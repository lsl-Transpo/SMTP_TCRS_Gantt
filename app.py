"""
app.py — SMTP/TCRS Gantt Flask API backed by Supabase Postgres.

Routes
──────
GET  /                      → HTML status page
GET  /api/schedule          → all tasks as JSON  (Gantt-compatible)
GET  /api/schedule?project= → filter by SMTP or TCRS
POST /api/schedule          → upsert a list of tasks
GET  /api/status            → health-check / row count
GET  /api/tasks/<task_id>   → single task
PUT  /api/tasks/<task_id>   → update single task
DELETE /api/tasks/<task_id> → delete single task
"""

import json
import os

from flask import Flask, jsonify, request
from flask_cors import CORS
from dotenv import load_dotenv
from postgrest.exceptions import APIError

from db import (
    fetch_all_tasks,
    fetch_tasks_by_project,
    upsert_tasks,
    delete_task,
    get_connection,
    TABLE,
    TASK_FIELDS,
)

load_dotenv()

app = Flask(__name__)
CORS(app)


# ── Helper ──────────────────────────────────────────────────────────────────

def _error(msg: str, code: int = 400):
    return jsonify({"status": "error", "message": msg}), code


# ── Routes ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    """Simple health / info page."""
    try:
        conn = get_connection()
        with conn.cursor() as cur:
            cur.execute(f"SELECT COUNT(*) FROM {TABLE};")
            (count,) = cur.fetchone()
        conn.close()
        db_status = f"✅ Connected — {count} tasks in database"
    except Exception as exc:
        db_status = f"❌ Database error: {exc}"

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>SMTP/TCRS Gantt API</title>
  <style>
    body {{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
          max-width:720px;margin:48px auto;padding:0 20px;color:#1e293b;line-height:1.6}}
    h1  {{color:#1e3a8a;border-bottom:3px solid #0284c7;padding-bottom:8px}}
    h2  {{color:#0369a1;margin-top:32px}}
    code{{background:#f1f5f9;padding:2px 6px;border-radius:4px;font-size:.9em}}
    pre {{background:#0f172a;color:#e2e8f0;padding:16px;border-radius:8px;overflow-x:auto}}
    .ok {{color:#059669;font-weight:bold}}
    .er {{color:#dc2626;font-weight:bold}}
    table{{border-collapse:collapse;width:100%}}
    td,th{{border:1px solid #e2e8f0;padding:8px 12px;text-align:left}}
    th{{background:#1e3a8a;color:#fff}}
    tr:nth-child(even){{background:#f8fafc}}
  </style>
</head>
<body>
  <h1>🗂️ WSDOT SMTP &amp; TCRS Gantt API</h1>
  <p class="{'ok' if '✅' in db_status else 'er'}">{db_status}</p>

  <h2>API Endpoints</h2>
  <table>
    <tr><th>Method</th><th>Path</th><th>Description</th></tr>
    <tr><td>GET</td><td><code>/api/schedule</code></td><td>All tasks (Gantt JSON)</td></tr>
    <tr><td>GET</td><td><code>/api/schedule?project=SMTP</code></td><td>Filter by project</td></tr>
    <tr><td>POST</td><td><code>/api/schedule</code></td><td>Upsert tasks — body: <code>{{"tasks":[…]}}</code></td></tr>
    <tr><td>GET</td><td><code>/api/status</code></td><td>Health check &amp; row count</td></tr>
    <tr><td>GET</td><td><code>/api/tasks/&lt;task_id&gt;</code></td><td>Single task by ID</td></tr>
    <tr><td>PUT</td><td><code>/api/tasks/&lt;task_id&gt;</code></td><td>Update a task</td></tr>
    <tr><td>DELETE</td><td><code>/api/tasks/&lt;task_id&gt;</code></td><td>Delete a task</td></tr>
  </table>

  <h2>Task Fields</h2>
  <pre>{json.dumps(TASK_FIELDS, indent=2)}</pre>
</body>
</html>"""
    return html


@app.route("/api/status")
def api_status():
    """Health check — returns row count and mtime of local schedule_data.json."""
    try:
        conn = get_connection()
        with conn.cursor() as cur:
            cur.execute(f"SELECT COUNT(*) FROM {TABLE};")
            (count,) = cur.fetchone()
            cur.execute(
                f"SELECT MAX(updated_at) FROM {TABLE};"
            )
            (last_updated,) = cur.fetchone()
        conn.close()

        json_mtime = 0.0
        if os.path.exists("schedule_data.json"):
            json_mtime = os.path.getmtime("schedule_data.json")

        return jsonify({
            "status": "ok",
            "task_count": count,
            "last_updated": str(last_updated) if last_updated else None,
            "json_mtime": json_mtime,
            "port": int(os.getenv("PORT", 5000)),
        })
    except Exception as exc:
        return _error(str(exc), 500)


@app.route("/api/schedule", methods=["GET"])
def api_get_schedule():
    """Return all tasks ordered by project + WBS."""
    try:
        project = request.args.get("project", "").strip().upper()
        if project in ("SMTP", "TCRS"):
            tasks = fetch_tasks_by_project(project)
        else:
            tasks = fetch_all_tasks()

        # Also return mtime for the Gantt HTML sync logic
        mtime = 0.0
        if os.path.exists("schedule_data.json"):
            mtime = os.path.getmtime("schedule_data.json")

        return jsonify({"tasks": tasks, "mtime": mtime, "count": len(tasks)})
    except Exception as exc:
        return _error(str(exc), 500)


@app.route("/api/schedule", methods=["POST"])
def api_post_schedule():
    """
    Upsert tasks from the Gantt HTML editor.
    Body: { "tasks": [ { <task fields> }, … ] }
    """
    try:
        body = request.get_json(force=True)
        if not body or "tasks" not in body:
            return _error('Request body must be JSON with a "tasks" key.')

        tasks = body["tasks"]
        if not isinstance(tasks, list):
            return _error('"tasks" must be a list.')

        n = upsert_tasks(tasks)

        # Keep local schedule_data.json in sync as backup
        try:
            all_tasks = fetch_all_tasks()
            with open("schedule_data.json", "w", encoding="utf-8") as f:
                import json as _json
                _json.dump(all_tasks, f, indent=2, ensure_ascii=False)
        except Exception:
            pass  # Non-fatal — Supabase is the source of truth

        mtime = 0.0
        if os.path.exists("schedule_data.json"):
            mtime = os.path.getmtime("schedule_data.json")

        return jsonify({
            "status": "success",
            "upserted": n,
            "mtime": mtime,
        })
    except Exception as exc:
        return _error(str(exc), 500)


@app.route("/api/tasks/<task_id>", methods=["GET"])
def api_get_task(task_id: str):
    """Fetch a single task by its task_id."""
    try:
        conn = get_connection()
        import psycopg2.extras
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                f"SELECT * FROM {TABLE} WHERE task_id = %s;", (task_id,)
            )
            row = cur.fetchone()
        conn.close()
        if row is None:
            return _error(f"Task '{task_id}' not found.", 404)
        from db import row_to_task
        return jsonify({"status": "success", "data": row_to_task(dict(row))})
    except Exception as exc:
        return _error(str(exc), 500)


@app.route("/api/tasks/<task_id>", methods=["PUT"])
def api_put_task(task_id: str):
    """Update a single task. Body is a partial or full task JSON object."""
    try:
        body = request.get_json(force=True)
        if not body:
            return _error("Request body must be a JSON task object.")
        # Inject the task_id from the URL so upsert works correctly
        body["Task ID"] = task_id
        n = upsert_tasks([body])
        return jsonify({"status": "success", "upserted": n})
    except Exception as exc:
        return _error(str(exc), 500)


@app.route("/api/tasks/<task_id>", methods=["DELETE"])
def api_delete_task(task_id: str):
    """Delete a single task by task_id."""
    try:
        deleted = delete_task(task_id)
        if not deleted:
            return _error(f"Task '{task_id}' not found.", 404)
        return jsonify({"status": "success", "deleted": task_id})
    except Exception as exc:
        return _error(str(exc), 500)


# ── Entry point ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    debug = os.getenv("FLASK_DEBUG", "1") == "1"
    print(f"Starting SMTP/TCRS Gantt API on http://localhost:{port}")
    app.run(debug=debug, port=port)
