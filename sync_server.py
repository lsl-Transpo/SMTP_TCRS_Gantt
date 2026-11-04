"""
sync_server.py — SMTP/TCRS Gantt live-sync server (Supabase Postgres backend).

Serves:
  GET  /                 → SMTP_TCRS_Interactive_Gantt.html
  GET  /api/status       → health check (DB row count + Excel mtime)
  GET  /api/schedule     → all tasks from Supabase (Gantt JSON format)
  POST /api/schedule     → upsert tasks to Supabase AND write schedule_data.json
                           + regenerates SMTP_TCRS_Schedule.xlsx

The Gantt HTML uses these same endpoints — no code changes are needed in the
HTML as long as it was talking to this server already.

Usage:
    python sync_server.py
    (or double-click Start_Gantt_Chart.bat)
"""

import json
import os
import subprocess
import sys
import time
import urllib.parse
import webbrowser
from http.server import HTTPServer, SimpleHTTPRequestHandler

from db import (
    fetch_all_tasks,
    upsert_tasks,
    get_connection,
    TABLE,
)

EXCEL_FILE = "SMTP_TCRS_Schedule.xlsx"
HTML_FILE  = "SMTP_TCRS_Interactive_Gantt.html"
PORT       = 8080


# ── Helpers ──────────────────────────────────────────────────────────────────

def _json_response(handler, data: dict, status: int = 200) -> None:
    body = json.dumps(data, default=str).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json")
    handler.end_headers()
    handler.wfile.write(body)


def read_db_tasks() -> list[dict]:
    """Return all tasks from Supabase, ordered by project + WBS."""
    return fetch_all_tasks(order_by_wbs=True)


def save_db_tasks(tasks: list[dict]) -> int:
    """
    Upsert tasks to Supabase.
    Also writes schedule_data.json as a local backup,
    then regenerates SMTP_TCRS_Schedule.xlsx.
    Returns the count of upserted rows.
    """
    n = upsert_tasks(tasks)

    # Keep local JSON in sync as a backup / offline fallback
    _write_json_backup(tasks)

    # Regenerate Excel
    _regenerate_excel()

    return n


def _write_json_backup(tasks: list[dict]) -> None:
    """Overwrite schedule_data.json with the full current task list from DB."""
    try:
        with open("schedule_data.json", "w", encoding="utf-8") as f:
            json.dump(tasks, f, indent=2, ensure_ascii=False)
    except Exception as exc:
        print(f"[warn] Could not write schedule_data.json: {exc}")


def _regenerate_excel() -> None:
    """Run create_excel_schedule.py to regenerate the .xlsx file."""
    script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "create_excel_schedule.py")
    if os.path.exists(script):
        try:
            subprocess.run([sys.executable, script], check=True, timeout=120)
        except Exception as exc:
            print(f"[warn] Excel regeneration failed: {exc}")
    else:
        print("[warn] create_excel_schedule.py not found — skipping Excel update")


def _db_row_count() -> int:
    try:
        conn = get_connection()
        with conn.cursor() as cur:
            cur.execute(f"SELECT COUNT(*) FROM {TABLE};")
            (n,) = cur.fetchone()
        conn.close()
        return n
    except Exception:
        return -1


# ── HTTP Handler ─────────────────────────────────────────────────────────────

class GanttSyncHandler(SimpleHTTPRequestHandler):
    # ── CORS headers on every response ────────────────────────────────────
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin",  "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def log_message(self, fmt, *args):
        # Suppress noisy static-file log entries; keep API entries
        path = args[0] if args else ""
        if "/api/" in str(path):
            super().log_message(fmt, *args)

    # ── GET ───────────────────────────────────────────────────────────────
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path   = parsed.path

        if path == "/api/status":
            count = _db_row_count()
            mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
            _json_response(self, {
                "status": "ok",
                "task_count": count,
                "excel_mtime": mtime,
                "port": PORT,
                "backend": "supabase-postgres",
            })

        elif path == "/api/schedule":
            try:
                tasks = read_db_tasks()
                mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
                _json_response(self, {"tasks": tasks, "mtime": mtime, "count": len(tasks)})
            except Exception as exc:
                _json_response(self, {"error": str(exc)}, 500)

        elif path in ("/", ""):
            self.path = "/" + HTML_FILE
            super().do_GET()

        else:
            super().do_GET()

    # ── POST ──────────────────────────────────────────────────────────────
    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path   = parsed.path

        if path == "/api/schedule":
            try:
                length   = int(self.headers.get("Content-Length", 0))
                raw_body = self.rfile.read(length)
                data     = json.loads(raw_body.decode("utf-8"))
                tasks    = data.get("tasks", [])

                if not tasks:
                    _json_response(self, {"error": "No tasks in request body"}, 400)
                    return

                n = save_db_tasks(tasks)
                mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
                _json_response(self, {
                    "success": True,
                    "upserted": n,
                    "count": len(tasks),
                    "mtime": mtime,
                })

            except Exception as exc:
                _json_response(self, {"error": str(exc)}, 500)
        else:
            self.send_response(404)
            self.end_headers()


# ── Server startup ───────────────────────────────────────────────────────────

def run_server():
    # Always run from the script's directory so static file serving works
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # Verify DB connectivity before opening the browser
    print("Checking Supabase connection…", end=" ", flush=True)
    count = _db_row_count()
    if count >= 0:
        print(f"✓  ({count} tasks in database)")
    else:
        print("⚠  Could not reach Supabase — check .env and network")

    httpd = HTTPServer(("", PORT), GanttSyncHandler)

    print("=" * 62)
    print(" WSDOT SMTP & TCRS Gantt Chart Live Sync Server Running")
    print(f" URL:          http://localhost:{PORT}")
    print(f" Backend:      Supabase Postgres  ({TABLE})")
    print(f" Excel target: {os.path.abspath(EXCEL_FILE)}")
    print("=" * 62)

    try:
        webbrowser.open(f"http://localhost:{PORT}")
    except Exception:
        pass

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server.")
        httpd.server_close()


if __name__ == "__main__":
    run_server()
