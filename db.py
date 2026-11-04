"""
db.py — Shared Supabase / psycopg2 database helpers for SMTP_TCRS_Gantt.

Provides:
  - get_connection()     : raw psycopg2 connection (for bulk ops / upserts)
  - get_supabase()       : supabase-py Client (for REST API calls)
  - TASK_FIELDS          : canonical ordered list of JSON/column names
  - row_to_task(row)     : convert a DB row dict back to the Gantt JSON shape
  - task_to_row(task)    : convert a Gantt JSON object to a DB-ready dict
"""

import os
import psycopg2
import psycopg2.extras
from dotenv import load_dotenv
from supabase import create_client, Client
from postgrest.exceptions import APIError  # re-export for callers

load_dotenv()

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
DATABASE_URL = os.getenv("DATABASE_URL")
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
TABLE = "gantt_tasks"

# Canonical field order — matches schedule_data.json exactly
TASK_FIELDS = [
    "Project",
    "Task ID",
    "Parent ID",
    "Parent Task",
    "Task Name",
    "Task Type",
    "Status",
    "Start Date",
    "Due Date",
    "Assignees",
    "Priority",
    "Tags",
    "Description",
    "Task Link",
    "WBS",
    "Level",
]

# Map from JSON key → Postgres column name (snake_case)
_JSON_TO_COL = {
    "Project":     "project",
    "Task ID":     "task_id",
    "Parent ID":   "parent_id",
    "Parent Task": "parent_task",
    "Task Name":   "task_name",
    "Task Type":   "task_type",
    "Status":      "status",
    "Start Date":  "start_date",
    "Due Date":    "due_date",
    "Assignees":   "assignees",
    "Priority":    "priority",
    "Tags":        "tags",
    "Description": "description",
    "Task Link":   "task_link",
    "WBS":         "wbs",
    "Level":       "level",
}

_COL_TO_JSON = {v: k for k, v in _JSON_TO_COL.items()}


# ---------------------------------------------------------------------------
# Connections
# ---------------------------------------------------------------------------
def get_connection():
    """Return a psycopg2 connection using DATABASE_URL from .env."""
    if not DATABASE_URL:
        raise ValueError(
            "DATABASE_URL is not set in .env — "
            "set it to postgresql://postgres.<project>:<password>@<host>:5432/postgres"
        )
    return psycopg2.connect(DATABASE_URL)


def get_supabase() -> Client:
    """Return a supabase-py Client using SUPABASE_URL/KEY from .env."""
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise ValueError(
            "SUPABASE_URL and SUPABASE_KEY must be set in .env"
        )
    return create_client(SUPABASE_URL, SUPABASE_KEY)


# ---------------------------------------------------------------------------
# Shape converters
# ---------------------------------------------------------------------------
def task_to_row(task: dict) -> dict:
    """Convert a Gantt JSON task object → Postgres column dict."""
    row = {}
    for json_key, col in _JSON_TO_COL.items():
        val = task.get(json_key, "")
        if col == "level":
            try:
                row[col] = int(val) if val != "" else 1
            except (TypeError, ValueError):
                row[col] = 1
        elif col in ("start_date", "due_date"):
            # Store blank dates as NULL
            row[col] = val if val else None
        else:
            row[col] = val if val is not None else ""
    return row


def row_to_task(row: dict) -> dict:
    """Convert a Postgres column dict → Gantt JSON task object."""
    task = {}
    for col, json_key in _COL_TO_JSON.items():
        val = row.get(col)
        if json_key == "Level":
            task[json_key] = int(val) if val is not None else 1
        elif json_key in ("Start Date", "Due Date"):
            # Return ISO string or empty string
            task[json_key] = str(val) if val else ""
        else:
            task[json_key] = val if val is not None else ""
    return task


# ---------------------------------------------------------------------------
# High-level DB helpers
# ---------------------------------------------------------------------------
def fetch_all_tasks(order_by_wbs: bool = True) -> list[dict]:
    """
    Fetch all rows from gantt_tasks and return them as Gantt JSON task dicts,
    ordered by WBS so the hierarchy is preserved.
    """
    conn = get_connection()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            order_clause = "ORDER BY project, wbs" if order_by_wbs else ""
            cur.execute(f"SELECT * FROM {TABLE} {order_clause};")
            rows = cur.fetchall()
        return [row_to_task(dict(r)) for r in rows]
    finally:
        conn.close()


def upsert_tasks(tasks: list[dict]) -> int:
    """
    Upsert a list of Gantt JSON task dicts into gantt_tasks.
    Uses task_id as the conflict key.
    Returns the number of rows affected.
    """
    if not tasks:
        return 0

    rows = [task_to_row(t) for t in tasks]
    cols = list(rows[0].keys())
    col_names = ", ".join(cols)
    placeholders = ", ".join([f"%({c})s" for c in cols])
    update_set = ", ".join(
        [f"{c} = EXCLUDED.{c}" for c in cols if c != "task_id"]
    )

    sql = f"""
        INSERT INTO {TABLE} ({col_names})
        VALUES ({placeholders})
        ON CONFLICT (task_id) DO UPDATE SET {update_set};
    """

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            psycopg2.extras.execute_batch(cur, sql, rows, page_size=200)
        conn.commit()
        return len(rows)
    finally:
        conn.close()


def delete_task(task_id: str) -> bool:
    """Delete a single task by task_id. Returns True if a row was deleted."""
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(f"DELETE FROM {TABLE} WHERE task_id = %s;", (task_id,))
            deleted = cur.rowcount > 0
        conn.commit()
        return deleted
    finally:
        conn.close()


def fetch_tasks_by_project(project: str) -> list[dict]:
    """Return all tasks for a given project (e.g. 'SMTP' or 'TCRS')."""
    conn = get_connection()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                f"SELECT * FROM {TABLE} WHERE project = %s ORDER BY wbs;",
                (project,),
            )
            rows = cur.fetchall()
        return [row_to_task(dict(r)) for r in rows]
    finally:
        conn.close()

