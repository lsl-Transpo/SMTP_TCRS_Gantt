"""
main.py — Supabase Postgres connection test for SMTP_TCRS_Gantt.

Run this to verify your DATABASE_URL is working and to show a summary
of what's currently stored in the gantt_tasks table.

Usage:
    python main.py
"""

from dotenv import load_dotenv
from db import get_connection, fetch_all_tasks, TABLE

load_dotenv()


def main():
    # ── 1. Test raw psycopg2 connection ───────────────────────────────────
    print("Testing Supabase Postgres connection…", end=" ", flush=True)
    conn = get_connection()
    with conn.cursor() as cur:
        cur.execute("SELECT version();")
        version = cur.fetchone()[0]
    conn.close()
    print("OK")
    print(f"  Server: {version[:60]}")

    # ── 2. Check gantt_tasks table ────────────────────────────────────────
    print(f"\nQuerying {TABLE}…", end=" ", flush=True)
    try:
        tasks = fetch_all_tasks()
        print(f"OK  ({len(tasks)} rows)")

        if tasks:
            projects = {}
            for t in tasks:
                p = t.get("Project", "?")
                projects[p] = projects.get(p, 0) + 1

            print("\n  Project breakdown:")
            for proj, count in sorted(projects.items()):
                print(f"    {proj}: {count} tasks")

            statuses = {}
            for t in tasks:
                s = t.get("Status", "?")
                statuses[s] = statuses.get(s, 0) + 1

            print("\n  Status breakdown:")
            for status, count in sorted(statuses.items()):
                print(f"    {status}: {count}")

            print(f"\n  Sample task (first row):")
            first = tasks[0]
            for key in ["Task ID", "Task Name", "WBS", "Project", "Status", "Start Date", "Due Date"]:
                print(f"    {key}: {first.get(key, '')}")
        else:
            print("\n  ! Table exists but has no rows yet.")
            print("  Run: python seed_supabase.py  to load schedule_data.json into Supabase.")

    except Exception as exc:
        if "PGRST205" in str(exc) or "gantt_tasks" in str(exc) or "does not exist" in str(exc).lower():
            print("!  Table not found.")
            print("  → Run the SQL in migrate_gantt.sql in your Supabase SQL Editor, then:")
            print("    python seed_supabase.py")
        else:
            print(f"✗  Error: {exc}")


if __name__ == "__main__":
    main()
