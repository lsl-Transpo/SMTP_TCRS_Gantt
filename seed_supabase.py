"""
seed_supabase.py
────────────────
One-time script: reads schedule_data.json and upserts all tasks into the
Supabase gantt_tasks table.  Run after executing migrate_gantt.sql.

Usage:
    python seed_supabase.py

Options (env vars):
    SEED_JSON_PATH   path to schedule_data.json (default: ./schedule_data.json)
    BATCH_SIZE       rows per INSERT batch      (default: 100)
"""

import json
import os
import sys

from dotenv import load_dotenv
from db import upsert_tasks, fetch_all_tasks, TABLE

load_dotenv()

JSON_PATH  = os.getenv("SEED_JSON_PATH", "schedule_data.json")
BATCH_SIZE = int(os.getenv("BATCH_SIZE", "100"))


def main() -> None:
    # ── Load JSON ──────────────────────────────────────────────────────────
    if not os.path.exists(JSON_PATH):
        print(f"ERROR: {JSON_PATH} not found.", file=sys.stderr)
        sys.exit(1)

    with open(JSON_PATH, "r", encoding="utf-8") as f:
        tasks = json.load(f)

    print(f"Loaded {len(tasks)} tasks from {JSON_PATH}")

    # ── Upsert in batches ──────────────────────────────────────────────────
    total_upserted = 0
    for i in range(0, len(tasks), BATCH_SIZE):
        batch = tasks[i : i + BATCH_SIZE]
        n = upsert_tasks(batch)
        total_upserted += n
        print(f"  Upserted rows {i + 1}–{i + len(batch)} ({n} rows)")

    # ── Verify ──────────────────────────────────────────────────────────────
    all_in_db = fetch_all_tasks()
    print(f"\nSeed complete: {total_upserted} rows upserted.")
    print(f"  {TABLE} now contains {len(all_in_db)} rows in Supabase.")

    # Quick sanity check: do we have both projects?
    projects = {t["Project"] for t in all_in_db}
    print(f"  Projects found: {', '.join(sorted(projects))}")


if __name__ == "__main__":
    main()
