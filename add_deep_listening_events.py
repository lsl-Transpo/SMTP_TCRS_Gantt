"""
add_deep_listening_events.py
────────────────────────────
Reads a CSV of SMTP Outreach events and upserts them as subtasks of the
"Deep Listening" parent task into Supabase Postgres, then refreshes
schedule_data.json and regenerates the Excel workbook.

Usage:
    python add_deep_listening_events.py [path/to/events.csv]

If no CSV path is given, the script looks for the default path set in
EVENTS_CSV_PATH below (or the env var of the same name).
"""

import csv
import json
import os
import subprocess
import sys

from dotenv import load_dotenv
from db import fetch_all_tasks, upsert_tasks

load_dotenv()

# ── Config ───────────────────────────────────────────────────────────────────
EVENTS_CSV_PATH = os.getenv(
    "EVENTS_CSV_PATH",
    r"TO2 - Task 4. Engagement\SMPT EnviroIssues Scheudle"
    r"\2026 SMTP Outreach Events and Conversations Tracker(Possible events).csv",
)

DEEP_LISTENING_TASK_ID   = "86bbh6vqk"
DEEP_LISTENING_TASK_NAME = "Deep Listening"
DEEP_LISTENING_START     = "2026-10-01"
DEEP_LISTENING_END       = "2026-12-15"


# ── Helpers ───────────────────────────────────────────────────────────────────

def clean_val(v) -> str:
    return str(v).strip() if v else ""


def _parse_dates(dates_raw: str) -> tuple[str, str, str]:
    """Return (start_date, due_date, status) from the raw DATES field."""
    dr = dates_raw.lower()
    if "october 17, 2026" in dr:
        return "2026-10-17", "2026-10-17", "Complete"
    if "october 16" in dr and "october 18" in dr:
        return "2026-10-16", "2026-10-18", "Complete"
    if "october 26, 2026" in dr:
        return "2026-10-26", "2026-10-26", "Complete"
    if "november 2, 2026" in dr:
        return "2026-11-02", "2026-11-02", "Complete"
    if "17-nov" in dr:
        return "2026-11-17", "2026-11-17", "To Do"
    if "december 5" in dr or "december 12" in dr:
        return "2026-12-05", "2026-12-12", "To Do"
    if "december 9, 2026" in dr:
        return "2026-12-09", "2026-12-09", "To Do"
    return "", "", "To Do"


def _build_task_name(region: str, activity: str, event_name: str, cbo: str, idx: int) -> str:
    prefix = f"[{region}]" if region else "[SMTP Outreach]"
    act_l  = activity.lower() if activity else ""
    if act_l == "listening session":
        core = f"Listening Session: {cbo or event_name}"
    elif "health fair" in act_l:
        core = f"Health Fair: {event_name or cbo}"
    elif "farmer" in act_l:
        core = f"Outreach: {event_name or cbo}"
    elif activity:
        core = f"{activity}: {cbo or event_name}"
    elif event_name and cbo:
        core = f"{event_name} - {cbo}"
    elif cbo:
        core = f"Outreach: {cbo}"
    elif event_name:
        core = f"Outreach: {event_name}"
    else:
        core = f"Outreach Event {idx}"
    return f"{prefix} {core}"


def _build_description(region, cbo, activity, location, dates_raw,
                        staffing, roles, lead_time, reg_proc, budget) -> str:
    parts = []
    if region:    parts.append(f"Region: {region}")
    if cbo:       parts.append(f"CBO / Partner: {cbo}")
    if activity:  parts.append(f"Activity: {activity}")
    if location:  parts.append(f"Location: {location}")
    if dates_raw: parts.append(f"Original Schedule Note: {dates_raw}")
    if staffing:  parts.append(f"Staffing: {staffing}")
    if roles:     parts.append(f"Roles: {roles}")
    if lead_time: parts.append(f"Lead Time: {lead_time}")
    if reg_proc:  parts.append(f"Registration Process: {reg_proc}")
    if budget:    parts.append(f"Estimated Budget: {budget}")
    return " | ".join(parts)


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    csv_path = sys.argv[1] if len(sys.argv) > 1 else EVENTS_CSV_PATH

    # ── 1. Load events CSV ────────────────────────────────────────────────
    if not os.path.exists(csv_path):
        print(f"ERROR: CSV not found: {csv_path}", file=sys.stderr)
        sys.exit(1)

    events = []
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            events.append(row)
    print(f"Read {len(events)} event rows from CSV.")

    # ── 2. Fetch all tasks from Supabase ──────────────────────────────────
    print("Fetching tasks from Supabase…")
    tasks = fetch_all_tasks()
    print(f"  Loaded {len(tasks)} tasks from database.")

    # ── 3. Find Deep Listening parent ─────────────────────────────────────
    deep_listening = next(
        (t for t in tasks
         if t.get("Task ID") == DEEP_LISTENING_TASK_ID
         or t.get("Task Name") == DEEP_LISTENING_TASK_NAME),
        None,
    )
    if not deep_listening:
        print(f"ERROR: Could not find '{DEEP_LISTENING_TASK_NAME}' task!", file=sys.stderr)
        sys.exit(1)

    print(
        f"Found parent: {deep_listening['Task Name']} "
        f"(ID: {deep_listening['Task ID']}, WBS: {deep_listening['WBS']})"
    )
    wbs_base = deep_listening.get("WBS", "5.3")

    # ── 4. Build subtasks ─────────────────────────────────────────────────
    new_subtasks = []
    for idx, ev in enumerate(events, start=1):
        region    = clean_val(ev.get("REGION"))
        event_name = clean_val(ev.get("EVENT"))
        cbo        = clean_val(ev.get("CBO"))
        activity   = clean_val(ev.get("ACTIVITY"))
        dates_raw  = clean_val(ev.get("DATES"))
        location   = clean_val(ev.get("Location"))
        staffing   = clean_val(ev.get("STAFFING"))
        roles      = clean_val(ev.get("ROLES"))
        budget     = clean_val(ev.get("ESTIMATED BUDGET"))
        reg_proc   = clean_val(ev.get("REGISTRATION PROCESS"))
        lead_time  = clean_val(ev.get("LEAD TIME"))

        start_date, due_date, status = _parse_dates(dates_raw)
        task_name   = _build_task_name(region, activity, event_name, cbo, idx)
        description = _build_description(
            region, cbo, activity, location, dates_raw,
            staffing, roles, lead_time, reg_proc, budget,
        )

        new_subtasks.append({
            "Project":     "SMTP",
            "Task ID":     f"dl_event_{idx:02d}",
            "Parent ID":   deep_listening["Task ID"],
            "Parent Task": DEEP_LISTENING_TASK_NAME,
            "Task Name":   task_name,
            "Task Type":   "Task",
            "Status":      status,
            "Start Date":  start_date,
            "Due Date":    due_date,
            "Assignees":   staffing,
            "Priority":    "Normal",
            "Tags":        "Deep Listening Outreach",
            "Description": description,
            "Task Link":   "",
            "WBS":         f"{wbs_base}.{idx}",
            "Level":       3,
        })

    print(f"Generated {len(new_subtasks)} subtasks.")

    # ── 5. Update Deep Listening parent dates ─────────────────────────────
    deep_listening["Start Date"] = DEEP_LISTENING_START
    deep_listening["Due Date"]   = DEEP_LISTENING_END

    # ── 6. Upsert parent + all subtasks to Supabase ───────────────────────
    print("Upserting to Supabase…")
    to_upsert = [deep_listening] + new_subtasks
    n = upsert_tasks(to_upsert)
    print(f"  Upserted {n} rows.")

    # ── 7. Refresh schedule_data.json ─────────────────────────────────────
    print("Refreshing schedule_data.json from Supabase…")
    all_tasks = fetch_all_tasks()
    json_path = "schedule_data.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(all_tasks, f, indent=2, ensure_ascii=False)
    print(f"  Wrote {len(all_tasks)} tasks to {json_path}.")

    # ── 8. Regenerate Excel ───────────────────────────────────────────────
    script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "create_excel_schedule.py")
    if os.path.exists(script):
        print("Regenerating Excel workbook…")
        try:
            subprocess.run([sys.executable, script], check=True, timeout=120)
            print("  Excel updated successfully.")
        except Exception as exc:
            print(f"  Warning: Excel regeneration failed: {exc}")
    else:
        print("  (create_excel_schedule.py not found — skipping Excel update)")

    print("\n✓ Done! Deep Listening subtasks are live in Supabase.")


if __name__ == "__main__":
    main()
