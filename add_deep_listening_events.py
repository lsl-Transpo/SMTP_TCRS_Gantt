import csv
import json
import os
import re

events_csv_path = r"TO2 - Task 4. Engagement\SMPT EnviroIssues Scheudle\2026 SMTP Outreach Events and Conversations Tracker(Possible events).csv"
json_path = r"Task 6.1 - Management & Administration\schedule_data.json"

print("Loading existing tasks...")
with open(json_path, "r", encoding="utf-8") as f:
    tasks = json.load(f)

# Find Deep Listening task
deep_listening = None
for t in tasks:
    if t.get("Task ID") == "86bbh6vqk" or t.get("Task Name") == "Deep Listening":
        deep_listening = t
        break

if not deep_listening:
    raise ValueError("Could not find Deep Listening task!")

print(f"Found parent: {deep_listening['Task Name']} (ID: {deep_listening['Task ID']}, WBS: {deep_listening['WBS']})")

# Read events CSV
events = []
with open(events_csv_path, "r", encoding="utf-8-sig") as f:
    reader = csv.DictReader(f)
    for row in reader:
        events.append(row)

print(f"Read {len(events)} event rows from CSV.")

def clean_val(v):
    return str(v).strip() if v else ""

new_subtasks = []
wbs_base = deep_listening.get("WBS", "5.3")

for idx, ev in enumerate(events, start=1):
    region = clean_val(ev.get("REGION"))
    event_name = clean_val(ev.get("EVENT"))
    cbo = clean_val(ev.get("CBO"))
    activity = clean_val(ev.get("ACTIVITY"))
    dates_raw = clean_val(ev.get("DATES"))
    location = clean_val(ev.get("Location"))
    staffing = clean_val(ev.get("STAFFING"))
    roles = clean_val(ev.get("ROLES"))
    budget = clean_val(ev.get("ESTIMATED BUDGET"))
    reg_proc = clean_val(ev.get("REGISTRATION PROCESS"))
    lead_time = clean_val(ev.get("LEAD TIME"))

    # Determine Title
    # Format: [Region] Type: Org/Event
    title_parts = []
    if region:
        prefix = f"[{region}]"
    else:
        prefix = "[SMTP Outreach]"

    if activity and activity.lower() == "listening session":
        core = f"Listening Session: {cbo or event_name}"
    elif activity and "health fair" in activity.lower():
        core = f"Health Fair: {event_name or cbo}"
    elif activity and "farmer" in activity.lower():
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

    task_name = f"{prefix} {core}"

    # Determine Dates
    start_date = ""
    due_date = ""
    status = "To Do"

    if "october 17, 2026" in dates_raw.lower():
        start_date = "2026-10-17"
        due_date = "2026-10-17"
        status = "Complete"
    elif "october 16" in dates_raw.lower() and "october 18" in dates_raw.lower():
        start_date = "2026-10-16"
        due_date = "2026-10-18"
        status = "Complete"
    elif "october 26, 2026" in dates_raw.lower():
        start_date = "2026-10-26"
        due_date = "2026-10-26"
        status = "Complete"
    elif "november 2, 2026" in dates_raw.lower():
        start_date = "2026-11-02"
        due_date = "2026-11-02"
        status = "Complete"
    elif "17-nov" in dates_raw.lower():
        start_date = "2026-11-17"
        due_date = "2026-11-17"
        status = "To Do"
    elif "december 5" in dates_raw.lower() or "december 12" in dates_raw.lower():
        start_date = "2026-12-05"
        due_date = "2026-12-12"
        status = "To Do"
    elif "december 9, 2026" in dates_raw.lower():
        start_date = "2026-12-09"
        due_date = "2026-12-09"
        status = "To Do"
    else:
        # TBD or unspecified dates
        start_date = ""
        due_date = ""
        status = "To Do"

    # Build Description Notes
    notes_parts = []
    if region:
        notes_parts.append(f"Region: {region}")
    if cbo:
        notes_parts.append(f"CBO / Partner: {cbo}")
    if activity:
        notes_parts.append(f"Activity: {activity}")
    if location:
        notes_parts.append(f"Location: {location}")
    if dates_raw:
        notes_parts.append(f"Original Schedule Note: {dates_raw}")
    if staffing:
        notes_parts.append(f"Staffing: {staffing}")
    if roles:
        notes_parts.append(f"Roles: {roles}")
    if lead_time:
        notes_parts.append(f"Lead Time: {lead_time}")
    if reg_proc:
        notes_parts.append(f"Registration Process: {reg_proc}")
    if budget:
        notes_parts.append(f"Estimated Budget: {budget}")

    description = " | ".join(notes_parts)

    tid = f"dl_event_{idx:02d}"
    wbs = f"{wbs_base}.{idx}"

    task_obj = {
        "Project": "SMTP",
        "WBS": wbs,
        "Level": 3,
        "Task Name": task_name,
        "Task Type": "Task",
        "Status": status,
        "Start Date": start_date,
        "Due Date": due_date,
        "Assignees": staffing,
        "Priority": "Normal",
        "Tags": "Deep Listening Outreach",
        "Parent Task": "Deep Listening",
        "Task ID": tid,
        "Parent ID": deep_listening["Task ID"],
        "Description": description,
        "Task Link": ""
    }
    new_subtasks.append(task_obj)

print(f"Generated {len(new_subtasks)} subtasks.")

# Update Deep Listening dates to span the events (2026-10-01 to 2026-12-15)
deep_listening["Start Date"] = "2026-10-01"
deep_listening["Due Date"] = "2026-12-15"

# Find index of Deep Listening in tasks array and insert right after it (or after existing children)
dl_index = -1
for i, t in enumerate(tasks):
    if t.get("Task ID") == deep_listening["Task ID"]:
        dl_index = i
        break

# Insert the new subtasks right after Deep Listening
updated_tasks = tasks[:dl_index + 1] + new_subtasks + tasks[dl_index + 1:]

print(f"Total tasks updated from {len(tasks)} to {len(updated_tasks)}.")

# Save updated JSON
with open(json_path, "w", encoding="utf-8") as f:
    json.dump(updated_tasks, f, indent=2)

print("Saved schedule_data.json successfully.")

