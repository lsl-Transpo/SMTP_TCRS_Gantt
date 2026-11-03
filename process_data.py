import os
import re
import json
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

csv_path = r"Task 6.1 - Management & Administration\SMTP_TCRS_Gantt_Chart_20261103.csv"
excel_path = r"Task 6.1 - Management & Administration\SMTP_TCRS_Schedule.xlsx"

print(f"Reading CSV from {csv_path}...")
df = pd.read_csv(csv_path)

def clean_text(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    return s

def clean_assignees(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.startswith("[") and s.endswith("]"):
        s = s[1:-1].strip()
    return s

def clean_tags(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.startswith("[") and s.endswith("]"):
        s = s[1:-1].strip()
    # Format 'task order 1' -> 'Task Order 1'
    tags = [t.strip().title() for t in s.split(",") if t.strip()]
    return ", ".join(tags)

def format_date(epoch_val, text_val):
    if pd.notna(epoch_val):
        try:
            # epoch ms
            ts = float(epoch_val) / 1000.0
            dt = datetime.datetime.fromtimestamp(ts)
            return dt.strftime("%Y-%m-%d")
        except Exception:
            pass
    if pd.notna(text_val) and str(text_val).strip():
        # Text format like '8/14/2026, 4:00:00 AM PDT'
        m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", str(text_val).strip())
        if m:
            mth, day, yr = m.groups()
            return f"{yr}-{int(mth):02d}-{int(day):02d}"
    return ""

def clean_status(val):
    if pd.isna(val):
        return "To Do"
    s = str(val).strip().lower()
    if s == "complete":
        return "Complete"
    elif s in ("in progress", "in_progress", "inprogress"):
        return "In Progress"
    else:
        return "To Do"

def clean_type(val):
    if pd.isna(val):
        return "Task"
    s = str(val).strip().lower()
    if s == "milestone":
        return "Milestone"
    elif s == "deliverable":
        return "Deliverable"
    elif s in ("decision point", "decisionpoint", "decision"):
        return "Decision Point"
    return "Task"

def clean_priority(val):
    if pd.isna(val):
        return "Normal"
    try:
        p = int(float(val))
        if p == 1:
            return "Urgent"
        elif p == 2:
            return "High"
        elif p == 3:
            return "Normal"
        elif p == 4:
            return "Low"
    except Exception:
        pass
    return "Normal"

def clean_project(row):
    list_name = str(row.get('List Name', ''))
    folder = str(row.get('Folder Name/Path', ''))
    if "SMTP" in list_name or "SMTP" in folder:
        return "SMTP"
    elif "TCRS" in list_name or "TCRS" in folder:
        return "TCRS"
    return "SMTP"

# Process rows
rows = []
task_id_to_name = {}

for idx, r in df.iterrows():
    tid = clean_text(r['Task ID'])
    tname = clean_text(r['Task Name'])
    task_id_to_name[tid] = tname

for idx, r in df.iterrows():
    tid = clean_text(r['Task ID'])
    pid = clean_text(r['Parent ID'])
    pname = task_id_to_name.get(pid, "")
    
    project = clean_project(r)
    ttype = clean_type(r['Task Type'])
    status = clean_status(r['Status'])
    s_date = format_date(r['Start Date'], r['Start Date Text'])
    d_date = format_date(r['Due Date'], r['Due Date Text'])
    assignees = clean_assignees(r['Assignees'])
    tags = clean_tags(r['Tags'])
    priority = clean_priority(r['Priority'])
    content = clean_text(r['Task Content'])
    
    rows.append({
        'Project': project,
        'Task ID': tid,
        'Parent ID': pid,
        'Parent Task': pname,
        'Task Name': clean_text(r['Task Name']),
        'Task Type': ttype,
        'Status': status,
        'Start Date': s_date,
        'Due Date': d_date,
        'Assignees': assignees,
        'Priority': priority,
        'Tags': tags,
        'Description': content,
        'Task Link': clean_text(r.get('Task Link', ''))
    })

print(f"Processed {len(rows)} tasks.")

# Build hierarchical tree and assign WBS numbers
# Separate by Project
projects = ['SMTP', 'TCRS']
all_ordered_rows = []

for proj in projects:
    proj_rows = [r for r in rows if r['Project'] == proj]
    id_map = {r['Task ID']: r for r in proj_rows}
    children_map = {}
    roots = []
    
    for r in proj_rows:
        pid = r['Parent ID']
        if pid and pid in id_map:
            children_map.setdefault(pid, []).append(r)
        else:
            roots.append(r)
            
    # Function to recursively traverse
    wbs_counter = 1
    def traverse(task, wbs_prefix):
        task['WBS'] = wbs_prefix
        task['Level'] = len(wbs_prefix.split('.'))
        all_ordered_rows.append(task)
        cid_list = children_map.get(task['Task ID'], [])
        for c_idx, child in enumerate(cid_list, start=1):
            traverse(child, f"{wbs_prefix}.{c_idx}")
            
    for root in roots:
        traverse(root, f"{wbs_counter}")
        wbs_counter += 1

print(f"Hierarchically ordered {len(all_ordered_rows)} tasks.")
with open(r"Task 6.1 - Management & Administration\schedule_data.json", "w", encoding="utf-8") as f:
    json.dump(all_ordered_rows, f, indent=2)

print("Saved schedule_data.json successfully.")

