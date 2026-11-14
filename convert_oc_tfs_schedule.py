import json
import re
from datetime import date, datetime
from pathlib import Path

import openpyxl


BASE_DIR = Path(__file__).resolve().parent
SOURCE_PATH = BASE_DIR / "oc_tfs_schedule.xlsx"
OUTPUT_PATH = BASE_DIR / "schedule_data.json"
PROJECT_NAME = "OC TFS"


def clean_text(value):
    if value is None:
        return ""
    return " ".join(str(value).replace("\xa0", " ").replace("\ufffd", "–").split())


def date_text(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return ""


def status_text(value):
    value = clean_text(value).lower()
    if value == "complete":
        return "Complete"
    if value == "in progress":
        return "In Progress"
    return "To Do"


def make_task(task_id, parent, name, task_type, status, start_date, due_date, assignees, description, wbs, level):
    return {
        "Project": PROJECT_NAME,
        "Task ID": task_id,
        "Parent ID": parent["Task ID"] if parent else "",
        "Parent Task": parent["Task Name"] if parent else "",
        "Task Name": name,
        "Task Type": task_type,
        "Status": status,
        "Start Date": start_date,
        "Due Date": due_date,
        "Assignees": assignees,
        "Priority": "Normal",
        "Tags": "",
        "Description": description,
        "Task Link": "",
        "WBS": wbs,
        "Level": level,
    }


def convert():
    workbook = openpyxl.load_workbook(SOURCE_PATH, data_only=True, read_only=True)
    sheet = workbook.active
    headers = [clean_text(value).lower() for value in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]
    columns = {header: index for index, header in enumerate(headers) if header}
    required_headers = {
        "subtask / deliverable name",
        "status",
        "type",
        "start date",
        "due date",
        "end date",
        "drafter(s)",
        "reviewer",
        "final reviewer",
        "notes",
        "notes2",
    }
    missing_headers = required_headers - columns.keys()
    if missing_headers:
        raise ValueError(f"Workbook is missing required columns: {', '.join(sorted(missing_headers))}")
    draft_header = next(
        (header for header in columns if "draft due date" in header.lower()),
        None,
    )
    if draft_header is None:
        raise ValueError("Workbook is missing the draft due date column")

    tasks = []
    section = None
    section_number = 0
    child_number = 0
    task_number = 0
    for row in sheet.iter_rows(min_row=2, values_only=True):
        name = clean_text(row[columns["subtask / deliverable name"]])
        if not name:
            continue

        section_match = re.match(r"^Subtask\s+(\d+)\s*[–-]\s*(.+)$", name, re.IGNORECASE)
        if section_match:
            section_number = int(section_match.group(1))
            child_number = 0
            section = make_task(
                f"section-{section_number}", None, name, "Task", "To Do", "", "", "", "",
                str(section_number), 1,
            )
            tasks.append(section)
            continue
        if section is None:
            raise ValueError(f"Schedule item appears before its section: {name}")

        status = status_text(row[columns["status"]])
        task_type = clean_text(row[columns["type"]]).title()
        start_date = date_text(row[columns["start date"]])
        due_value = row[columns["due date"]]
        due_date = date_text(due_value)
        draft_date = date_text(row[columns[draft_header]])
        end_date = date_text(row[columns["end date"]])
        due_note = clean_text(due_value) if isinstance(due_value, str) else ""
        if task_type not in {"Task", "Deliverable", "Milestone"}:
            if task_type:
                raise ValueError(f"Unsupported type {task_type!r} for schedule item {name!r}")
            task_type = "Deliverable" if due_date or draft_date else "Task"

        assignees = clean_text(row[columns["drafter(s)"]])
        notes = []
        for label, header in (
            ("Reviewer", "reviewer"),
            ("Final reviewer", "final reviewer"),
            ("Notes", "notes"),
            ("Notes2", "notes2"),
        ):
            value = row[columns[header]]
            text = clean_text(value)
            if text:
                notes.append(f"{label}: {text}")
        if due_note:
            notes.append(f"Schedule: {due_note}")
        description = "; ".join(notes)

        if task_type == "Deliverable":
            checkpoints = []
            if draft_date and due_date:
                if draft_date != due_date:
                    checkpoints.append((f"{name} (Draft due)", draft_date))
                checkpoints.append((f"{name} (Due)", due_date))
            elif due_date:
                checkpoints.append((name, due_date))
            elif draft_date:
                checkpoints.append((f"{name} (Draft due)", draft_date))
            if not checkpoints:
                checkpoints.append((name, ""))
            for checkpoint_name, checkpoint_date in checkpoints:
                task_number += 1
                child_number += 1
                tasks.append(make_task(
                    f"item-{task_number}", section, checkpoint_name, "Deliverable", status, "",
                    checkpoint_date, assignees, description, f"{section_number}.{child_number}", 2,
                ))
        else:
            task_number += 1
            child_number += 1
            if task_type == "Milestone":
                start_date = ""
                due_date = due_date or draft_date or end_date or start_date
            else:
                due_date = end_date or due_date or start_date
            tasks.append(make_task(
                f"item-{task_number}", section, name, task_type, status, start_date,
                due_date, assignees, description, f"{section_number}.{child_number}", 2,
            ))

    for section in (task for task in tasks if task["Level"] == 1):
        children = [task for task in tasks if task["Parent ID"] == section["Task ID"]]
        child_statuses = {task["Status"] for task in children}
        if "In Progress" in child_statuses:
            section["Status"] = "In Progress"
        elif "To Do" in child_statuses:
            section["Status"] = "To Do"
        else:
            section["Status"] = "Complete"

        section_dates = [
            value
            for child in children
            for value in (child["Start Date"], child["Due Date"])
            if value
        ]
        if section_dates:
            section["Start Date"] = min(section_dates)
            section["Due Date"] = max(section_dates)

    schedule_dates = [
        value
        for task in tasks
        for value in (task["Start Date"], task["Due Date"])
        if value
    ]
    schedule = {
        "metadata": {
            "projectName": PROJECT_NAME,
            "clientName": "",
            "startDate": min(schedule_dates) if schedule_dates else "",
            "endDate": max(schedule_dates) if schedule_dates else "",
        },
        "tasks": tasks,
    }
    OUTPUT_PATH.write_text(json.dumps(schedule, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"Converted {len(tasks)} Gantt items from {SOURCE_PATH.name} to {OUTPUT_PATH.name}")


if __name__ == "__main__":
    convert()