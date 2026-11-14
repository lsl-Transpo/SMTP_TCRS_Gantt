# Interactive Gantt Chart Template

Use this repository to create a standalone interactive Gantt chart for a new project. The setup form creates the schedule input; the Python builder generates `index.html` from it.

## Create a chart

1. Open [`setup.html`](setup.html) in a browser.
2. Enter the project name, client name, overall timeline, and task dates. Turn on **Include subtasks** or **Include milestones** as needed; subtasks can be assigned to a parent task.
3. Choose **Download schedule file**. Save the downloaded `schedule_data.json` over the file with the same name in this repository.
4. Run `python build_interactive_gantt.py` from this folder. This replaces `index.html` with the chart for the entered project.
5. Open [`index.html`](index.html), or double-click [`Start_Gantt_Chart.bat`](Start_Gantt_Chart.bat).

Python 3 is required to build the chart. The generated chart itself runs in a modern browser without a server. Task edits are saved in that browser; use its Excel or CSV export buttons to download a copy.

## Files

- [`setup.html`](setup.html): project and task entry form.
- [`schedule_data.json`](schedule_data.json): schedule input. The form-generated format includes project metadata and task records; the builder also accepts the existing flat task-list format.
- [`build_interactive_gantt.py`](build_interactive_gantt.py): generates the interactive chart.
- [`index.html`](index.html): generated chart. Rebuild it after changing the setup form data.
- [`Start_Gantt_Chart.bat`](Start_Gantt_Chart.bat): opens the generated chart in the default browser.

The form starts with one task row. Add more items as needed. A milestone uses its due date as its marker; tasks require a start and due date. Subtasks are one level below their selected parent.

