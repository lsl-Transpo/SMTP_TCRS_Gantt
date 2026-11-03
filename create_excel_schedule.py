import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

json_path = r"Task 6.1 - Management & Administration\schedule_data.json"
excel_path = r"Task 6.1 - Management & Administration\SMTP_TCRS_Schedule.xlsx"

with open(json_path, "r", encoding="utf-8") as f:
    tasks = json.load(f)

wb = openpyxl.Workbook()
# Remove default sheet
wb.remove(wb.active)

# Color palettes
NAVY_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
SMTP_FILL = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid")
TCRS_FILL = PatternFill(start_color="059669", end_color="059669", fill_type="solid")
DASH_FILL = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
WHITE_BOLD = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
REGULAR_FONT = Font(name="Segoe UI", size=10)
BOLD_FONT = Font(name="Segoe UI", size=10, bold=True)
MUTED_FONT = Font(name="Segoe UI", size=9, color="64748B")

ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

THIN_SIDE = Side(border_style="thin", color="E2E8F0")
CELL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

COLUMNS = [
    ("Project", 10, ALIGN_CENTER),
    ("WBS", 10, ALIGN_CENTER),
    ("Task Name", 42, ALIGN_LEFT),
    ("Task Type", 15, ALIGN_CENTER),
    ("Status", 14, ALIGN_CENTER),
    ("Start Date", 13, ALIGN_CENTER),
    ("Due Date", 13, ALIGN_CENTER),
    ("Duration", 10, ALIGN_CENTER),
    ("Assignees", 22, ALIGN_LEFT),
    ("Priority", 11, ALIGN_CENTER),
    ("Task Order", 16, ALIGN_CENTER),
    ("Parent Task", 32, ALIGN_LEFT),
    ("Task ID", 12, ALIGN_CENTER),
    ("Parent ID", 12, ALIGN_CENTER),
    ("Description", 45, ALIGN_LEFT),
]

def add_schedule_sheet(sheet_title, task_list, header_fill):
    ws = wb.create_sheet(title=sheet_title)
    ws.views.sheetView[0].showGridLines = True
    
    # Write headers
    for col_idx, (col_name, col_width, align) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = WHITE_BOLD
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = CELL_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = col_width
        
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    
    # Write rows
    for row_idx, t in enumerate(task_list, start=2):
        level = t.get("Level", 1)
        indent = "    " * (level - 1)
        name_display = f"{indent}{t['Task Name']}"
        is_parent = level == 1
        row_font = BOLD_FONT if is_parent else REGULAR_FONT
        row_fill = ZEBRA_FILL if row_idx % 2 == 0 else WHITE_FILL
        
        row_data = [
            t.get("Project", ""),
            t.get("WBS", ""),
            name_display,
            t.get("Task Type", "Task"),
            t.get("Status", "To Do"),
            t.get("Start Date", ""),
            t.get("Due Date", ""),
            f'=IF(AND(G{row_idx}<>"",F{row_idx}<>""), G{row_idx}-F{row_idx}+1, "")',
            t.get("Assignees", ""),
            t.get("Priority", "Normal"),
            t.get("Tags", ""),
            t.get("Parent Task", ""),
            t.get("Task ID", ""),
            t.get("Parent ID", ""),
            t.get("Description", "")
        ]
        
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = row_font
            cell.fill = row_fill
            cell.alignment = COLUMNS[col_idx-1][2]
            cell.border = CELL_BORDER
            
            # Format date cells
            if col_idx in (6, 7) and val:
                cell.number_format = "YYYY-MM-DD"
            elif col_idx == 8:
                cell.number_format = "#,##0"
                
        ws.row_dimensions[row_idx].height = 22
        
    # Add Data Validations
    max_row = len(task_list) + 100
    
    dv_proj = DataValidation(type="list", formula1='"SMTP,TCRS"', allow_blank=True)
    ws.add_data_validation(dv_proj)
    dv_proj.add(f"A2:A{max_row}")
    
    dv_type = DataValidation(type="list", formula1='"Task,Deliverable,Milestone,Decision Point"', allow_blank=True)
    ws.add_data_validation(dv_type)
    dv_type.add(f"D2:D{max_row}")
    
    dv_status = DataValidation(type="list", formula1='"To Do,In Progress,Complete"', allow_blank=True)
    ws.add_data_validation(dv_status)
    dv_status.add(f"E2:E{max_row}")
    
    dv_prio = DataValidation(type="list", formula1='"Urgent,High,Normal,Low"', allow_blank=True)
    ws.add_data_validation(dv_prio)
    dv_prio.add(f"J2:J{max_row}")
    
    # Conditional formatting for Status
    green_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    green_font = Font(color="15803D", bold=True, name="Segoe UI")
    ws.conditional_formatting.add(f"E2:E{max_row}", CellIsRule(operator="equal", formula=['"Complete"'], fill=green_fill, font=green_font))
    
    blue_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    blue_font = Font(color="1D4ED8", bold=True, name="Segoe UI")
    ws.conditional_formatting.add(f"E2:E{max_row}", CellIsRule(operator="equal", formula=['"In Progress"'], fill=blue_fill, font=blue_font))
    
    gray_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    gray_font = Font(color="64748B", bold=False, name="Segoe UI")
    ws.conditional_formatting.add(f"E2:E{max_row}", CellIsRule(operator="equal", formula=['"To Do"'], fill=gray_fill, font=gray_font))
    
    # Conditional formatting for Task Type
    amber_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    amber_font = Font(color="B45309", bold=True, name="Segoe UI")
    ws.conditional_formatting.add(f"D2:D{max_row}", CellIsRule(operator="equal", formula=['"Milestone"'], fill=amber_fill, font=amber_font))
    
    purple_fill = PatternFill(start_color="F3E8FF", end_color="F3E8FF", fill_type="solid")
    purple_font = Font(color="7E22CE", bold=True, name="Segoe UI")
    ws.conditional_formatting.add(f"D2:D{max_row}", CellIsRule(operator="equal", formula=['"Deliverable"'], fill=purple_fill, font=purple_font))
    
    rose_fill = PatternFill(start_color="FFE4E6", end_color="FFE4E6", fill_type="solid")
    rose_font = Font(color="BE123C", bold=True, name="Segoe UI")
    ws.conditional_formatting.add(f"D2:D{max_row}", CellIsRule(operator="equal", formula=['"Decision Point"'], fill=rose_fill, font=rose_font))

# 1. Master Schedule (Both SMTP & TCRS)
add_schedule_sheet("All Tasks (Master)", tasks, NAVY_FILL)

# 2. SMTP Project Sheet
smtp_tasks = [t for t in tasks if t.get("Project") == "SMTP"]
add_schedule_sheet("SMTP Tasks", smtp_tasks, SMTP_FILL)

# 3. TCRS Project Sheet
tcrs_tasks = [t for t in tasks if t.get("Project") == "TCRS"]
add_schedule_sheet("TCRS Tasks", tcrs_tasks, TCRS_FILL)

# 4. Project Dashboard Sheet
ws_dash = wb.create_sheet(title="Dashboard & Summary", index=0)
ws_dash.views.sheetView[0].showGridLines = True

# Title banner
ws_dash.merge_cells("A1:G2")
banner = ws_dash["A1"]
banner.value = "WSDOT SMTP & TCRS Project Schedule Dashboard"
banner.fill = DASH_FILL
banner.font = Font(name="Segoe UI", size=16, bold=True, color="FFFFFF")
banner.alignment = Alignment(horizontal="center", vertical="center")

# Subtitle
ws_dash["A3"] = "Combined Statewide Multimodal Transportation Plan (SMTP) & Carbon Reduction Strategy (TCRS)"
ws_dash["A3"].font = Font(name="Segoe UI", size=11, italic=True, color="475569")
ws_dash.row_dimensions[1].height = 24
ws_dash.row_dimensions[2].height = 24
ws_dash.row_dimensions[3].height = 20

# KPI Cards Setup
kpis = [
    ("B5:B6", "Total Tasks", f'=COUNTA(\'All Tasks (Master)\'!C2:C{len(tasks)+1})', "3B82F6"),
    ("C5:C6", "Completed", f'=COUNTIF(\'All Tasks (Master)\'!E2:E{len(tasks)+1}, "Complete")', "10B981"),
    ("D5:D6", "In Progress", f'=COUNTIF(\'All Tasks (Master)\'!E2:E{len(tasks)+1}, "In Progress")', "6366F1"),
    ("E5:E6", "To Do", f'=COUNTIF(\'All Tasks (Master)\'!E2:E{len(tasks)+1}, "To Do")', "64748B"),
    ("F5:F6", "% Complete", f'=C6/B6', "059669")
]

for cell_range, label, formula, hex_color in kpis:
    top_cell_ref = cell_range.split(":")[0]
    bot_cell_ref = cell_range.split(":")[1]
    
    c_label = ws_dash[top_cell_ref]
    c_label.value = label
    c_label.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
    c_label.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
    c_label.alignment = Alignment(horizontal="center", vertical="center")
    
    c_val = ws_dash[bot_cell_ref]
    c_val.value = formula
    c_val.font = Font(name="Segoe UI", size=16, bold=True, color="1E293B")
    c_val.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    c_val.alignment = Alignment(horizontal="center", vertical="center")
    c_val.border = CELL_BORDER
    if label == "% Complete":
        c_val.number_format = "0.0%"
    else:
        c_val.number_format = "#,##0"

ws_dash.row_dimensions[5].height = 18
ws_dash.row_dimensions[6].height = 32

# Project Breakdown Table
ws_dash["B9"] = "Project Breakdown"
ws_dash["B9"].font = Font(name="Segoe UI", size=13, bold=True, color="1E293B")

dash_headers = ["Project", "Total Items", "Complete", "In Progress", "To Do", "% Progress"]
for idx, h in enumerate(dash_headers, start=2):
    c = ws_dash.cell(row=10, column=idx, value=h)
    c.fill = DASH_FILL
    c.font = WHITE_BOLD
    c.alignment = ALIGN_CENTER
    c.border = CELL_BORDER
ws_dash.row_dimensions[10].height = 24

proj_rows = [
    ("SMTP", 11, f'=COUNTIF(\'All Tasks (Master)\'!A2:A{len(tasks)+1}, "SMTP")',
     f'=COUNTIFS(\'All Tasks (Master)\'!A2:A{len(tasks)+1}, "SMTP", \'All Tasks (Master)\'!E2:E{len(tasks)+1}, "Complete")',
     f'=COUNTIFS(\'All Tasks (Master)\'!A2:A{len(tasks)+1}, "SMTP", \'All Tasks (Master)\'!E2:E{len(tasks)+1}, "In Progress")',
     f'=COUNTIFS(\'All Tasks (Master)\'!A2:A{len(tasks)+1}, "SMTP", \'All Tasks (Master)\'!E2:E{len(tasks)+1}, "To Do")',
     '=D11/C11'),
    ("TCRS", 12, f'=COUNTIF(\'All Tasks (Master)\'!A2:A{len(tasks)+1}, "TCRS")',
     f'=COUNTIFS(\'All Tasks (Master)\'!A2:A{len(tasks)+1}, "TCRS", \'All Tasks (Master)\'!E2:E{len(tasks)+1}, "Complete")',
     f'=COUNTIFS(\'All Tasks (Master)\'!A2:A{len(tasks)+1}, "TCRS", \'All Tasks (Master)\'!E2:E{len(tasks)+1}, "In Progress")',
     f'=COUNTIFS(\'All Tasks (Master)\'!A2:A{len(tasks)+1}, "TCRS", \'All Tasks (Master)\'!E2:E{len(tasks)+1}, "To Do")',
     '=D12/C12'),
    ("Total", 13, '=SUM(C11:C12)', '=SUM(D11:D12)', '=SUM(E11:E12)', '=SUM(F11:F12)', '=D13/C13')
]

for proj_name, row_num, total_f, comp_f, prog_f, todo_f, pct_f in proj_rows:
    ws_dash.cell(row=row_num, column=2, value=proj_name).font = BOLD_FONT
    ws_dash.cell(row=row_num, column=2).alignment = ALIGN_CENTER
    ws_dash.cell(row=row_num, column=2).border = CELL_BORDER
    
    ws_dash.cell(row=row_num, column=3, value=total_f).alignment = ALIGN_CENTER
    ws_dash.cell(row=row_num, column=3).number_format = "#,##0"
    ws_dash.cell(row=row_num, column=3).border = CELL_BORDER
    
    ws_dash.cell(row=row_num, column=4, value=comp_f).alignment = ALIGN_CENTER
    ws_dash.cell(row=row_num, column=4).number_format = "#,##0"
    ws_dash.cell(row=row_num, column=4).border = CELL_BORDER
    
    ws_dash.cell(row=row_num, column=5, value=prog_f).alignment = ALIGN_CENTER
    ws_dash.cell(row=row_num, column=5).number_format = "#,##0"
    ws_dash.cell(row=row_num, column=5).border = CELL_BORDER
    
    ws_dash.cell(row=row_num, column=6, value=todo_f).alignment = ALIGN_CENTER
    ws_dash.cell(row=row_num, column=6).number_format = "#,##0"
    ws_dash.cell(row=row_num, column=6).border = CELL_BORDER
    
    pct_cell = ws_dash.cell(row=row_num, column=7, value=pct_f)
    pct_cell.alignment = ALIGN_CENTER
    pct_cell.number_format = "0.0%"
    pct_cell.font = BOLD_FONT
    pct_cell.border = CELL_BORDER
    
    if proj_name == "Total":
        for c_idx in range(2, 8):
            ws_dash.cell(row=row_num, column=c_idx).fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            ws_dash.cell(row=row_num, column=c_idx).font = BOLD_FONT
            
    ws_row_dim = ws_dash.row_dimensions[row_num]
    ws_row_dim.height = 22

# Item Type Breakdown Table
ws_dash["B15"] = "Breakdown by Item Type"
ws_dash["B15"].font = Font(name="Segoe UI", size=13, bold=True, color="1E293B")

type_headers = ["Item Type", "Count", "Complete", "In Progress", "To Do", "% Done"]
for idx, h in enumerate(type_headers, start=2):
    c = ws_dash.cell(row=16, column=idx, value=h)
    c.fill = DASH_FILL
    c.font = WHITE_BOLD
    c.alignment = ALIGN_CENTER
    c.border = CELL_BORDER
ws_dash.row_dimensions[16].height = 24

types_info = [
    ("Tasks", "Task", 17),
    ("Deliverables", "Deliverable", 18),
    ("Milestones", "Milestone", 19),
    ("Decision Points", "Decision Point", 20)
]

for label, type_code, row_num in types_info:
    ws_dash.cell(row=row_num, column=2, value=label).font = BOLD_FONT
    ws_dash.cell(row=row_num, column=2).alignment = ALIGN_CENTER
    ws_dash.cell(row=row_num, column=2).border = CELL_BORDER
    
    tot_f = f'=COUNTIF(\'All Tasks (Master)\'!D2:D{len(tasks)+1}, "{type_code}")'
    comp_f = f'=COUNTIFS(\'All Tasks (Master)\'!D2:D{len(tasks)+1}, "{type_code}", \'All Tasks (Master)\'!E2:E{len(tasks)+1}, "Complete")'
    prog_f = f'=COUNTIFS(\'All Tasks (Master)\'!D2:D{len(tasks)+1}, "{type_code}", \'All Tasks (Master)\'!E2:E{len(tasks)+1}, "In Progress")'
    todo_f = f'=COUNTIFS(\'All Tasks (Master)\'!D2:D{len(tasks)+1}, "{type_code}", \'All Tasks (Master)\'!E2:E{len(tasks)+1}, "To Do")'
    pct_f = f'=IF(C{row_num}>0, D{row_num}/C{row_num}, 0)'
    
    ws_dash.cell(row=row_num, column=3, value=tot_f).alignment = ALIGN_CENTER
    ws_dash.cell(row=row_num, column=3).number_format = "#,##0"
    ws_dash.cell(row=row_num, column=3).border = CELL_BORDER
    
    ws_dash.cell(row=row_num, column=4, value=comp_f).alignment = ALIGN_CENTER
    ws_dash.cell(row=row_num, column=4).number_format = "#,##0"
    ws_dash.cell(row=row_num, column=4).border = CELL_BORDER
    
    ws_dash.cell(row=row_num, column=5, value=prog_f).alignment = ALIGN_CENTER
    ws_dash.cell(row=row_num, column=5).number_format = "#,##0"
    ws_dash.cell(row=row_num, column=5).border = CELL_BORDER
    
    ws_dash.cell(row=row_num, column=6, value=todo_f).alignment = ALIGN_CENTER
    ws_dash.cell(row=row_num, column=6).number_format = "#,##0"
    ws_dash.cell(row=row_num, column=6).border = CELL_BORDER
    
    pct_cell = ws_dash.cell(row=row_num, column=7, value=pct_f)
    pct_cell.alignment = ALIGN_CENTER
    pct_cell.number_format = "0.0%"
    pct_cell.font = BOLD_FONT
    pct_cell.border = CELL_BORDER
    ws_dash.row_dimensions[row_num].height = 22

# Instructions box on the side
ws_dash["B23"] = "💡 How to Use This Workbook"
ws_dash["B23"].font = Font(name="Segoe UI", size=13, bold=True, color="1E293B")

instructions = [
    "1. Shared Editing: Team members can edit dates, statuses, assignees, or add new tasks in any sheet.",
    "2. Synchronized Views: Use the 'All Tasks (Master)' sheet for the whole picture, or project tabs for focused work.",
    "3. Status & Type Dropdowns: Click dropdown arrows in Status, Type, and Priority columns to select standardized values.",
    "4. Interactive Gantt Sync: You can load this Excel file directly into the Interactive Gantt Chart (SMTP_TCRS_Interactive_Gantt.html) anytime,",
    "   or export updated Excel files directly from the web interface to keep everyone aligned."
]

for idx, text in enumerate(instructions, start=24):
    c = ws_dash.cell(row=idx, column=2, value=text)
    c.font = Font(name="Segoe UI", size=10, color="334155")
    ws_dash.row_dimensions[idx].height = 20

# Column widths for dashboard
ws_dash.column_dimensions["A"].width = 4
ws_dash.column_dimensions["B"].width = 24
ws_dash.column_dimensions["C"].width = 16
ws_dash.column_dimensions["D"].width = 16
ws_dash.column_dimensions["E"].width = 16
ws_dash.column_dimensions["F"].width = 16
ws_dash.column_dimensions["G"].width = 16

wb.save(excel_path)
print(f"Successfully generated {excel_path}!")

