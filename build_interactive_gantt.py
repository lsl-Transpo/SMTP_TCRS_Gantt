import json
import os

json_path = r"Task 6.1 - Management & Administration\schedule_data.json"
html_path = r"Task 6.1 - Management & Administration\SMTP_TCRS_Interactive_Gantt.html"

with open(json_path, "r", encoding="utf-8") as f:
    tasks = json.load(f)

smtp_count = sum(1 for t in tasks if t["Project"] == "SMTP")
tcrs_count = sum(1 for t in tasks if t["Project"] == "TCRS")
total_count = len(tasks)

json_tasks_str = json.dumps(tasks, indent=2)

html_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>WSDOT SMTP & TCRS Interactive Project Gantt</title>
  <!-- Load SheetJS for Excel .xlsx read and write -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    :root {{
      --primary-smtp: #0284c7;
      --primary-smtp-dark: #0369a1;
      --primary-tcrs: #059669;
      --primary-tcrs-dark: #047857;
      --bg-dark: #0f172a;
      --bg-header: #1e293b;
      --bg-surface: #ffffff;
      --bg-subtle: #f8fafc;
      --bg-muted: #f1f5f9;
      --border-color: #e2e8f0;
      --border-dark: #cbd5e1;
      --text-main: #0f172a;
      --text-muted: #64748b;
      --text-light: #94a3b8;
      
      --status-complete: #10b981;
      --status-complete-bg: #ecfdf5;
      --status-progress: #3b82f6;
      --status-progress-bg: #eff6ff;
      --status-todo: #94a3b8;
      --status-todo-bg: #f8fafc;
      
      --type-milestone: #f59e0b;
      --type-deliverable: #8b5cf6;
      --type-decision: #ec4899;
      --type-task: #3b82f6;

      --row-height: 38px;
    }}

    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }}

    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      background: var(--bg-subtle);
      color: var(--text-main);
      height: 100vh;
      display: flex;
      flex-direction: column;
      overflow: hidden;
      user-select: none;
    }}

    /* Top App Header */
    header.app-header {{
      background: var(--bg-header);
      color: #ffffff;
      padding: 10px 20px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      box-shadow: 0 2px 8px rgba(0,0,0,0.15);
      z-index: 30;
      flex-shrink: 0;
    }}

    .header-branding {{
      display: flex;
      align-items: center;
      gap: 14px;
    }}

    .logo-badge {{
      background: linear-gradient(135deg, #0284c7, #059669);
      color: white;
      font-weight: 800;
      font-size: 16px;
      letter-spacing: 0.5px;
      padding: 6px 12px;
      border-radius: 6px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }}

    .header-title-box h1 {{
      font-size: 17px;
      font-weight: 700;
      letter-spacing: -0.2px;
      display: flex;
      align-items: center;
      gap: 10px;
    }}

    .header-title-box p {{
      font-size: 12px;
      color: #94a3b8;
    }}

    .header-controls {{
      display: flex;
      align-items: center;
      gap: 10px;
    }}

    /* Action Buttons */
    .btn {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 6px 12px;
      border-radius: 6px;
      font-size: 12.5px;
      font-weight: 600;
      cursor: pointer;
      border: 1px solid transparent;
      transition: all 0.15s ease;
      white-space: nowrap;
    }}

    .btn-primary {{
      background: #2563eb;
      color: #fff;
    }}
    .btn-primary:hover {{
      background: #1d4ed8;
    }}

    .btn-smtp {{
      background: #0284c7;
      color: #fff;
    }}
    .btn-smtp:hover {{
      background: #0369a1;
    }}

    .btn-tcrs {{
      background: #059669;
      color: #fff;
    }}
    .btn-tcrs:hover {{
      background: #047857;
    }}

    .btn-outline {{
      background: rgba(255,255,255,0.08);
      color: #f1f5f9;
      border-color: rgba(255,255,255,0.2);
    }}
    .btn-outline:hover {{
      background: rgba(255,255,255,0.18);
      color: #fff;
    }}

    .btn-light {{
      background: #ffffff;
      color: #1e293b;
      border-color: var(--border-color);
    }}
    .btn-light:hover {{
      background: #f8fafc;
      border-color: #cbd5e1;
    }}

    .btn-success {{
      background: #10b981;
      color: white;
    }}
    .btn-success:hover {{
      background: #059669;
    }}

    /* Project Filter Tab Pill */
    .project-tabs {{
      display: flex;
      background: rgba(0,0,0,0.3);
      padding: 3px;
      border-radius: 8px;
      gap: 3px;
    }}

    .tab-btn {{
      padding: 5px 12px;
      font-size: 12px;
      font-weight: 600;
      border-radius: 6px;
      color: #cbd5e1;
      cursor: pointer;
      border: none;
      background: transparent;
      display: flex;
      align-items: center;
      gap: 6px;
      transition: all 0.15s;
    }}

    .tab-btn:hover {{
      color: #fff;
    }}

    .tab-btn.active {{
      background: #ffffff;
      color: #0f172a;
      box-shadow: 0 1px 3px rgba(0,0,0,0.2);
    }}

    .tab-btn.active[data-project="SMTP"] {{
      background: #38bdf8;
      color: #0c4a6e;
    }}

    .tab-btn.active[data-project="TCRS"] {{
      background: #34d399;
      color: #064e3b;
    }}

    /* Sub Toolbar */
    .toolbar {{
      background: #ffffff;
      border-bottom: 1px solid var(--border-color);
      padding: 8px 16px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      flex-wrap: wrap;
      z-index: 20;
      flex-shrink: 0;
    }}

    .toolbar-left, .toolbar-right {{
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }}

    .search-box {{
      position: relative;
      display: flex;
      align-items: center;
    }}

    .search-box input {{
      padding: 6px 10px 6px 30px;
      font-size: 12.5px;
      border: 1px solid var(--border-color);
      border-radius: 6px;
      width: 220px;
      outline: none;
      transition: all 0.15s;
    }}
    .search-box input:focus {{
      border-color: #3b82f6;
      width: 260px;
      box-shadow: 0 0 0 2px rgba(59,130,246,0.15);
    }}
    .search-box .search-icon {{
      position: absolute;
      left: 9px;
      color: var(--text-muted);
      font-size: 13px;
      pointer-events: none;
    }}

    .select-filter {{
      padding: 6px 10px;
      font-size: 12px;
      border: 1px solid var(--border-color);
      border-radius: 6px;
      background: #ffffff;
      color: var(--text-main);
      cursor: pointer;
      outline: none;
    }}

    .select-filter:focus {{
      border-color: #3b82f6;
    }}

    /* View Zoom Controls */
    .zoom-group {{
      display: flex;
      border: 1px solid var(--border-color);
      border-radius: 6px;
      overflow: hidden;
    }}

    .zoom-btn {{
      padding: 5px 11px;
      font-size: 11.5px;
      font-weight: 600;
      background: #ffffff;
      border: none;
      border-right: 1px solid var(--border-color);
      cursor: pointer;
      color: var(--text-muted);
      transition: all 0.15s;
    }}
    .zoom-btn:last-child {{
      border-right: none;
    }}
    .zoom-btn:hover {{
      background: #f1f5f9;
      color: var(--text-main);
    }}
    .zoom-btn.active {{
      background: #0f172a;
      color: #ffffff;
    }}

    /* Stats Quick Strip */
    .stats-strip {{
      background: #f8fafc;
      border-bottom: 1px solid var(--border-color);
      padding: 5px 16px;
      display: flex;
      align-items: center;
      gap: 16px;
      font-size: 12px;
      color: var(--text-muted);
      flex-shrink: 0;
    }}

    .stat-pill {{
      display: flex;
      align-items: center;
      gap: 6px;
      font-weight: 500;
    }}

    .stat-pill strong {{
      color: var(--text-main);
    }}

    .dot {{
      width: 8px;
      height: 8px;
      border-radius: 50%;
      display: inline-block;
    }}
    .dot-complete {{ background: var(--status-complete); }}
    .dot-progress {{ background: var(--status-progress); }}
    .dot-todo {{ background: var(--status-todo); }}

    .sync-status-badge {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 3px 8px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: 600;
      background: #ecfdf5;
      color: #047857;
      border: 1px solid #a7f3d0;
    }}
    .sync-status-badge.offline {{
      background: #f8fafc;
      color: #64748b;
      border-color: #cbd5e1;
    }}

    /* Main Container: Split into Left Table and Right Gantt */
    .gantt-main-container {{
      flex: 1;
      display: flex;
      overflow: hidden;
      position: relative;
    }}

    /* Left Pane: Hierarchical Table */
    .tree-pane {{
      width: 665px;
      min-width: 380px;
      max-width: 900px;
      background: #ffffff;
      border-right: 1px solid var(--border-color);
      display: flex;
      flex-direction: column;
      flex-shrink: 0;
      overflow: hidden;
      z-index: 10;
    }}

    .pane-header {{
      height: 50px;
      background: #f8fafc;
      border-bottom: 1px solid var(--border-color);
      display: flex;
      align-items: center;
      font-size: 11.5px;
      font-weight: 700;
      color: #475569;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      flex-shrink: 0;
      width: 665px;
      min-width: 665px;
      overflow-x: hidden;
    }}

    .th-col {{
      padding: 0 8px;
      display: flex;
      align-items: center;
      height: 100%;
      border-right: 1px solid var(--border-color);
      flex-shrink: 0;
      box-sizing: border-box;
    }}

    .tree-scroll {{
      flex: 1;
      overflow-y: hidden; /* Synced with timeline scroll */
      overflow-x: auto;
    }}

    .tree-table {{
      width: 665px;
      min-width: 665px;
      border-collapse: collapse;
      table-layout: fixed;
    }}

    .tree-row {{
      height: var(--row-height);
      border-bottom: 1px solid var(--border-color);
      cursor: pointer;
      transition: background 0.1s ease;
      font-size: 12px;
    }}

    .tree-row:hover {{
      background: #f1f5f9;
    }}

    .tree-row.selected {{
      background: #e0f2fe !important;
    }}

    .tree-row.is-parent {{
      background: #fbfcfe;
      font-weight: 600;
    }}

    .tree-row td {{
      padding: 0 8px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      vertical-align: middle;
      border-right: 1px solid #f1f5f9;
    }}

    .col-wbs {{
      width: 65px;
      min-width: 65px;
      max-width: 65px;
      text-align: center;
      color: #64748b;
      font-size: 11px;
      font-weight: 600;
      flex-shrink: 0;
      box-sizing: border-box;
    }}

    .col-name {{
      width: 280px;
      min-width: 280px;
      max-width: 280px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      flex-shrink: 0;
      box-sizing: border-box;
    }}

    .col-name-wrapper {{
      display: flex;
      align-items: center;
      gap: 6px;
      width: 100%;
      max-width: 100%;
      min-width: 0;
      height: var(--row-height);
      overflow: hidden;
      box-sizing: border-box;
    }}

    .col-name-text {{
      flex: 1 1 0%;
      min-width: 0;
      max-width: 100%;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      display: block;
    }}

    .col-type {{
      width: 105px;
      min-width: 105px;
      max-width: 105px;
      text-align: center;
      flex-shrink: 0;
      box-sizing: border-box;
    }}

    .col-status {{
      width: 95px;
      min-width: 95px;
      max-width: 95px;
      text-align: center;
      flex-shrink: 0;
      box-sizing: border-box;
    }}

    .col-assignee {{
      width: 120px;
      min-width: 120px;
      max-width: 120px;
      color: #475569;
      flex-shrink: 0;
      box-sizing: border-box;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }}

    .expander {{
      width: 16px;
      height: 16px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      border-radius: 3px;
      color: #64748b;
      font-size: 10px;
      transition: transform 0.15s ease;
      flex-shrink: 0;
    }}

    .expander:hover {{
      background: #e2e8f0;
      color: #0f172a;
    }}

    .expander.collapsed {{
      transform: rotate(-90deg);
    }}

    .indent-spacer {{
      display: inline-block;
      width: 16px;
      flex-shrink: 0;
    }}

    /* Badges */
    .badge {{
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 2px 7px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: 600;
      line-height: 1.3;
    }}

    .badge-smtp {{
      background: #e0f2fe;
      color: #0369a1;
      border: 1px solid #bae6fd;
    }}

    .badge-tcrs {{
      background: #dcfce7;
      color: #15803d;
      border: 1px solid #bbf7d0;
    }}

    .badge-complete {{
      background: var(--status-complete-bg);
      color: #065f46;
      border: 1px solid #a7f3d0;
    }}

    .badge-progress {{
      background: var(--status-progress-bg);
      color: #1e40af;
      border: 1px solid #bfdbfe;
    }}

    .badge-todo {{
      background: var(--status-todo-bg);
      color: #475569;
      border: 1px solid #cbd5e1;
    }}

    .badge-milestone {{
      background: #fef3c7;
      color: #92400e;
      border: 1px solid #fde68a;
    }}

    .badge-deliverable {{
      background: #f3e8ff;
      color: #6b21a8;
      border: 1px solid #e9d5ff;
    }}

    .badge-decision {{
      background: #ffe4e6;
      color: #9f1239;
      border: 1px solid #fecdd3;
    }}

    .badge-task {{
      background: #f1f5f9;
      color: #334155;
      border: 1px solid #e2e8f0;
    }}

    /* Resizer Splitter */
    .resizer {{
      width: 6px;
      background: #e2e8f0;
      cursor: col-resize;
      transition: background 0.15s ease;
      z-index: 15;
      flex-shrink: 0;
    }}
    .resizer:hover, .resizer.dragging {{
      background: #3b82f6;
    }}

    /* Right Pane: Gantt Timeline */
    .timeline-pane {{
      flex: 1;
      display: flex;
      flex-direction: column;
      overflow: hidden;
      background: #ffffff;
      position: relative;
    }}

    /* Timeline Header */
    .timeline-header-container {{
      height: 50px;
      background: #f8fafc;
      border-bottom: 1px solid var(--border-color);
      overflow-x: hidden;
      position: relative;
      flex-shrink: 0;
    }}

    .timeline-header-months, .timeline-header-weeks {{
      display: flex;
      height: 25px;
      border-bottom: 1px solid #e2e8f0;
      font-size: 11px;
      font-weight: 600;
      color: #475569;
    }}

    .timeline-header-weeks {{
      border-bottom: none;
      font-size: 10px;
      color: #64748b;
      font-weight: 500;
    }}

    .timeline-header-cell {{
      display: flex;
      align-items: center;
      justify-content: center;
      border-right: 1px solid #e2e8f0;
      overflow: hidden;
      white-space: nowrap;
      text-overflow: ellipsis;
      flex-shrink: 0;
    }}

    /* Timeline Body */
    .timeline-body-scroll {{
      flex: 1;
      overflow: auto;
      position: relative;
    }}

    .timeline-canvas {{
      position: relative;
      min-height: 100%;
    }}

    /* Grid Lines */
    .grid-column {{
      position: absolute;
      top: 0;
      bottom: 0;
      border-right: 1px solid #f8fafc;
      pointer-events: none;
    }}

    .grid-column.month-boundary {{
      border-right: 1px solid #e2e8f0;
    }}

    .timeline-row {{
      height: var(--row-height);
      position: relative;
      border-bottom: 1px solid #f1f5f9;
      transition: background 0.1s ease;
    }}

    .timeline-row:hover {{
      background: rgba(241, 245, 249, 0.6);
    }}

    .timeline-row.is-parent {{
      background: rgba(248, 250, 252, 0.7);
    }}

    .timeline-row.selected {{
      background: #f0f9ff !important;
    }}

    /* Today Vertical Line Marker */
    .today-line {{
      position: absolute;
      top: 0;
      bottom: 0;
      width: 2px;
      background: #ef4444;
      z-index: 12;
      pointer-events: none;
    }}

    .today-tag {{
      position: sticky;
      top: 0;
      left: -28px;
      background: #ef4444;
      color: white;
      font-size: 9.5px;
      font-weight: 700;
      padding: 1px 6px;
      border-radius: 4px;
      white-space: nowrap;
      box-shadow: 0 1px 3px rgba(0,0,0,0.2);
    }}

    /* Gantt Bars */
    .gantt-bar {{
      position: absolute;
      top: 6px;
      height: 26px;
      border-radius: 6px;
      cursor: pointer;
      display: flex;
      align-items: center;
      padding: 0 8px;
      font-size: 11px;
      font-weight: 600;
      color: #ffffff;
      box-shadow: 0 1px 3px rgba(0,0,0,0.12);
      transition: box-shadow 0.15s, transform 0.1s;
      overflow: hidden;
      white-space: nowrap;
      text-overflow: ellipsis;
      z-index: 5;
    }}

    .gantt-bar:hover {{
      box-shadow: 0 4px 10px rgba(0,0,0,0.2);
      transform: translateY(-1px);
      z-index: 9;
    }}

    /* Type Specific Bar Styles */
    .gantt-bar.task-bar-smtp {{
      background: linear-gradient(135deg, #0284c7, #0369a1);
      border: 1px solid #0284c7;
    }}

    .gantt-bar.task-bar-tcrs {{
      background: linear-gradient(135deg, #059669, #047857);
      border: 1px solid #059669;
    }}

    .gantt-bar.deliverable-bar {{
      background: linear-gradient(135deg, #7c3aed, #6d28d9);
      border: 1.5px solid #a78bfa;
      border-left: 4px solid #c4b5fd;
    }}

    .gantt-bar.parent-summary-bar {{
      height: 8px;
      top: 15px;
      border-radius: 3px;
      background: #334155;
    }}

    /* Milestone Diamond */
    .milestone-marker {{
      position: absolute;
      top: 10px;
      width: 18px;
      height: 18px;
      background: #f59e0b;
      transform: rotate(45deg);
      border: 2px solid #ffffff;
      box-shadow: 0 2px 5px rgba(0,0,0,0.25);
      cursor: pointer;
      z-index: 6;
      transition: all 0.15s ease;
    }}

    .milestone-marker:hover {{
      transform: rotate(45deg) scale(1.25);
      background: #d97706;
      z-index: 10;
    }}

    /* Decision Point Marker */
    .decision-marker {{
      position: absolute;
      top: 8px;
      width: 22px;
      height: 22px;
      background: #ec4899;
      border-radius: 5px;
      border: 2px solid #ffffff;
      box-shadow: 0 2px 5px rgba(0,0,0,0.25);
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 11px;
      color: white;
      cursor: pointer;
      z-index: 6;
      transition: all 0.15s ease;
    }}

    .decision-marker:hover {{
      transform: scale(1.25);
      background: #db2777;
      z-index: 10;
    }}

    .bar-label-outside {{
      position: absolute;
      left: calc(100% + 8px);
      top: 5px;
      font-size: 11px;
      font-weight: 600;
      color: #334155;
      white-space: nowrap;
      pointer-events: none;
    }}

    /* Resize Handles on Bars */
    .resize-handle {{
      position: absolute;
      top: 0;
      bottom: 0;
      width: 6px;
      cursor: ew-resize;
      z-index: 7;
    }}
    .resize-handle.left {{
      left: 0;
      border-top-left-radius: 6px;
      border-bottom-left-radius: 6px;
    }}
    .resize-handle.right {{
      right: 0;
      border-top-right-radius: 6px;
      border-bottom-right-radius: 6px;
    }}
    .resize-handle:hover {{
      background: rgba(255,255,255,0.4);
    }}

    /* Floating Tooltip */
    #gantt-tooltip {{
      position: fixed;
      background: rgba(15, 23, 42, 0.95);
      color: #ffffff;
      padding: 10px 14px;
      border-radius: 8px;
      font-size: 12px;
      line-height: 1.45;
      max-width: 340px;
      pointer-events: none;
      z-index: 1000;
      box-shadow: 0 4px 14px rgba(0,0,0,0.3);
      display: none;
      backdrop-filter: blur(4px);
      border: 1px solid rgba(255,255,255,0.1);
    }}

    #gantt-tooltip h4 {{
      font-size: 13px;
      font-weight: 700;
      margin-bottom: 5px;
      color: #38bdf8;
    }}

    #gantt-tooltip .meta-line {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 3px;
      font-size: 11px;
      color: #cbd5e1;
    }}

    #gantt-tooltip .desc-line {{
      margin-top: 6px;
      padding-top: 6px;
      border-top: 1px solid rgba(255,255,255,0.15);
      font-size: 11px;
      color: #94a3b8;
    }}

    /* Bottom Status Bar */
    footer.app-footer {{
      background: #ffffff;
      border-top: 1px solid var(--border-color);
      padding: 6px 16px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      font-size: 12px;
      color: var(--text-muted);
      flex-shrink: 0;
    }}

    .unscheduled-pill {{
      background: #fef2f2;
      color: #b91c1c;
      border: 1px solid #fecaca;
      padding: 2px 8px;
      border-radius: 12px;
      font-weight: 600;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 5px;
    }}

    .unscheduled-pill:hover {{
      background: #fee2e2;
    }}

    /* Modal Dialogs */
    .modal-overlay {{
      position: fixed;
      inset: 0;
      background: rgba(15, 23, 42, 0.6);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 2000;
      opacity: 0;
      pointer-events: none;
      transition: opacity 0.2s ease;
      backdrop-filter: blur(2px);
    }}

    .modal-overlay.active {{
      opacity: 1;
      pointer-events: auto;
    }}

    .modal-card {{
      background: #ffffff;
      border-radius: 12px;
      width: 580px;
      max-width: 95vw;
      max-height: 90vh;
      display: flex;
      flex-direction: column;
      box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.2), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
      overflow: hidden;
      transform: scale(0.96);
      transition: transform 0.2s ease;
    }}

    .modal-overlay.active .modal-card {{
      transform: scale(1);
    }}

    .modal-header {{
      background: var(--bg-header);
      color: #ffffff;
      padding: 14px 20px;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }}

    .modal-header h3 {{
      font-size: 15px;
      font-weight: 700;
    }}

    .modal-close-btn {{
      background: transparent;
      border: none;
      color: #94a3b8;
      font-size: 20px;
      cursor: pointer;
      line-height: 1;
    }}
    .modal-close-btn:hover {{
      color: #ffffff;
    }}

    .modal-body {{
      padding: 20px;
      overflow-y: auto;
      display: flex;
      flex-direction: column;
      gap: 14px;
    }}

    .form-row {{
      display: flex;
      gap: 12px;
    }}

    .form-group {{
      display: flex;
      flex-direction: column;
      gap: 4px;
      flex: 1;
    }}

    .form-group label {{
      font-size: 11.5px;
      font-weight: 600;
      color: #475569;
    }}

    .form-control {{
      padding: 8px 10px;
      border: 1px solid var(--border-color);
      border-radius: 6px;
      font-size: 13px;
      outline: none;
      background: #ffffff;
      color: var(--text-main);
    }}
    .form-control:focus {{
      border-color: #3b82f6;
      box-shadow: 0 0 0 2px rgba(59,130,246,0.15);
    }}

    textarea.form-control {{
      resize: vertical;
      min-height: 70px;
    }}

    .modal-footer {{
      padding: 12px 20px;
      background: #f8fafc;
      border-top: 1px solid var(--border-color);
      display: flex;
      align-items: center;
      justify-content: space-between;
    }}

    /* Toast Notification */
    .toast {{
      position: fixed;
      bottom: 24px;
      right: 24px;
      background: #0f172a;
      color: white;
      padding: 10px 18px;
      border-radius: 8px;
      box-shadow: 0 10px 15px -3px rgba(0,0,0,0.3);
      font-size: 13px;
      font-weight: 500;
      z-index: 3000;
      opacity: 0;
      transform: translateY(10px);
      transition: all 0.25s ease;
      display: flex;
      align-items: center;
      gap: 8px;
    }}

    .toast.show {{
      opacity: 1;
      transform: translateY(0);
    }}

    /* Drag & Drop Overlay */
    .dropzone-overlay {{
      position: fixed;
      inset: 0;
      background: rgba(37, 99, 235, 0.85);
      color: white;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      z-index: 5000;
      opacity: 0;
      pointer-events: none;
      transition: opacity 0.2s ease;
      font-size: 20px;
      font-weight: 700;
      gap: 12px;
    }}

    .dropzone-overlay.active {{
      opacity: 1;
      pointer-events: auto;
    }}

    /* Print styles */
    @media print {{
      header.app-header, .toolbar, .stats-strip, footer.app-footer, .resizer {{
        display: none !important;
      }}
      body, .gantt-main-container, .tree-pane, .timeline-pane {{
        height: auto !important;
        overflow: visible !important;
      }}
      .tree-pane {{
        width: 40% !important;
      }}
      .timeline-pane {{
        width: 60% !important;
      }}
    }}
  </style>
</head>
<body>

  <!-- Top App Header -->
  <header class="app-header">
    <div class="header-branding">
      <div class="logo-badge">WSDOT</div>
      <div class="header-title-box">
        <h1>SMTP & TCRS Project Schedule Gantt</h1>
        <p>Statewide Multimodal Transportation Plan (SMTP) & Transportation Carbon Reduction Strategy (TCRS)</p>
      </div>
    </div>

    <!-- Project Tabs -->
    <div class="project-tabs">
      <button class="tab-btn active" data-project="ALL">🌐 All Projects ({total_count})</button>
      <button class="tab-btn" data-project="SMTP">🔵 SMTP ({smtp_count})</button>
      <button class="tab-btn" data-project="TCRS">🟢 TCRS ({tcrs_count})</button>
    </div>

    <!-- Global Actions -->
    <div class="header-controls">
      <button class="btn btn-outline" id="btn-import-file" title="Load Excel (.xlsx) or CSV">📂 Load File</button>
      <input type="file" id="file-input" accept=".xlsx, .xls, .csv" style="display:none;" />
      
      <button class="btn btn-primary" id="btn-export-excel" title="Download updated Excel workbook">📥 Export Excel (.xlsx)</button>
      <button class="btn btn-outline" id="btn-export-csv" title="Download updated CSV">📊 Export CSV</button>
      <button class="btn btn-outline" id="btn-print" title="Print or save as PDF">🖨️ Print</button>
      <button class="btn btn-success" id="btn-add-task">➕ Add Task</button>
    </div>
  </header>

  <!-- Sub Toolbar with Search & Filters -->
  <div class="toolbar">
    <div class="toolbar-left">
      <!-- Search Input -->
      <div class="search-box">
        <span class="search-icon">🔍</span>
        <input type="text" id="search-input" placeholder="Search task name, note, or ID... (Press /)" />
      </div>

      <!-- Assignee Filter -->
      <select id="filter-assignee" class="select-filter">
        <option value="">👤 All Assignees</option>
      </select>

      <!-- Status Filter -->
      <select id="filter-status" class="select-filter">
        <option value="">⚡ All Statuses</option>
        <option value="Complete">✅ Complete</option>
        <option value="In Progress">⏳ In Progress</option>
        <option value="To Do">📋 To Do</option>
      </select>

      <!-- Type Filter -->
      <select id="filter-type" class="select-filter">
        <option value="">🏷️ All Item Types</option>
        <option value="Task">📋 Tasks</option>
        <option value="Deliverable">📦 Deliverables</option>
        <option value="Milestone">🔶 Milestones</option>
        <option value="Decision Point">🎯 Decision Points</option>
      </select>

      <button class="btn btn-light" id="btn-collapse-all" title="Collapse all parent tasks">▲ Collapse All</button>
      <button class="btn btn-light" id="btn-expand-all" title="Expand all parent tasks">▼ Expand All</button>
    </div>

    <div class="toolbar-right">
      <!-- Jump to Today -->
      <button class="btn btn-light" id="btn-jump-today" title="Center timeline on Today">📍 Go to Today</button>

      <!-- Zoom Level Selector -->
      <div class="zoom-group">
        <button class="zoom-btn" data-zoom="quarter">Quarter</button>
        <button class="zoom-btn active" data-zoom="month">Month</button>
        <button class="zoom-btn" data-zoom="week">Week</button>
      </div>
    </div>
  </div>

  <!-- KPI / Stats Summary Strip -->
  <div class="stats-strip">
    <div class="stat-pill"><span class="dot dot-complete"></span> Complete: <strong id="stat-complete">28</strong></div>
    <div class="stat-pill"><span class="dot dot-progress"></span> In Progress: <strong id="stat-progress">38</strong></div>
    <div class="stat-pill"><span class="dot dot-todo"></span> To Do: <strong id="stat-todo">86</strong></div>
    <div style="height:14px; width:1px; background:#cbd5e1; margin:0 4px;"></div>
    <div class="stat-pill">📦 Deliverables: <strong id="stat-deliverables">21</strong></div>
    <div class="stat-pill">🔶 Milestones: <strong id="stat-milestones">30</strong></div>
    <div class="stat-pill">🎯 Decision Points: <strong id="stat-decisions">7</strong></div>
    <div style="flex:1"></div>
    <div class="sync-status-badge" id="sync-status-indicator">● Checking disk sync...</div>
  </div>

  <!-- Main Work Area -->
  <div class="gantt-main-container">
    <!-- Left Pane: WBS Table -->
    <div class="tree-pane" id="tree-pane">
      <div class="pane-header">
        <div class="th-col col-wbs">WBS</div>
        <div class="th-col col-name">Task / Deliverable Name</div>
        <div class="th-col col-type">Type</div>
        <div class="th-col col-status">Status</div>
        <div class="th-col col-assignee">Assignee</div>
      </div>
      <div class="tree-scroll" id="tree-scroll">
        <table class="tree-table" id="tree-table">
          <tbody id="tree-tbody"></tbody>
        </table>
      </div>
    </div>

    <!-- Draggable Resizer Splitter -->
    <div class="resizer" id="resizer"></div>

    <!-- Right Pane: Gantt Timeline -->
    <div class="timeline-pane">
      <!-- Sticky Timeline Header -->
      <div class="timeline-header-container" id="timeline-header-container">
        <div class="timeline-header-months" id="header-months"></div>
        <div class="timeline-header-weeks" id="header-weeks"></div>
      </div>

      <!-- Scrollable Timeline Body -->
      <div class="timeline-body-scroll" id="timeline-body-scroll">
        <div class="timeline-canvas" id="timeline-canvas">
          <!-- Background Grid Columns -->
          <div id="grid-columns-container"></div>
          
          <!-- Today Vertical Line -->
          <div class="today-line" id="today-marker">
            <span class="today-tag" id="today-tag">TODAY</span>
          </div>

          <!-- Timeline Rows and Task Bars -->
          <div id="timeline-rows-container"></div>
        </div>
      </div>
    </div>
  </div>

  <!-- Bottom Status Bar -->
  <footer class="app-footer">
    <div style="display:flex; align-items:center; gap:12px;">
      <span id="footer-item-count">Showing 152 items</span>
      <span class="unscheduled-pill" id="unscheduled-pill" style="display:none;">⚠️ 0 Unscheduled Tasks</span>
    </div>
    <div style="display:flex; align-items:center; gap:12px;">
      <span>💡 Drag bars to reschedule • Click any task to edit details • Drag & drop .xlsx files here to load</span>
      <button class="btn btn-light" id="btn-reset-data" style="padding:2px 8px; font-size:11px;">Reset to Original Data</button>
    </div>
  </footer>

  <!-- Floating Tooltip -->
  <div id="gantt-tooltip"></div>

  <!-- Task Edit / Create Modal -->
  <div class="modal-overlay" id="task-modal">
    <div class="modal-card">
      <div class="modal-header">
        <h3 id="modal-title">Edit Task Details</h3>
        <button class="modal-close-btn" id="modal-close">&times;</button>
      </div>
      <div class="modal-body">
        <input type="hidden" id="edit-task-id" />

        <div class="form-row">
          <div class="form-group" style="flex:1;">
            <label>Project</label>
            <select class="form-control" id="edit-project">
              <option value="SMTP">SMTP (Statewide Multimodal)</option>
              <option value="TCRS">TCRS (Carbon Reduction)</option>
            </select>
          </div>
          <div class="form-group" style="flex:1;">
            <label>Parent Task</label>
            <select class="form-control" id="edit-parent-id">
              <option value="">(None - Top Level Task)</option>
            </select>
          </div>
        </div>

        <div class="form-group">
          <label>Task Name</label>
          <input type="text" class="form-control" id="edit-task-name" placeholder="Enter task title..." />
        </div>

        <div class="form-row">
          <div class="form-group">
            <label>Item Type</label>
            <select class="form-control" id="edit-task-type">
              <option value="Task">Task</option>
              <option value="Deliverable">Deliverable</option>
              <option value="Milestone">Milestone</option>
              <option value="Decision Point">Decision Point</option>
            </select>
          </div>
          <div class="form-group">
            <label>Status</label>
            <select class="form-control" id="edit-status">
              <option value="To Do">To Do</option>
              <option value="In Progress">In Progress</option>
              <option value="Complete">Complete</option>
            </select>
          </div>
          <div class="form-group">
            <label>Priority</label>
            <select class="form-control" id="edit-priority">
              <option value="Urgent">Urgent</option>
              <option value="High">High</option>
              <option value="Normal">Normal</option>
              <option value="Low">Low</option>
            </select>
          </div>
        </div>

        <div class="form-row">
          <div class="form-group">
            <label>Start Date</label>
            <input type="date" class="form-control" id="edit-start-date" />
          </div>
          <div class="form-group">
            <label>Due Date</label>
            <input type="date" class="form-control" id="edit-due-date" />
          </div>
        </div>

        <div class="form-row">
          <div class="form-group">
            <label>Assignee(s)</label>
            <input type="text" class="form-control" id="edit-assignees" placeholder="e.g. Leslie Langley, Paul Sharman" />
          </div>
          <div class="form-group">
            <label>Task Order / Tags</label>
            <input type="text" class="form-control" id="edit-tags" placeholder="e.g. Task Order 1" />
          </div>
        </div>

        <div class="form-group">
          <label>Task Content / Description</label>
          <textarea class="form-control" id="edit-description" placeholder="Notes, scope description, meeting details..."></textarea>
        </div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-outline" style="color:#ef4444; border-color:#fca5a5;" id="btn-delete-task">🗑️ Delete Task</button>
        <div style="display:flex; gap:10px;">
          <button class="btn btn-light" id="btn-modal-cancel">Cancel</button>
          <button class="btn btn-primary" id="btn-modal-save">💾 Save Changes</button>
        </div>
      </div>
    </div>
  </div>

  <!-- Unscheduled Tasks Modal Tray -->
  <div class="modal-overlay" id="unscheduled-modal">
    <div class="modal-card" style="width:680px;">
      <div class="modal-header">
        <h3>Unscheduled Tasks (Missing Dates)</h3>
        <button class="modal-close-btn" id="unscheduled-close">&times;</button>
      </div>
      <div class="modal-body" id="unscheduled-list" style="max-height:60vh;"></div>
      <div class="modal-footer">
        <div></div>
        <button class="btn btn-primary" id="unscheduled-ok">Done</button>
      </div>
    </div>
  </div>

  <!-- Toast Notification element -->
  <div class="toast" id="toast"></div>

  <!-- Drag & Drop Overlay -->
  <div class="dropzone-overlay" id="dropzone-overlay">
    <div>📥 Drop Excel (.xlsx) or CSV file here to update schedule!</div>
    <div style="font-size:14px; font-weight:400; opacity:0.9;">Both SMTP and TCRS tasks will automatically sync</div>
  </div>

  <script>
    // Initial Baseline Dataset Embedded
    const BASELINE_TASKS = {json_tasks_str};

    // State
    let tasks = [];
    let activeProject = "ALL"; // "ALL", "SMTP", "TCRS"
    let currentZoom = "month"; // "quarter", "month", "week"
    let collapsedSet = new Set(); // set of parent Task IDs currently collapsed
    let selectedTaskId = null;
    let serverSyncEnabled = false;
    let lastKnownServerMtime = 0;
    
    // Zoom scale configuration (pixels per day)
    const ZOOM_CONFIG = {{
      quarter: {{ pxPerDay: 2.2, label: "Quarter" }},
      month: {{ pxPerDay: 6.5, label: "Month" }},
      week: {{ pxPerDay: 18.0, label: "Week" }}
    }};

    // Project timeline bounds: May 1, 2026 to Dec 31, 2027
    const TIMELINE_START = new Date("2026-04-15T00:00:00");
    const TIMELINE_END = new Date("2028-01-15T00:00:00");
    const TODAY = new Date("2026-11-03T00:00:00"); // Matches project current context

    // Initialize application
    async function init() {{
      setupEventListeners();
      await checkServerSync();
      if (!serverSyncEnabled) {{
        loadSavedTasks();
      }}
      populateAssigneeDropdown();
      render();
      
      // Auto center on Today on first load
      setTimeout(() => {{
        jumpToToday();
      }}, 150);
    }}

    // Check if Python live sync server is running on localhost
    async function checkServerSync() {{
      const badge = document.getElementById("sync-status-indicator");
      try {{
        const res = await fetch("/api/status", {{ cache: "no-store" }});
        if (res.ok) {{
          const data = await res.json();
          serverSyncEnabled = true;
          lastKnownServerMtime = data.excel_mtime;
          badge.className = "sync-status-badge";
          badge.innerHTML = "🟢 Live Excel Sync Active (Disk/OneDrive)";
          
          // Fetch latest data from server
          await fetchServerTasks();
          // Poll for disk updates every 4 seconds
          setInterval(pollServerForChanges, 4000);
          return;
        }}
      }} catch (e) {{
        // Running as standalone file://
      }}
      serverSyncEnabled = false;
      badge.className = "sync-status-badge offline";
      badge.innerHTML = "💾 Standalone Mode (Browser / Drag & Drop)";
    }}

    async function fetchServerTasks() {{
      try {{
        const res = await fetch("/api/schedule", {{ cache: "no-store" }});
        if (res.ok) {{
          const data = await res.json();
          if (data.tasks && data.tasks.length > 0) {{
            tasks = data.tasks;
            lastKnownServerMtime = data.mtime;
            render();
            showToast("Synced with SMTP_TCRS_Schedule.xlsx on disk!");
          }}
        }}
      }} catch (err) {{
        console.error("Error fetching tasks from server:", err);
      }}
    }}

    async function pollServerForChanges() {{
      if (!serverSyncEnabled) return;
      try {{
        const res = await fetch("/api/status", {{ cache: "no-store" }});
        if (res.ok) {{
          const data = await res.json();
          if (data.excel_mtime && data.excel_mtime > lastKnownServerMtime) {{
            console.log("Excel file modified on disk, refreshing...");
            lastKnownServerMtime = data.excel_mtime;
            await fetchServerTasks();
          }}
        }}
      }} catch (e) {{}}
    }}

    // Load tasks from localStorage or baseline
    function loadSavedTasks() {{
      const saved = localStorage.getItem("smtp_tcrs_schedule_data");
      if (saved) {{
        try {{
          tasks = JSON.parse(saved);
          showToast("Loaded saved schedule from browser storage");
          return;
        }} catch(e) {{
          console.error("Error reading saved tasks", e);
        }}
      }}
      tasks = JSON.parse(JSON.stringify(BASELINE_TASKS));
    }}

    async function saveTasks() {{
      // Always save to browser localStorage
      localStorage.setItem("smtp_tcrs_schedule_data", JSON.stringify(tasks));
      
      // If server is connected, save directly to Excel on disk!
      if (serverSyncEnabled) {{
        try {{
          const res = await fetch("/api/schedule", {{
            method: "POST",
            headers: {{ "Content-Type": "application/json" }},
            body: JSON.stringify({{ tasks: tasks }})
          }});
          if (res.ok) {{
            const data = await res.json();
            lastKnownServerMtime = data.mtime;
            showToast("Saved to SMTP_TCRS_Schedule.xlsx on disk!");
            return;
          }}
        }} catch(e) {{
          console.warn("Server save failed, saved locally", e);
        }}
      }}

      showToast("Saved changes to browser storage!");
    }}

    // Setup UI event listeners
    function setupEventListeners() {{
      // Project Tabs
      document.querySelectorAll(".tab-btn").forEach(btn => {{
        btn.addEventListener("click", () => {{
          document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
          btn.classList.add("active");
          activeProject = btn.dataset.project;
          render();
        }});
      }});

      // Zoom Buttons
      document.querySelectorAll(".zoom-btn").forEach(btn => {{
        btn.addEventListener("click", () => {{
          document.querySelectorAll(".zoom-btn").forEach(b => b.classList.remove("active"));
          btn.classList.add("active");
          currentZoom = btn.dataset.zoom;
          renderTimeline();
        }});
      }});

      // Filters
      document.getElementById("search-input").addEventListener("input", debounce(render, 200));
      document.getElementById("filter-assignee").addEventListener("change", render);
      document.getElementById("filter-status").addEventListener("change", render);
      document.getElementById("filter-type").addEventListener("change", render);

      // Keyboard Shortcuts
      window.addEventListener("keydown", (e) => {{
        if (e.key === "/" && document.activeElement.tagName !== "INPUT" && document.activeElement.tagName !== "TEXTAREA") {{
          e.preventDefault();
          document.getElementById("search-input").focus();
        }} else if (e.key === "Escape") {{
          closeModal();
          closeUnscheduledModal();
        }}
      }});

      // Expand / Collapse All
      document.getElementById("btn-expand-all").addEventListener("click", () => {{
        collapsedSet.clear();
        render();
      }});
      document.getElementById("btn-collapse-all").addEventListener("click", () => {{
        tasks.forEach(t => {{
          if (hasChildren(t.id || t["Task ID"])) {{
            collapsedSet.add(t.id || t["Task ID"]);
          }}
        }});
        render();
      }});

      // Jump to Today
      document.getElementById("btn-jump-today").addEventListener("click", jumpToToday);

      // Resizer Dragging
      initResizer();

      // Synchronize Table and Timeline Vertical Scroll
      const treeScroll = document.getElementById("tree-scroll");
      const timelineScroll = document.getElementById("timeline-body-scroll");
      const headerScroll = document.getElementById("timeline-header-container");

      timelineScroll.addEventListener("scroll", () => {{
        treeScroll.scrollTop = timelineScroll.scrollTop;
        headerScroll.scrollLeft = timelineScroll.scrollLeft;
      }});

      treeScroll.addEventListener("scroll", () => {{
        timelineScroll.scrollTop = treeScroll.scrollTop;
        const paneHdr = document.querySelector(".pane-header");
        if (paneHdr) {{
          paneHdr.scrollLeft = treeScroll.scrollLeft;
        }}
      }});

      // Modal Controls
      document.getElementById("modal-close").addEventListener("click", closeModal);
      document.getElementById("btn-modal-cancel").addEventListener("click", closeModal);
      document.getElementById("btn-modal-save").addEventListener("click", saveTaskModal);
      document.getElementById("btn-delete-task").addEventListener("click", deleteTaskModal);
      document.getElementById("btn-add-task").addEventListener("click", () => openTaskModal(null));

      // Unscheduled modal
      document.getElementById("unscheduled-pill").addEventListener("click", openUnscheduledModal);
      document.getElementById("unscheduled-close").addEventListener("click", closeUnscheduledModal);
      document.getElementById("unscheduled-ok").addEventListener("click", closeUnscheduledModal);

      // Reset Data
      document.getElementById("btn-reset-data").addEventListener("click", () => {{
        if (confirm("Reset all schedule modifications back to the original baseline?")) {{
          localStorage.removeItem("smtp_tcrs_schedule_data");
          tasks = JSON.parse(JSON.stringify(BASELINE_TASKS));
          render();
          showToast("Reset to baseline schedule!");
        }}
      }});

      // File Import / Export
      document.getElementById("btn-import-file").addEventListener("click", () => {{
        document.getElementById("file-input").click();
      }});
      document.getElementById("file-input").addEventListener("change", handleFileSelect);

      document.getElementById("btn-export-excel").addEventListener("click", exportExcel);
      document.getElementById("btn-export-csv").addEventListener("click", exportCSV);
      document.getElementById("btn-print").addEventListener("click", () => window.print());

      // Drag and drop onto page
      setupDragAndDrop();
    }}

    function hasChildren(tid) {{
      return tasks.some(t => (t["Parent ID"] || t.parent_id) === tid);
    }}

    function getVisibleTasks() {{
      const searchTerm = document.getElementById("search-input").value.toLowerCase().trim();
      const assigneeFilter = document.getElementById("filter-assignee").value;
      const statusFilter = document.getElementById("filter-status").value;
      const typeFilter = document.getElementById("filter-type").value;

      // Filter by project
      let filtered = tasks.filter(t => {{
        if (activeProject !== "ALL" && t["Project"] !== activeProject) return false;
        if (assigneeFilter && !String(t["Assignees"] || "").includes(assigneeFilter)) return false;
        if (statusFilter && t["Status"] !== statusFilter) return false;
        if (typeFilter && t["Task Type"] !== typeFilter) return false;
        if (searchTerm) {{
          const txt = `${{t["Task Name"]}} ${{t["Description"]}} ${{t["Task ID"]}} ${{t["Assignees"]}}`.toLowerCase();
          if (!txt.includes(searchTerm)) return false;
        }}
        return true;
      }});

      // Filter out children of collapsed parents
      const visible = [];
      const hiddenIds = new Set();

      filtered.forEach(t => {{
        const pid = t["Parent ID"] || t.parent_id;
        if (pid && (hiddenIds.has(pid) || collapsedSet.has(pid))) {{
          hiddenIds.add(t["Task ID"] || t.id);
        }} else {{
          visible.push(t);
        }}
      }});

      return visible;
    }}

    // Render both Left Tree Table and Right Timeline
    function render() {{
      const visible = getVisibleTasks();
      updateStats();
      renderTreeTable(visible);
      renderTimeline(visible);
      updateFooter(visible.length);
    }}

    function renderTreeTable(visibleTasks) {{
      const tbody = document.getElementById("tree-tbody");
      tbody.innerHTML = "";

      visibleTasks.forEach(t => {{
        const tid = t["Task ID"] || t.id;
        const level = t["Level"] || 1;
        const isParent = hasChildren(tid);
        const isCollapsed = collapsedSet.has(tid);

        const tr = document.createElement("tr");
        tr.className = `tree-row ${{isParent ? "is-parent" : ""}} ${{selectedTaskId === tid ? "selected" : ""}}`;
        tr.dataset.taskId = tid;

        // Click to select & open modal
        tr.addEventListener("click", (e) => {{
          if (e.target.closest(".expander")) return;
          selectTask(tid);
        }});

        tr.addEventListener("dblclick", () => {{
          openTaskModal(tid);
        }});

        // WBS
        const tdWbs = document.createElement("td");
        tdWbs.className = "col-wbs";
        tdWbs.textContent = t["WBS"] || "";
        tr.appendChild(tdWbs);

        // Name + Expander + Indent (contained with set max width)
        const tdName = document.createElement("td");
        tdName.className = "col-name";

        const nameWrapper = document.createElement("div");
        nameWrapper.className = "col-name-wrapper";

        // Indent spacers
        for (let i = 1; i < level; i++) {{
          const sp = document.createElement("span");
          sp.className = "indent-spacer";
          nameWrapper.appendChild(sp);
        }}

        // Expander icon
        if (isParent) {{
          const exp = document.createElement("span");
          exp.className = `expander ${{isCollapsed ? "collapsed" : ""}}`;
          exp.textContent = "▼";
          exp.addEventListener("click", (e) => {{
            e.stopPropagation();
            if (collapsedSet.has(tid)) {{
              collapsedSet.delete(tid);
            }} else {{
              collapsedSet.add(tid);
            }}
            render();
          }});
          nameWrapper.appendChild(exp);
        }} else {{
          const sp = document.createElement("span");
          sp.className = "indent-spacer";
          nameWrapper.appendChild(sp);
        }}

        // Project badge dot
        const dot = document.createElement("span");
        dot.style.cssText = `width:6px; height:6px; border-radius:50%; background:${{t["Project"] === "SMTP" ? "var(--primary-smtp)" : "var(--primary-tcrs)"}}; flex-shrink:0;`;
        nameWrapper.appendChild(dot);

        // Name text (strictly contained with ellipsis and full title tooltip)
        const nameSpan = document.createElement("span");
        nameSpan.className = "col-name-text";
        nameSpan.textContent = t["Task Name"];
        nameSpan.title = t["Task Name"];
        nameWrapper.appendChild(nameSpan);

        tdName.appendChild(nameWrapper);
        tr.appendChild(tdName);

        // Type
        const tdType = document.createElement("td");
        tdType.className = "col-type";
        const typeBadge = document.createElement("span");
        const tType = t["Task Type"] || "Task";
        typeBadge.className = `badge badge-${{tType.toLowerCase().replace(" ", "")}}`;
        typeBadge.textContent = tType;
        tdType.appendChild(typeBadge);
        tr.appendChild(tdType);

        // Status
        const tdStatus = document.createElement("td");
        tdStatus.className = "col-status";
        const statusBadge = document.createElement("span");
        const status = t["Status"] || "To Do";
        statusBadge.className = `badge badge-${{status.toLowerCase().replace(" ", "")}}`;
        statusBadge.textContent = status;
        tdStatus.appendChild(statusBadge);
        tr.appendChild(tdStatus);

        // Assignee
        const tdAss = document.createElement("td");
        tdAss.className = "col-assignee";
        tdAss.textContent = t["Assignees"] || "";
        tdAss.title = t["Assignees"] || "";
        tr.appendChild(tdAss);

        tbody.appendChild(tr);
      }});
    }}

    function renderTimeline(visibleTasks) {{
      if (!visibleTasks) visibleTasks = getVisibleTasks();
      const pxPerDay = ZOOM_CONFIG[currentZoom].pxPerDay;
      const totalDays = Math.ceil((TIMELINE_END - TIMELINE_START) / (1000 * 60 * 60 * 24));
      const totalWidth = totalDays * pxPerDay;

      // Render Header
      renderTimelineHeader(totalWidth, pxPerDay);

      // Render Canvas and Grid
      const canvas = document.getElementById("timeline-canvas");
      canvas.style.width = `${{totalWidth}}px`;
      canvas.style.height = `${{visibleTasks.length * 38}}px`;

      renderGridColumns(totalWidth, pxPerDay);
      positionTodayMarker(pxPerDay);

      // Render Task Rows & Bars
      const rowsContainer = document.getElementById("timeline-rows-container");
      rowsContainer.innerHTML = "";

      visibleTasks.forEach((t, rowIdx) => {{
        const tid = t["Task ID"] || t.id;
        const rowDiv = document.createElement("div");
        rowDiv.className = `timeline-row ${{hasChildren(tid) ? "is-parent" : ""}} ${{selectedTaskId === tid ? "selected" : ""}}`;
        rowDiv.dataset.taskId = tid;

        rowDiv.addEventListener("click", () => selectTask(tid));
        rowDiv.addEventListener("dblclick", () => openTaskModal(tid));

        // Render Bar or Marker if dates exist
        const sDateStr = t["Start Date"];
        const dDateStr = t["Due Date"];

        if (sDateStr || dDateStr) {{
          let sDate = sDateStr ? new Date(sDateStr + "T00:00:00") : null;
          let dDate = dDateStr ? new Date(dDateStr + "T00:00:00") : null;

          if (!sDate && dDate) sDate = new Date(dDate.getTime() - 2 * 24 * 60 * 60 * 1000);
          if (!dDate && sDate) dDate = new Date(sDate.getTime() + 2 * 24 * 60 * 60 * 1000);

          const left = getXForDate(sDate, pxPerDay);
          const right = getXForDate(dDate, pxPerDay);
          const width = Math.max(right - left, 12);

          const itemType = t["Task Type"] || "Task";
          const isMilestone = itemType === "Milestone";
          const isDecision = itemType === "Decision Point";
          const isDeliverable = itemType === "Deliverable";
          const isParent = hasChildren(tid);

          if (isMilestone) {{
            // Render Milestone Diamond
            const diamond = document.createElement("div");
            diamond.className = "milestone-marker";
            diamond.style.left = `${{right - 9}}px`;
            attachTooltip(diamond, t);
            attachDragHandlers(diamond, t, "move-single", pxPerDay);
            rowDiv.appendChild(diamond);

            // Label outside
            const lbl = document.createElement("span");
            lbl.className = "bar-label-outside";
            lbl.style.left = `${{right + 12}}px`;
            lbl.textContent = `🔶 ${{t["Task Name"]}} (${{formatDateDisplay(dDate)}})`;
            rowDiv.appendChild(lbl);
          }} else if (isDecision) {{
            // Render Decision Marker
            const dec = document.createElement("div");
            dec.className = "decision-marker";
            dec.style.left = `${{right - 11}}px`;
            dec.innerHTML = "🎯";
            attachTooltip(dec, t);
            attachDragHandlers(dec, t, "move-single", pxPerDay);
            rowDiv.appendChild(dec);

            const lbl = document.createElement("span");
            lbl.className = "bar-label-outside";
            lbl.style.left = `${{right + 14}}px`;
            lbl.textContent = `🎯 ${{t["Task Name"]}}`;
            rowDiv.appendChild(lbl);
          }} else {{
            // Render Bar
            const bar = document.createElement("div");
            const projClass = t["Project"] === "SMTP" ? "task-bar-smtp" : "task-bar-tcrs";
            bar.className = `gantt-bar ${{isDeliverable ? "deliverable-bar" : projClass}} ${{isParent ? "parent-summary-bar" : ""}}`;
            bar.style.left = `${{left}}px`;
            bar.style.width = `${{width}}px`;

            if (!isParent) {{
              const icon = isDeliverable ? "📦 " : "";
              bar.innerHTML = `<span style="overflow:hidden; text-overflow:ellipsis;">${{icon}}${{t["Task Name"]}}</span>`;

              // Resize handles
              const hLeft = document.createElement("div");
              hLeft.className = "resize-handle left";
              const hRight = document.createElement("div");
              hRight.className = "resize-handle right";

              bar.appendChild(hLeft);
              bar.appendChild(hRight);

              attachDragHandlers(hLeft, t, "resize-left", pxPerDay);
              attachDragHandlers(hRight, t, "resize-right", pxPerDay);
              attachDragHandlers(bar, t, "move", pxPerDay);
            }}

            attachTooltip(bar, t);
            rowDiv.appendChild(bar);

            // If bar is short, add outside label
            if (width < 90) {{
              const lbl = document.createElement("span");
              lbl.className = "bar-label-outside";
              lbl.style.left = `${{left + width + 8}}px`;
              lbl.textContent = t["Task Name"];
              rowDiv.appendChild(lbl);
            }}
          }}
        }}

        rowsContainer.appendChild(rowDiv);
      }});
    }}

    function renderTimelineHeader(totalWidth, pxPerDay) {{
      const monthsDiv = document.getElementById("header-months");
      const weeksDiv = document.getElementById("header-weeks");
      monthsDiv.innerHTML = "";
      weeksDiv.innerHTML = "";
      monthsDiv.style.width = `${{totalWidth}}px`;
      weeksDiv.style.width = `${{totalWidth}}px`;

      let cur = new Date(TIMELINE_START);
      while (cur < TIMELINE_END) {{
        const yr = cur.getFullYear();
        const mo = cur.getMonth();
        const startMo = new Date(yr, mo, 1);
        const endMo = new Date(yr, mo + 1, 1);

        const x1 = getXForDate(startMo, pxPerDay);
        const x2 = getXForDate(endMo, pxPerDay);
        const width = x2 - x1;

        const cell = document.createElement("div");
        cell.className = "timeline-header-cell";
        cell.style.width = `${{width}}px`;
        cell.textContent = startMo.toLocaleDateString("en-US", {{ month: "short", year: "numeric" }});
        monthsDiv.appendChild(cell);

        // Sub-divisions (weeks)
        let wCur = new Date(startMo);
        while (wCur < endMo) {{
          const wEnd = new Date(wCur.getTime() + 7 * 24 * 60 * 60 * 1000);
          const wx1 = getXForDate(wCur, pxPerDay);
          const wx2 = Math.min(getXForDate(wEnd, pxPerDay), x2);
          const wWidth = wx2 - wx1;

          if (wWidth > 5) {{
            const wCell = document.createElement("div");
            wCell.className = "timeline-header-cell";
            wCell.style.width = `${{wWidth}}px`;
            wCell.textContent = currentZoom === "week" ? `Wk ${{getWeekNumber(wCur)}}` : `${{wCur.getDate()}}`;
            weeksDiv.appendChild(wCell);
          }}
          wCur = wEnd;
        }}

        cur = endMo;
      }}
    }}

    function renderGridColumns(totalWidth, pxPerDay) {{
      const grid = document.getElementById("grid-columns-container");
      grid.innerHTML = "";
      let cur = new Date(TIMELINE_START);
      while (cur < TIMELINE_END) {{
        const endMo = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
        const x = getXForDate(endMo, pxPerDay);
        const col = document.createElement("div");
        col.className = "grid-column month-boundary";
        col.style.left = `${{x}}px`;
        grid.appendChild(col);
        cur = endMo;
      }}
    }}

    function positionTodayMarker(pxPerDay) {{
      const todayX = getXForDate(TODAY, pxPerDay);
      const marker = document.getElementById("today-marker");
      marker.style.left = `${{todayX}}px`;
      document.getElementById("today-tag").textContent = `TODAY (${{TODAY.toLocaleDateString("en-US", {{ month: "short", day: "numeric" }})}})`;
    }}

    function jumpToToday() {{
      const pxPerDay = ZOOM_CONFIG[currentZoom].pxPerDay;
      const todayX = getXForDate(TODAY, pxPerDay);
      const scroll = document.getElementById("timeline-body-scroll");
      const clientWidth = scroll.clientWidth;
      scroll.scrollLeft = Math.max(0, todayX - clientWidth / 2);
    }}

    function getXForDate(d, pxPerDay) {{
      const days = (d - TIMELINE_START) / (1000 * 60 * 60 * 24);
      return Math.round(days * pxPerDay);
    }}

    function getDateForX(x, pxPerDay) {{
      const days = x / pxPerDay;
      const ms = TIMELINE_START.getTime() + days * 24 * 60 * 60 * 1000;
      return new Date(ms);
    }}

    function formatDateDisplay(d) {{
      if (!d) return "";
      return d.toLocaleDateString("en-US", {{ month: "short", day: "numeric", year: "numeric" }});
    }}

    function formatDateISO(d) {{
      if (!d) return "";
      const yr = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, "0");
      const day = String(d.getDate()).padStart(2, "0");
      return `${{yr}}-${{m}}-${{day}}`;
    }}

    function getWeekNumber(d) {{
      const date = new Date(d.getTime());
      date.setHours(0, 0, 0, 0);
      date.setDate(date.getDate() + 3 - (date.getDay() + 6) % 7);
      const week1 = new Date(date.getFullYear(), 0, 4);
      return 1 + Math.round(((date.getTime() - week1.getTime()) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
    }}

    // Drag to Reschedule / Resize Handlers
    function attachDragHandlers(el, task, mode, pxPerDay) {{
      el.addEventListener("mousedown", (e) => {{
        if (e.button !== 0) return; // only left click
        e.stopPropagation();

        const startX = e.clientX;
        const initialStartDate = task["Start Date"] ? new Date(task["Start Date"] + "T00:00:00") : new Date(TODAY);
        const initialDueDate = task["Due Date"] ? new Date(task["Due Date"] + "T00:00:00") : new Date(TODAY);
        let moved = false;

        function onMouseMove(ev) {{
          const dx = ev.clientX - startX;
          if (Math.abs(dx) > 3) moved = true;
          const daysDelta = Math.round(dx / pxPerDay);

          if (mode === "move") {{
            const newStart = new Date(initialStartDate.getTime() + daysDelta * 86400000);
            const newDue = new Date(initialDueDate.getTime() + daysDelta * 86400000);
            task["Start Date"] = formatDateISO(newStart);
            task["Due Date"] = formatDateISO(newDue);
          }} else if (mode === "move-single") {{
            const newDue = new Date(initialDueDate.getTime() + daysDelta * 86400000);
            task["Due Date"] = formatDateISO(newDue);
            task["Start Date"] = formatDateISO(newDue);
          }} else if (mode === "resize-left") {{
            const newStart = new Date(initialStartDate.getTime() + daysDelta * 86400000);
            if (newStart <= initialDueDate) {{
              task["Start Date"] = formatDateISO(newStart);
            }}
          }} else if (mode === "resize-right") {{
            const newDue = new Date(initialDueDate.getTime() + daysDelta * 86400000);
            if (newDue >= initialStartDate) {{
              task["Due Date"] = formatDateISO(newDue);
            }}
          }}

          renderTimeline();
        }}

        function onMouseUp() {{
          window.removeEventListener("mousemove", onMouseMove);
          window.removeEventListener("mouseup", onMouseUp);
          if (moved) {{
            saveTasks();
            render();
            showToast(`Updated ${{task["Task Name"]}} schedule!`);
          }}
        }}

        window.addEventListener("mousemove", onMouseMove);
        window.addEventListener("mouseup", onMouseUp);
      }});
    }}

    // Tooltip handling
    const tooltip = document.getElementById("gantt-tooltip");
    function attachTooltip(el, task) {{
      el.addEventListener("mouseenter", (e) => {{
        tooltip.innerHTML = `
          <h4>${{task["Task Name"]}}</h4>
          <div class="meta-line"><span>Project: <strong>${{task["Project"]}}</strong></span><span>WBS: <strong>${{task["WBS"] || "-"}}</strong></span></div>
          <div class="meta-line"><span>Type: <strong>${{task["Task Type"]}}</strong></span><span>Status: <strong>${{task["Status"]}}</strong></span></div>
          <div class="meta-line"><span>Start: ${{task["Start Date"] || "Not set"}}</span><span>Due: ${{task["Due Date"] || "Not set"}}</span></div>
          ${{task["Assignees"] ? `<div class="meta-line"><span>Assignee: <strong>${{task["Assignees"]}}</strong></span></div>` : ""}}
          ${{task["Tags"] ? `<div class="meta-line"><span>Order / Tags: <strong>${{task["Tags"]}}</strong></span></div>` : ""}}
          ${{task["Description"] ? `<div class="desc-line">${{task["Description"]}}</div>` : ""}}
        `;
        tooltip.style.display = "block";
        positionTooltip(e);
      }});

      el.addEventListener("mousemove", positionTooltip);
      el.addEventListener("mouseleave", () => {{
        tooltip.style.display = "none";
      }});
    }}

    function positionTooltip(e) {{
      const x = e.clientX + 14;
      const y = e.clientY + 14;
      tooltip.style.left = `${{Math.min(x, window.innerWidth - 360)}}px`;
      tooltip.style.top = `${{Math.min(y, window.innerHeight - 200)}}px`;
    }}

    function selectTask(tid) {{
      selectedTaskId = tid;
      document.querySelectorAll(".tree-row").forEach(r => r.classList.toggle("selected", r.dataset.taskId === tid));
      document.querySelectorAll(".timeline-row").forEach(r => r.classList.toggle("selected", r.dataset.taskId === tid));
    }}

    // Modal Operations
    function openTaskModal(tid) {{
      const modal = document.getElementById("task-modal");
      const parentSelect = document.getElementById("edit-parent-id");
      parentSelect.innerHTML = '<option value="">(None - Top Level Task)</option>';

      // Populate possible parents
      tasks.forEach(t => {{
        if (t["Task ID"] !== tid && (!t["Level"] || t["Level"] < 3)) {{
          const opt = document.createElement("option");
          opt.value = t["Task ID"];
          opt.textContent = `[${{t["Project"]}}] ${{t["Task Name"]}}`;
          parentSelect.appendChild(opt);
        }}
      }});

      if (tid) {{
        // Edit existing
        const task = tasks.find(t => (t["Task ID"] || t.id) === tid);
        if (!task) return;
        document.getElementById("modal-title").textContent = "Edit Task Details";
        document.getElementById("edit-task-id").value = tid;
        document.getElementById("edit-project").value = task["Project"] || "SMTP";
        document.getElementById("edit-parent-id").value = task["Parent ID"] || "";
        document.getElementById("edit-task-name").value = task["Task Name"] || "";
        document.getElementById("edit-task-type").value = task["Task Type"] || "Task";
        document.getElementById("edit-status").value = task["Status"] || "To Do";
        document.getElementById("edit-priority").value = task["Priority"] || "Normal";
        document.getElementById("edit-start-date").value = task["Start Date"] || "";
        document.getElementById("edit-due-date").value = task["Due Date"] || "";
        document.getElementById("edit-assignees").value = task["Assignees"] || "";
        document.getElementById("edit-tags").value = task["Tags"] || "";
        document.getElementById("edit-description").value = task["Description"] || "";
        document.getElementById("btn-delete-task").style.display = "inline-flex";
      }} else {{
        // Create new
        document.getElementById("modal-title").textContent = "Add New Schedule Item";
        document.getElementById("edit-task-id").value = "";
        document.getElementById("edit-project").value = activeProject === "ALL" ? "SMTP" : activeProject;
        document.getElementById("edit-parent-id").value = "";
        document.getElementById("edit-task-name").value = "";
        document.getElementById("edit-task-type").value = "Task";
        document.getElementById("edit-status").value = "To Do";
        document.getElementById("edit-priority").value = "Normal";
        document.getElementById("edit-start-date").value = formatDateISO(TODAY);
        document.getElementById("edit-due-date").value = formatDateISO(new Date(TODAY.getTime() + 14 * 86400000));
        document.getElementById("edit-assignees").value = "";
        document.getElementById("edit-tags").value = "";
        document.getElementById("edit-description").value = "";
        document.getElementById("btn-delete-task").style.display = "none";
      }}

      modal.classList.add("active");
    }}

    function closeModal() {{
      document.getElementById("task-modal").classList.remove("active");
    }}

    function saveTaskModal() {{
      const tid = document.getElementById("edit-task-id").value;
      const name = document.getElementById("edit-task-name").value.trim();
      if (!name) {{
        alert("Please enter a task name.");
        return;
      }}

      if (tid) {{
        // Update
        const task = tasks.find(t => (t["Task ID"] || t.id) === tid);
        if (task) {{
          task["Project"] = document.getElementById("edit-project").value;
          task["Parent ID"] = document.getElementById("edit-parent-id").value;
          task["Task Name"] = name;
          task["Task Type"] = document.getElementById("edit-task-type").value;
          task["Status"] = document.getElementById("edit-status").value;
          task["Priority"] = document.getElementById("edit-priority").value;
          task["Start Date"] = document.getElementById("edit-start-date").value;
          task["Due Date"] = document.getElementById("edit-due-date").value;
          task["Assignees"] = document.getElementById("edit-assignees").value;
          task["Tags"] = document.getElementById("edit-tags").value;
          task["Description"] = document.getElementById("edit-description").value;
        }}
      }} else {{
        // New task
        const newId = "task_" + Date.now().toString(36);
        const newTask = {{
          "Project": document.getElementById("edit-project").value,
          "Task ID": newId,
          "Parent ID": document.getElementById("edit-parent-id").value,
          "Task Name": name,
          "Task Type": document.getElementById("edit-task-type").value,
          "Status": document.getElementById("edit-status").value,
          "Priority": document.getElementById("edit-priority").value,
          "Start Date": document.getElementById("edit-start-date").value,
          "Due Date": document.getElementById("edit-due-date").value,
          "Assignees": document.getElementById("edit-assignees").value,
          "Tags": document.getElementById("edit-tags").value,
          "Description": document.getElementById("edit-description").value,
          "Level": 2,
          "WBS": "New"
        }};
        tasks.push(newTask);
      }}

      saveTasks();
      closeModal();
      render();
    }}

    function deleteTaskModal() {{
      const tid = document.getElementById("edit-task-id").value;
      if (!tid) return;
      if (confirm("Are you sure you want to delete this task?")) {{
        tasks = tasks.filter(t => (t["Task ID"] || t.id) !== tid);
        saveTasks();
        closeModal();
        render();
        showToast("Task deleted.");
      }}
    }}

    // Unscheduled modal
    function openUnscheduledModal() {{
      const modal = document.getElementById("unscheduled-modal");
      const list = document.getElementById("unscheduled-list");
      list.innerHTML = "";

      const unscheduled = tasks.filter(t => !t["Start Date"] && !t["Due Date"]);
      if (unscheduled.length === 0) {{
        list.innerHTML = "<p>All tasks currently have scheduled dates!</p>";
      }} else {{
        unscheduled.forEach(t => {{
          const item = document.createElement("div");
          item.style.cssText = "display:flex; align-items:center; justify-content:space-between; padding:8px 12px; border-bottom:1px solid #e2e8f0; font-size:12.5px;";
          item.innerHTML = `
            <div>
              <span class="badge badge-${{t["Project"].toLowerCase()}}" style="margin-right:6px;">${{t["Project"]}}</span>
              <strong>${{t["Task Name"]}}</strong>
              <div style="font-size:11px; color:#64748b;">${{t["Task Type"]}} • ${{t["Status"]}}</div>
            </div>
            <button class="btn btn-primary" style="padding:4px 8px; font-size:11.5px;">Schedule Dates</button>
          `;
          item.querySelector("button").addEventListener("click", () => {{
            closeUnscheduledModal();
            openTaskModal(t["Task ID"] || t.id);
          }});
          list.appendChild(item);
        }});
      }}

      modal.classList.add("active");
    }}

    function closeUnscheduledModal() {{
      document.getElementById("unscheduled-modal").classList.remove("active");
    }}

    // Update stats and badge counts
    function updateStats() {{
      const activeTasks = activeProject === "ALL" ? tasks : tasks.filter(t => t["Project"] === activeProject);
      const complete = activeTasks.filter(t => t["Status"] === "Complete").length;
      const progress = activeTasks.filter(t => t["Status"] === "In Progress").length;
      const todo = activeTasks.filter(t => t["Status"] === "To Do").length;

      document.getElementById("stat-complete").textContent = complete;
      document.getElementById("stat-progress").textContent = progress;
      document.getElementById("stat-todo").textContent = todo;

      document.getElementById("stat-deliverables").textContent = activeTasks.filter(t => t["Task Type"] === "Deliverable").length;
      document.getElementById("stat-milestones").textContent = activeTasks.filter(t => t["Task Type"] === "Milestone").length;
      document.getElementById("stat-decisions").textContent = activeTasks.filter(t => t["Task Type"] === "Decision Point").length;

      const unscheduled = activeTasks.filter(t => !t["Start Date"] && !t["Due Date"]).length;
      const pill = document.getElementById("unscheduled-pill");
      if (unscheduled > 0) {{
        pill.style.display = "inline-flex";
        pill.textContent = `⚠️ ${{unscheduled}} Unscheduled Tasks`;
      }} else {{
        pill.style.display = "none";
      }}
    }}

    function updateFooter(visibleCount) {{
      document.getElementById("footer-item-count").textContent = `Showing ${{visibleCount}} of ${{tasks.length}} items`;
    }}

    function populateAssigneeDropdown() {{
      const assignees = new Set();
      tasks.forEach(t => {{
        if (t["Assignees"]) {{
          t["Assignees"].split(",").forEach(a => {{
            const clean = a.trim();
            if (clean) assignees.add(clean);
          }});
        }}
      }});

      const select = document.getElementById("filter-assignee");
      select.innerHTML = '<option value="">👤 All Assignees</option>';
      Array.from(assignees).sort().forEach(a => {{
        const opt = document.createElement("option");
        opt.value = a;
        opt.textContent = a;
        select.appendChild(opt);
      }});
    }}

    // Splitter Resizer
    function initResizer() {{
      const resizer = document.getElementById("resizer");
      const treePane = document.getElementById("tree-pane");
      let startX, startWidth;

      resizer.addEventListener("mousedown", (e) => {{
        startX = e.clientX;
        startWidth = treePane.offsetWidth;
        resizer.classList.add("dragging");

        function onMouseMove(ev) {{
          const newWidth = Math.max(300, Math.min(800, startWidth + (ev.clientX - startX)));
          treePane.style.width = `${{newWidth}}px`;
        }}

        function onMouseUp() {{
          resizer.classList.remove("dragging");
          window.removeEventListener("mousemove", onMouseMove);
          window.removeEventListener("mouseup", onMouseUp);
        }}

        window.addEventListener("mousemove", onMouseMove);
        window.addEventListener("mouseup", onMouseUp);
      }});
    }}

    // Drag and Drop of Files
    function setupDragAndDrop() {{
      const overlay = document.getElementById("dropzone-overlay");
      window.addEventListener("dragover", (e) => {{
        e.preventDefault();
        overlay.classList.add("active");
      }});
      window.addEventListener("dragleave", (e) => {{
        if (e.clientX === 0 || e.clientY === 0) overlay.classList.remove("active");
      }});
      window.addEventListener("drop", (e) => {{
        e.preventDefault();
        overlay.classList.remove("active");
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {{
          processUploadedFile(e.dataTransfer.files[0]);
        }}
      }});
    }}

    function handleFileSelect(e) {{
      if (e.target.files && e.target.files[0]) {{
        processUploadedFile(e.target.files[0]);
      }}
    }}

    function processUploadedFile(file) {{
      const reader = new FileReader();
      const ext = file.name.split(".").pop().toLowerCase();

      if (ext === "xlsx" || ext === "xls") {{
        reader.onload = function(e) {{
          try {{
            if (typeof XLSX === "undefined") {{
              alert("SheetJS library is loading. Please check network connection or use CSV export.");
              return;
            }}
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {{ type: "array" }});
            // Prefer 'All Tasks (Master)' sheet or first sheet
            const sheetName = workbook.SheetNames.includes("All Tasks (Master)") ? "All Tasks (Master)" : workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet);
            
            // Map rows to task structure
            const newTasks = rows.map((r, i) => ({{
              "Project": r["Project"] || "SMTP",
              "WBS": r["WBS"] || `${{i+1}}`,
              "Task Name": String(r["Task Name"] || "").trim(),
              "Task Type": r["Task Type"] || "Task",
              "Status": r["Status"] || "To Do",
              "Start Date": r["Start Date"] ? String(r["Start Date"]).split("T")[0] : "",
              "Due Date": r["Due Date"] ? String(r["Due Date"]).split("T")[0] : "",
              "Assignees": r["Assignees"] || "",
              "Priority": r["Priority"] || "Normal",
              "Tags": r["Task Order"] || r["Tags"] || "",
              "Parent Task": r["Parent Task"] || "",
              "Task ID": r["Task ID"] || `task_${{i}}`,
              "Parent ID": r["Parent ID"] || "",
              "Description": r["Description"] || "",
              "Level": String(r["WBS"] || "").split(".").length
            }}));

            if (newTasks.length > 0) {{
              tasks = newTasks;
              saveTasks();
              populateAssigneeDropdown();
              render();
              showToast(`Loaded ${{newTasks.length}} tasks from ${{file.name}}!`);
            }}
          }} catch(err) {{
            console.error("Excel load error", err);
            alert("Error parsing Excel file: " + err.message);
          }}
        }};
        reader.readAsArrayBuffer(file);
      }} else if (ext === "csv") {{
        reader.onload = function(e) {{
          try {{
            parseCSVData(e.target.result);
            showToast(`Loaded tasks from ${{file.name}}!`);
          }} catch(err) {{
            alert("Error reading CSV: " + err.message);
          }}
        }};
        reader.readAsText(file);
      }}
    }}

    function parseCSVData(csvText) {{
      const lines = csvText.split(/\\r?\\n/).filter(l => l.trim());
      if (lines.length < 2) return;
      // Simple CSV parser
      const headers = lines[0].split(",").map(h => h.trim().replace(/^"|"$/g, ""));
      const newTasks = [];

      for (let i = 1; i < lines.length; i++) {{
        const cols = lines[i].match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g) || lines[i].split(",");
        const row = {{}};
        headers.forEach((h, idx) => {{
          row[h] = cols[idx] ? cols[idx].replace(/^"|"$/g, "").trim() : "";
        }});

        newTasks.push({{
          "Project": row["List Name"] && row["List Name"].includes("TCRS") ? "TCRS" : (row["Project"] || "SMTP"),
          "Task ID": row["Task ID"] || `task_${{i}}`,
          "Parent ID": row["Parent ID"] || "",
          "Task Name": row["Task Name"] || "",
          "Task Type": row["Task Type"] || "Task",
          "Status": row["Status"] || "To Do",
          "Start Date": row["Start Date"] || "",
          "Due Date": row["Due Date"] || "",
          "Assignees": row["Assignees"] || "",
          "Priority": row["Priority"] || "Normal",
          "Tags": row["Tags"] || "",
          "Description": row["Task Content"] || row["Description"] || "",
          "WBS": row["WBS"] || `${{i}}`,
          "Level": 1
        }});
      }}

      if (newTasks.length > 0) {{
        tasks = newTasks;
        saveTasks();
        populateAssigneeDropdown();
        render();
      }}
    }}

    // Export to Excel (.xlsx) using SheetJS
    function exportExcel() {{
      if (typeof XLSX === "undefined") {{
        alert("SheetJS is not loaded yet. Exporting as CSV instead.");
        exportCSV();
        return;
      }}

      const wb = XLSX.utils.book_new();

      function formatSheetData(taskList) {{
        return taskList.map(t => ({{
          "Project": t["Project"],
          "WBS": t["WBS"] || "",
          "Task Name": t["Task Name"],
          "Task Type": t["Task Type"],
          "Status": t["Status"],
          "Start Date": t["Start Date"],
          "Due Date": t["Due Date"],
          "Assignees": t["Assignees"],
          "Priority": t["Priority"],
          "Task Order": t["Tags"],
          "Parent Task": t["Parent Task"] || "",
          "Task ID": t["Task ID"],
          "Parent ID": t["Parent ID"],
          "Description": t["Description"]
        }}));
      }}

      // 1. Master sheet
      const wsAll = XLSX.utils.json_to_sheet(formatSheetData(tasks));
      XLSX.utils.book_append_sheet(wb, wsAll, "All Tasks (Master)");

      // 2. SMTP sheet
      const smtpList = tasks.filter(t => t["Project"] === "SMTP");
      const wsSmtp = XLSX.utils.json_to_sheet(formatSheetData(smtpList));
      XLSX.utils.book_append_sheet(wb, wsSmtp, "SMTP Tasks");

      // 3. TCRS sheet
      const tcrsList = tasks.filter(t => t["Project"] === "TCRS");
      const wsTcrs = XLSX.utils.json_to_sheet(formatSheetData(tcrsList));
      XLSX.utils.book_append_sheet(wb, wsTcrs, "TCRS Tasks");

      XLSX.writeFile(wb, "SMTP_TCRS_Schedule_Updated.xlsx");
      showToast("Downloaded SMTP_TCRS_Schedule_Updated.xlsx!");
    }}

    // Export to CSV
    function exportCSV() {{
      const headers = ["Project", "WBS", "Task Name", "Task Type", "Status", "Start Date", "Due Date", "Assignees", "Priority", "Tags", "Parent Task", "Task ID", "Parent ID", "Description"];
      let csv = headers.join(",") + "\\n";

      tasks.forEach(t => {{
        const row = [
          t["Project"],
          `"${{t["WBS"] || ""}}"`,
          `"${{(t["Task Name"] || "").replace(/"/g, '""')}}"`,
          t["Task Type"],
          t["Status"],
          t["Start Date"] || "",
          t["Due Date"] || "",
          `"${{(t["Assignees"] || "").replace(/"/g, '""')}}"`,
          t["Priority"] || "",
          `"${{(t["Tags"] || "").replace(/"/g, '""')}}"`,
          `"${{(t["Parent Task"] || "").replace(/"/g, '""')}}"`,
          t["Task ID"] || "",
          t["Parent ID"] || "",
          `"${{(t["Description"] || "").replace(/"/g, '""')}}"`
        ];
        csv += row.join(",") + "\\n";
      }});

      const blob = new Blob([csv], {{ type: "text/csv;charset=utf-8;" }});
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "SMTP_TCRS_Schedule_Updated.csv";
      link.click();
      showToast("Downloaded SMTP_TCRS_Schedule_Updated.csv!");
    }}

    // Toast Notification helper
    function showToast(msg) {{
      const t = document.getElementById("toast");
      t.textContent = msg;
      t.classList.add("show");
      setTimeout(() => {{
        t.classList.remove("show");
      }}, 3000);
    }}

    function debounce(func, wait) {{
      let timeout;
      return function(...args) {{
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), wait);
      }};
    }}

    // Run on startup
    window.addEventListener("DOMContentLoaded", init);
  </script>
</body>
</html>
'''

with open(html_path, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"Generated {html_path} successfully (size: {os.path.getsize(html_path)} bytes)")
