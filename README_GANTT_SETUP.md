# WSDOT SMTP & TCRS Interactive Gantt Chart & Schedule

This package recreates and modernizes the interactive Gantt chart for the **WSDOT Statewide Multimodal Transportation Plan (SMTP)** and **Transportation Carbon Reduction Strategy (TCRS)**.

It provides a seamless **dual-update workflow**: team members can update project dates and statuses either directly in the **shared Excel workbook** or via the **rich interactive Gantt chart interface**.

---

## 📁 Key Files Created

| File | Purpose |
| :--- | :--- |
| **`SMTP_TCRS_Schedule.xlsx`** | The master formatted Excel workbook for team collaboration on OneDrive. Contains **Dashboard & Summary**, **All Tasks (Master)**, **SMTP Tasks**, and **TCRS Tasks** with dropdowns and conditional formatting. |
| **`SMTP_TCRS_Interactive_Gantt.html`** | The standalone interactive web-based Gantt chart. Zero-install, opens in any web browser (Chrome, Edge). Supports drag-and-drop rescheduling, filtering, editing, and Excel import/export. |
| **`Start_Gantt_Chart.bat`** | **1-click launcher** for Windows. Double-click this to automatically start the two-way live disk sync server and open the Gantt chart in your browser. |
| **`sync_server.py`** | Lightweight background server that monitors `SMTP_TCRS_Schedule.xlsx` and keeps the web interface and Excel file in real-time two-way synchronization. |
| **`schedule_data.json`** | Cleaned and structured JSON database of all 171 tasks, deliverables, milestones, decision points, and regional outreach events. |

---

## 🚀 How to Use & Update the Schedule

You can choose either of two ways to update the schedule:

### Option A: Update via the Shared Excel File (`.xlsx`)
1. Open [`SMTP_TCRS_Schedule.xlsx`](file:///c:/Users/leslie.langley/OneDrive%20-%20Transpo%20Group/WSDOT%20SMTP%20&%20CRS%20-%20General/Task%206%20-%20Carbon%20Reduction%20Strategy%20Update/Task%206.1%20-%20Management%20&%20Administration/SMTP_TCRS_Schedule.xlsx) in Microsoft Excel or Excel Online.
2. Edit start/due dates (`YYYY-MM-DD`), assignees, status (`Complete`, `In Progress`, `To Do`), or item types using the built-in dropdown menus.
3. Save the file. Because it sits in OneDrive, team members receive updates automatically.
4. If you have the interactive Gantt chart open, you can simply **drag and drop** the updated `.xlsx` file onto the browser window, or click **"📂 Load File"** to refresh it immediately.

### Option B: Update via the Interactive Gantt Chart
1. **Double-click [`Start_Gantt_Chart.bat`](file:///c:/Users/leslie.langley/OneDrive%20-%20Transpo%20Group/WSDOT%20SMTP%20&%20CRS%20-%20General/Task%206%20-%20Carbon%20Reduction%20Strategy%20Update/Task%206.1%20-%20Management%20&%20Administration/Start_Gantt_Chart.bat)** (or open [`SMTP_TCRS_Interactive_Gantt.html`](file:///c:/Users/leslie.langley/OneDrive%20-%20Transpo%20Group/WSDOT%20SMTP%20&%20CRS%20-%20General/Task%206%20-%20Carbon%20Reduction%20Strategy%20Update/Task%206.1%20-%20Management%20&%20Administration/SMTP_TCRS_Interactive_Gantt.html) in your browser).
2. **Interactive Controls**:
   - **Switch Projects**: Click `All Projects`, `SMTP`, or `TCRS` in the top bar.
   - **Reschedule**: Click and drag any timeline bar left/right to adjust dates, or drag the bar edges to adjust duration.
   - **Edit / Add Tasks**: Double-click any row or click the **"➕ Add Task"** button to open the Task Modal and modify title, dates, priority, notes, or assignees.
   - **Filter & Search**: Type `/` to search by keywords, filter by Assignee (`Leslie Langley`, `Paul Sharman`, etc.), Status, or Item Type.
   - **Zoom**: Toggle between **Month**, **Week**, or **Quarter** views.
   - **Today Marker**: Click **"📍 Go to Today"** to instantly center on the current progress line.
3. **Saving & Exporting**:
   - When running via `Start_Gantt_Chart.bat`, changes save directly to `SMTP_TCRS_Schedule.xlsx` on disk!
   - In standalone browser mode, click **"📥 Export Excel (.xlsx)"** anytime to download an updated, formatted Excel workbook.

---

## 🎨 Visual Item Types
- 📋 **Task**: Standard horizontal bar colored by project (Azure blue for SMTP, Emerald green for TCRS).
- 📦 **Deliverable**: Distinct outlined card bar with a package icon.
- 🔶 **Milestone**: Amber diamond marker at the target completion date.
- 🎯 **Decision Point**: Rose badge highlighting critical WSDOT decision milestones.

