import os
import sys
import json
import time
import datetime
import urllib.parse
from http.server import HTTPServer, SimpleHTTPRequestHandler
import openpyxl

EXCEL_FILE = "SMTP_TCRS_Schedule.xlsx"
HTML_FILE = "SMTP_TCRS_Interactive_Gantt.html"
PORT = 8080

def read_excel_tasks():
    if not os.path.exists(EXCEL_FILE):
        return []
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    sheet_name = "All Tasks (Master)" if "All Tasks (Master)" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]
    
    headers = [cell.value for cell in ws[1]]
    tasks = []
    
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        row_dict = dict(zip(headers, row))
        
        # Clean dates
        def parse_d(val):
            if val is None:
                return ""
            if isinstance(val, (datetime.datetime, datetime.date)):
                return val.strftime("%Y-%m-%d")
            s = str(val).strip()
            if " " in s:
                s = s.split(" ")[0]
            return s

        t_name = str(row_dict.get("Task Name") or "").strip()
        wbs = str(row_dict.get("WBS") or "").strip()
        level = len(wbs.split(".")) if wbs else 1

        task = {
            "Project": str(row_dict.get("Project") or "SMTP").strip(),
            "WBS": wbs,
            "Task Name": t_name,
            "Task Type": str(row_dict.get("Task Type") or "Task").strip(),
            "Status": str(row_dict.get("Status") or "To Do").strip(),
            "Start Date": parse_d(row_dict.get("Start Date")),
            "Due Date": parse_d(row_dict.get("Due Date")),
            "Assignees": str(row_dict.get("Assignees") or "").strip(),
            "Priority": str(row_dict.get("Priority") or "Normal").strip(),
            "Tags": str(row_dict.get("Task Order") or row_dict.get("Tags") or "").strip(),
            "Parent Task": str(row_dict.get("Parent Task") or "").strip(),
            "Task ID": str(row_dict.get("Task ID") or "").strip(),
            "Parent ID": str(row_dict.get("Parent ID") or "").strip(),
            "Description": str(row_dict.get("Description") or "").strip(),
            "Level": level
        }
        tasks.append(task)
    return tasks

def save_excel_tasks(new_tasks):
    # Call the existing generation script logic to update SMTP_TCRS_Schedule.xlsx cleanly
    from create_excel_schedule import add_schedule_sheet, wb, NAVY_FILL, SMTP_FILL, TCRS_FILL
    import create_excel_schedule
    # Save JSON data first
    with open("schedule_data.json", "w", encoding="utf-8") as f:
        json.dump(new_tasks, f, indent=2)
    # Re-run generator
    os.system(f'python "{os.path.join(os.path.dirname(__file__), "create_excel_schedule.py")}"')
    return True

class GanttSyncHandler(SimpleHTTPRequestHandler):
    def end_headers(self):
        # Enable CORS
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/status":
            mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
            res = {"status": "ok", "excel_mtime": mtime, "port": PORT}
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps(res).encode("utf-8"))
            return
        elif parsed.path == "/api/schedule":
            tasks = read_excel_tasks()
            mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
            res = {"tasks": tasks, "mtime": mtime}
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps(res).encode("utf-8"))
            return
        elif parsed.path in ("/", ""):
            self.path = "/" + HTML_FILE
            return super().do_GET()
        else:
            return super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/schedule":
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length)
            try:
                data = json.loads(post_data.decode('utf-8'))
                tasks_list = data.get("tasks", [])
                if tasks_list:
                    # Update JSON and Excel
                    with open("schedule_data.json", "w", encoding="utf-8") as f:
                        json.dump(tasks_list, f, indent=2)
                    
                    import subprocess
                    script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "create_excel_schedule.py")
                    subprocess.run([sys.executable, script_path], check=True)
                    
                    mtime = os.path.getmtime(EXCEL_FILE)
                    res = {"success": True, "count": len(tasks_list), "mtime": mtime}
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps(res).encode("utf-8"))
                    return
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode("utf-8"))
                return

def run_server():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    server_address = ('', PORT)
    httpd = HTTPServer(server_address, GanttSyncHandler)
    print(f"============================================================")
    print(f" WSDOT SMTP & TCRS Gantt Chart Live Sync Server Running")
    print(f" URL: http://localhost:{PORT}")
    print(f" Excel Target: {os.path.abspath(EXCEL_FILE)}")
    print(f"============================================================")
    try:
        import webbrowser
        webbrowser.open(f"http://localhost:{PORT}")
    except Exception:
        pass
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server.")
        httpd.server_close()

if __name__ == "__main__":
    run_server()

