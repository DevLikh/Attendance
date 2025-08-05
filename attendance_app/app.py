import sys
import os
import json
from datetime import datetime, timedelta
import win32com.client
import atexit
import jpholiday

from PySide6.QtWidgets import QApplication, QMainWindow
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtCore import QObject, Slot, Signal

# --- Constants ---
if getattr(sys, 'frozen', False):
    app_path = os.path.dirname(sys.executable)
else:
    app_path = os.path.dirname(os.path.abspath(__file__))

EXCEL_FILE_PATH = os.path.join(app_path, "attendance_data.xlsx")

# --- Helper Functions ---
def round_up_time(dt):
    discard = timedelta(minutes=dt.minute % 15, seconds=dt.second, microseconds=dt.microsecond)
    dt -= discard
    if discard > timedelta(0):
        dt += timedelta(minutes=15)
    return dt

def round_down_time(dt):
    discard = timedelta(minutes=dt.minute % 15, seconds=dt.second, microseconds=dt.microsecond)
    dt -= discard
    return dt

# --- Excel Management ---
class ExcelManager:
    def __init__(self, filepath):
        self.filepath = filepath
        self.excel_app = None
        self.workbook = None
        try:
            self.excel_app = win32com.client.Dispatch("Excel.Application")
            self.excel_app.Visible = False
            self.excel_app.DisplayAlerts = False # Suppress alerts
        except Exception as e:
            print(f"Excelの起動エラー: {e}")
            return

        if os.path.exists(self.filepath):
            try:
                self.workbook = self.excel_app.Workbooks.Open(self.filepath)
            except Exception as e:
                print(f"Excelファイルの読み込みエラー: {e}")
                self._create_new_workbook()
        else:
            self._create_new_workbook()
        
        atexit.register(self.shutdown)

    def _create_new_workbook(self):
        self.workbook = self.excel_app.Workbooks.Add()
        try:
            self.workbook.Worksheets(1).Name = "Attendance"
            self.workbook.Worksheets.Add().Name = "Tasks"
            self.workbook.Worksheets.Add().Name = "Announcements"
            self.workbook.SaveAs(self.filepath)
            self.workbook.Close()
            self.workbook = self.excel_app.Workbooks.Open(self.filepath)
        except Exception as e:
            print(f"新規Excelファイルの作成エラー: {e}")

    def load_all_data(self):
        all_data = {"attendance": {}, "tasks": {}, "announcements": {}}
        print("--- Excelデータ読み込み開始 ---")
        try:
            # Load Attendance
            ws = self.workbook.Worksheets("Attendance")
            print(f"Attendanceシートの使用範囲行数: {ws.UsedRange.Rows.Count}")
            if ws.UsedRange.Rows.Count > 1: # Check if there's data beyond headers
                for row in range(2, ws.UsedRange.Rows.Count + 1):
                    try:
                        raw_employee_id = ws.Cells(row, 1).Value
                        employee_id = str(int(raw_employee_id)) if isinstance(raw_employee_id, float) else str(raw_employee_id).strip()
                        
                        raw_date_val = ws.Cells(row, 2).Value
                        date_str = raw_date_val.strftime('%Y-%m-%d') if isinstance(raw_date_val, datetime) else str(raw_date_val).strip()
                        
                        work_type = str(ws.Cells(row, 3).Value or "").strip()
                        check_in = str(ws.Cells(row, 4).Value or "").strip().lstrip("'") # Remove leading '
                        check_out = str(ws.Cells(row, 5).Value or "").strip().lstrip("'") # Remove leading '
                        rest_time = str(ws.Cells(row, 6).Value or "01:00").strip().lstrip("'") # Remove leading '
                        subtasks_json_raw = str(ws.Cells(row, 7).Value or '[]').strip()
                        
                        subtasks = []
                        try:
                            subtasks = json.loads(subtasks_json_raw)
                        except json.JSONDecodeError as json_e:
                            print(f"サブタスクのJSONデコードエラー (行 {row}): {json_e} - 生データ: {subtasks_json_raw}")
                            subtasks = [] # Default to empty list on error

                        if employee_id not in all_data["attendance"]:
                            all_data["attendance"][employee_id] = {}
                        
                        all_data["attendance"][employee_id][date_str] = {
                            'work_type': work_type,
                            'check_in': check_in,
                            'check_out': check_out,
                            'rest_time': rest_time,
                            'subtasks': subtasks
                        }
                        print(f"勤怠データ読み込み: 社員ID='{employee_id}', 日付='{date_str}', 勤務タイプ='{work_type}', 出勤='{check_in}', 退勤='{check_out}', 休憩='{rest_time}', サブタスク={subtasks}")
                    except Exception as row_e:
                        print(f"勤怠データ読み込みエラー (行 {row}): {row_e}")

            # Load Tasks
            ws = self.workbook.Worksheets("Tasks")
            print(f"Tasksシートの使用範囲行数: {ws.UsedRange.Rows.Count}")
            if ws.UsedRange.Rows.Count > 1:
                for row in range(2, ws.UsedRange.Rows.Count + 1):
                    try:
                        raw_employee_id = ws.Cells(row, 1).Value
                        employee_id = str(int(raw_employee_id)) if isinstance(raw_employee_id, float) else str(raw_employee_id).strip()
                        category = str(ws.Cells(row, 2).Value or "").strip()
                        task_name = str(ws.Cells(row, 3).Value or "").strip()
                        if employee_id not in all_data["tasks"]:
                            all_data["tasks"][employee_id] = {"顧客": [], "社内": []}
                        if category and task_name and task_name not in all_data["tasks"][employee_id][category]:
                            all_data["tasks"][employee_id][category].append(task_name)
                        print(f"タスク読み込み: 社員ID='{employee_id}', カテゴリ='{category}', タスク名='{task_name}'")
                    except Exception as row_e:
                        print(f"タスク読み込みエラー (行 {row}): {row_e}")

            # Load Announcements
            ws = self.workbook.Worksheets("Announcements")
            print(f"Announcementsシートの使用範囲行数: {ws.UsedRange.Rows.Count}")
            if ws.UsedRange.Rows.Count > 1:
                for row in range(2, ws.UsedRange.Rows.Count + 1):
                    try:
                        raw_employee_id = ws.Cells(row, 1).Value
                        employee_id = str(int(raw_employee_id)) if isinstance(raw_employee_id, float) else str(raw_employee_id).strip()
                        if employee_id not in all_data["announcements"]:
                            all_data["announcements"][employee_id] = []
                        
                        raw_announcement_date = ws.Cells(row, 2).Value
                        announcement_date = raw_announcement_date.strftime('%Y-%m-%d') if isinstance(raw_announcement_date, datetime) else str(raw_announcement_date or "").strip()
                        announcement_title = str(ws.Cells(row, 3).Value or "").strip()
                        announcement_content = str(ws.Cells(row, 4).Value or "").strip()

                        all_data["announcements"][employee_id].insert(0, { # Insert to beginning to maintain order
                            'date': announcement_date,
                            'title': announcement_title,
                            'content': announcement_content
                        })
                        print(f"お知らせ読み込み: 社員ID='{employee_id}', 日付='{announcement_date}', タイトル='{announcement_title}'")
                    except Exception as row_e:
                        print(f"お知らせ読み込みエラー (行 {row}): {row_e}")

        except Exception as e:
            print(f"Excelデータ読み込み中に致命的なエラーが発生しました: {e}")
        print("--- Excelデータ読み込み完了 ---")
        return all_data

    def save_all_data(self, all_data):
        print("--- Excelデータ保存開始 ---")
        try:
            # Save Attendance
            ws = self.workbook.Worksheets("Attendance")
            ws.UsedRange.ClearContents() # Clear only contents, not formatting
            ws.Cells(1, 1).Value = 'EmployeeID'
            ws.Cells(1, 2).Value = 'Date'
            ws.Cells(1, 3).Value = 'WorkType'
            ws.Cells(1, 4).Value = 'CheckIn'
            ws.Cells(1, 5).Value = 'CheckOut'
            ws.Cells(1, 6).Value = 'RestTime'
            ws.Cells(1, 7).Value = 'Subtasks'
            row = 2
            for employee_id, attendance_by_date in all_data["attendance"].items():
                for date, day_data in sorted(attendance_by_date.items()):
                    ws.Cells(row, 1).Value = employee_id
                    ws.Cells(row, 2).Value = date
                    ws.Cells(row, 3).Value = day_data.get('work_type', '')
                    ws.Cells(row, 4).Value = "'" + str(day_data.get('check_in', '')) # Prepend ' to save as string
                    ws.Cells(row, 5).Value = "'" + str(day_data.get('check_out', '')) # Prepend ' to save as string
                    ws.Cells(row, 6).Value = "'" + str(day_data.get('rest_time', '01:00')) # Prepend ' to save as string
                    ws.Cells(row, 7).Value = json.dumps(day_data.get('subtasks', []), ensure_ascii=False)
                    print(f"勤怠データ保存: 社員ID={employee_id}, 日付={date}, 出勤={day_data.get('check_in', '')}, 退勤={day_data.get('check_out', '')}")
                    row += 1

            # Save Tasks
            ws = self.workbook.Worksheets("Tasks")
            ws.UsedRange.ClearContents()
            ws.Cells(1, 1).Value = 'EmployeeID'
            ws.Cells(1, 2).Value = 'Category'
            ws.Cells(1, 3).Value = 'TaskName'
            row = 2
            for employee_id, tasks_by_category in all_data["tasks"].items():
                for category, tasks in tasks_by_category.items():
                    for task_name in tasks:
                        ws.Cells(row, 1).Value = employee_id
                        ws.Cells(row, 2).Value = category
                        ws.Cells(row, 3).Value = task_name
                        print(f"タスク保存: 社員ID={employee_id}, カテゴリ={category}, タスク名={task_name}")
                        row += 1

            # Save Announcements
            ws = self.workbook.Worksheets("Announcements")
            ws.UsedRange.ClearContents()
            ws.Cells(1, 1).Value = 'EmployeeID'
            ws.Cells(1, 2).Value = 'Date'
            ws.Cells(1, 3).Value = 'Title'
            ws.Cells(1, 4).Value = 'Content'
            row = 2
            for employee_id, announcements_list in all_data["announcements"].items():
                # Announcements are stored newest first in Python, save in that order
                for announcement in announcements_list:
                    ws.Cells(row, 1).Value = employee_id
                    ws.Cells(row, 2).Value = announcement.get('date')
                    ws.Cells(row, 3).Value = announcement.get('title')
                    ws.Cells(row, 4).Value = announcement.get('content')
                    print(f"お知らせ保存: 社員ID={employee_id}, タイトル={announcement.get('title')}")
                    row += 1

            self.workbook.Save()
        except Exception as e:
            print(f"Excelデータの保存エラー: {e}")
        print("--- Excelデータ保存完了 ---")

    def shutdown(self):
        if self.workbook:
            self.workbook.Close(SaveChanges=True) # Ensure changes are saved on close
        if self.excel_app:
            self.excel_app.Quit()
            print("Excelプロセスを終了しました。")

# --- Backend Class ---
class Backend(QObject):
    dataLoaded = Signal(dict)
    dayDataChanged = Signal(str, dict)
    taskUpdated = Signal(dict)
    announcementUpdated = Signal(list)
    showEmployeeIdPrompt = Signal() # New signal to show prompt

    def __init__(self):
        super().__init__()
        self.excel_manager = ExcelManager(EXCEL_FILE_PATH)
        self.all_app_data = self.excel_manager.load_all_data() # Load all data initially
        self.employee_id = None
        self.showEmployeeIdPrompt.emit() # Emit signal to show prompt on startup

    @Slot(str)
    def setEmployeeId(self, employee_id):
        self.employee_id = employee_id
        self._load_employee_data()
        print(f"社員番号が設定されました: {self.employee_id}")

    def _load_employee_data(self):
        # Filter data for the current employee
        employee_data = {
            "attendance": self.all_app_data["attendance"].get(self.employee_id, {}),
            "tasks": self.all_app_data["tasks"].get(self.employee_id, {"顧客": [], "社内": []}),
            "announcements": self.all_app_data["announcements"].get(self.employee_id, []),
            "holidays": [] # Initialize holidays list
        }
        
        # Get holidays for the current month
        today = datetime.now()
        for holiday_date, holiday_name in jpholiday.month_holidays(today.year, today.month):
            employee_data["holidays"].append(holiday_date.strftime("%Y-%m-%d"))

        self.dataLoaded.emit(employee_data)

    @Slot()
    def requestInitialData(self):
        if self.employee_id:
            self._load_employee_data()
        else:
            print("社員番号が設定されていないため、初期データを要求できません。")

    @Slot()
    def checkIn(self):
        if not self.employee_id: return print("社員番号が設定されていません。")
        now = datetime.now()
        today_str = now.strftime("%Y-%m-%d")
        check_in_time = round_up_time(now).strftime("%H:%M")
        
        if self.employee_id not in self.all_app_data["attendance"]:
            self.all_app_data["attendance"][self.employee_id] = {}
        if today_str not in self.all_app_data["attendance"][self.employee_id]:
            self.all_app_data["attendance"][self.employee_id][today_str] = {}
        
        day_data = self.all_app_data["attendance"][self.employee_id][today_str]
        day_data["check_in"] = check_in_time
        if "work_type" not in day_data: day_data["work_type"] = "出勤"
        
        self.excel_manager.save_all_data(self.all_app_data)
        self.dayDataChanged.emit(today_str, day_data)
        print(f"✅ 出勤処理: {today_str} {check_in_time}")

    @Slot()
    def checkOut(self):
        if not self.employee_id: return print("社員番号が設定されていません。")
        now = datetime.now()
        today_str = now.strftime("%Y-%m-%d")
        check_out_time = round_down_time(now).strftime("%H:%M")

        if self.employee_id not in self.all_app_data["attendance"]:
            self.all_app_data["attendance"][self.employee_id] = {}
        if today_str not in self.all_app_data["attendance"][self.employee_id]:
            self.all_app_data["attendance"][self.employee_id][today_str] = {}

        day_data = self.all_app_data["attendance"][self.employee_id][today_str]
        day_data["check_out"] = check_out_time
        if "work_type" not in day_data: day_data["work_type"] = "出勤"

        self.excel_manager.save_all_data(self.all_app_data)
        self.dayDataChanged.emit(today_str, day_data)
        print(f"✅ 退勤処理: {today_str} {check_out_time}")

    @Slot(str, dict)
    def updateDayData(self, date, new_data):
        if not self.employee_id: return print("社員番号が設定されていません。")
        if self.employee_id not in self.all_app_data["attendance"]:
            self.all_app_data["attendance"][self.employee_id] = {}
        if date not in self.all_app_data['attendance'][self.employee_id]:
            self.all_app_data['attendance'][self.employee_id][date] = {}
        
        self.all_app_data['attendance'][self.employee_id][date].update(new_data)
        self.excel_manager.save_all_data(self.all_app_data)
        self.dayDataChanged.emit(date, self.all_app_data['attendance'][self.employee_id][date])
        print(f"✅ データ更新と信号送信: {date}")

    @Slot(str, str)
    def defineTask(self, category, task_name):
        if not self.employee_id: return print("社員番号が設定されていません。")
        if self.employee_id not in self.all_app_data["tasks"]:
            self.all_app_data["tasks"][self.employee_id] = {"顧客": [], "社内": []}
        if category in self.all_app_data["tasks"][self.employee_id] and task_name not in self.all_app_data["tasks"][self.employee_id][category]:
            self.all_app_data["tasks"][self.employee_id][category].append(task_name)
            self.excel_manager.save_all_data(self.all_app_data)
            self.taskUpdated.emit(self.all_app_data["tasks"][self.employee_id])
            print(f"✅ タスク追加: [{category}] {task_name}")

    @Slot(str, str)
    def deleteTask(self, category, task_name):
        if not self.employee_id: return print("社員番号が設定されていません。")
        if self.employee_id in self.all_app_data["tasks"] and category in self.all_app_data["tasks"][self.employee_id] and task_name in self.all_app_data["tasks"][self.employee_id][category]:
            self.all_app_data["tasks"][self.employee_id][category].remove(task_name)
            self.excel_manager.save_all_data(self.all_app_data)
            self.taskUpdated.emit(self.all_app_data["tasks"][self.employee_id])
            print(f"✅ タスク削除: [{category}] {task_name}")

    @Slot(str, str)
    def addAnnouncement(self, title, content):
        if not self.employee_id: return print("社員番号が設定されていません。")
        new_announcement = {"title": title, "content": content, "date": datetime.now().strftime("%Y-%m-%d")}
        if self.employee_id not in self.all_app_data["announcements"]:
            self.all_app_data["announcements"][self.employee_id] = []
        self.all_app_data["announcements"][self.employee_id].insert(0, new_announcement)
        self.excel_manager.save_all_data(self.all_app_data)
        self.announcementUpdated.emit(self.all_app_data["announcements"][self.employee_id])
        print(f"✅ お知らせ追加: {title}")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("勤怠管理システム")
        self.setGeometry(100, 100, 1600, 900)
        self.view = QWebEngineView()
        html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "index.html")
        self.view.load(f"file:///{html_path.replace(os.sep, '/')}")
        self.setCentralWidget(self.view)
        self.backend = Backend()
        self.channel = QWebChannel()
        self.channel.registerObject("backend", self.backend)
        self.view.page().setWebChannel(self.channel)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())