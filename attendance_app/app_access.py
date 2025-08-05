import sys
import os
import json
from datetime import datetime, timedelta
import win32com.client
import atexit
import jpholiday
import calendar

from PySide6.QtWidgets import QApplication, QMainWindow
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtCore import QObject, Slot, Signal

# --- Constants ---
if getattr(sys, 'frozen', False):
    app_path = os.path.dirname(sys.executable)
else:
    app_path = os.path.dirname(os.path.abspath(__file__))

DB_FILE_PATH = os.path.join(app_path, "attendance_data.accdb")

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

# --- Database Management (ADO Version) ---
class DatabaseManager:
    def __init__(self, filepath):
        self.filepath = filepath
        self.connection = None
        self.provider = "Microsoft.ACE.OLEDB.12.0" # For .accdb, common provider

        db_exists = os.path.exists(self.filepath)

        if not db_exists:
            print("データベースファイルが見つかりません。新しいファイルを作成します。")
            try:
                catalog = win32com.client.Dispatch("ADOX.Catalog")
                connection_string = f'Provider={self.provider};Data Source={self.filepath};'
                catalog.Create(connection_string)
                catalog = None
                print("データベースファイルの作成に成功しました。")
            except Exception as e:
                print(f"ADOXを使用したデータベース作成エラー: {e}")
                print("\n***\nエラー: Microsoft Access Database Engineが見つからない可能性があります。\n")
                print("お使いのPythonのビット数（32ビットまたは64ビット）に合った「Microsoft Access Database Engine 2016 Redistributable」をインストールする必要があるかもしれません。\n***\n")
                raise

        try:
            self.connection = win32com.client.Dispatch("ADODB.Connection")
            self.connection.Open(f'Provider={self.provider};Data Source={self.filepath};')
            print("データベースに正常に接続しました。")
        except Exception as e:
            print(f"ADOを使用したデータベース接続エラー: {e}")
            print("\n***\nエラー: Microsoft Access Database Engineが見つからない可能性があります。\n")
            print("お使いのPythonのビット数（32ビットまたは64ビット）に合った「Microsoft Access Database Engine 2016 Redistributable」をインストールする必要があるかもしれません。\n***\n")
            raise

        if not db_exists:
            self._create_tables()

        atexit.register(self.shutdown)

    def _execute(self, sql):
        try:
            self.connection.Execute(sql)
        except Exception as e:
            print(f"SQL実行エラー: {sql} - {e}")

    def _query(self, sql):
        try:
            recordset = win32com.client.Dispatch("ADODB.Recordset")
            recordset.Open(sql, self.connection, 1, 3) # adOpenKeyset, adLockOptimistic
            
            if recordset.EOF and recordset.BOF:
                return []

            fields = [field.Name for field in recordset.Fields]
            data = recordset.GetRows()
            recordset.Close()

            if not data:
                return []
            
            return [dict(zip(fields, row)) for row in zip(*data)]
        except Exception as e:
            print(f"SQLクエリエラー: {sql} - {e}")
            return []

    def _create_tables(self):
        print("テーブルの作成を開始します...")
        try:
            self._execute("""
                CREATE TABLE Attendance (
                    ID AUTOINCREMENT PRIMARY KEY,
                    EmployeeID TEXT(50),
                    AttendanceDate DATE,
                    WorkType TEXT(50),
                    CheckIn TEXT(10),
                    CheckOut TEXT(10),
                    RestTime TEXT(10),
                    Subtasks MEMO
                );
            """)
            self._execute("""
                CREATE TABLE Tasks (
                    ID AUTOINCREMENT PRIMARY KEY,
                    EmployeeID TEXT(50),
                    Category TEXT(50),
                    TaskName TEXT(255)
                );
            """)
            self._execute("""
                CREATE TABLE Announcements (
                    ID AUTOINCREMENT PRIMARY KEY,
                    EmployeeID TEXT(50),
                    AnnouncementDate DATE,
                    Title TEXT(255),
                    Content MEMO
                );
            """)
            print("テーブルの作成が完了しました。")
        except Exception as e:
            print(f"テーブル作成エラー: {e}")

    def load_employee_data(self, employee_id):
        print(f"--- {employee_id}のデータベース読み込み開始 ---")
        
        attendance_sql = f"SELECT * FROM Attendance WHERE EmployeeID='{employee_id}'"
        attendance_records = self._query(attendance_sql)
        attendance_data = {}
        for rec in attendance_records:
            # ADO might return datetime objects, handle them carefully
            raw_date = rec['AttendanceDate']
            date_str = datetime(raw_date.year, raw_date.month, raw_date.day).strftime('%Y-%m-%d')
            attendance_data[date_str] = {
                'work_type': rec.get('WorkType', ''),
                'check_in': rec.get('CheckIn', ''),
                'check_out': rec.get('CheckOut', ''),
                'rest_time': rec.get('RestTime', '01:00'),
                'subtasks': json.loads(rec.get('Subtasks', '[]') or '[]')
            }

        tasks_sql = f"SELECT Category, TaskName FROM Tasks WHERE EmployeeID='{employee_id}'"
        task_records = self._query(tasks_sql)
        tasks_data = {"顧客": [], "社内": []}
        for rec in task_records:
            category = rec.get('Category')
            task_name = rec.get('TaskName')
            if category in tasks_data and task_name:
                tasks_data[category].append(task_name)

        announcements_sql = f"SELECT AnnouncementDate, Title, Content FROM Announcements WHERE EmployeeID='{employee_id}' ORDER BY AnnouncementDate DESC"
        announcement_records = self._query(announcements_sql)
        announcements_data = []
        for rec in announcement_records:
            raw_date = rec['AnnouncementDate']
            date_str = datetime(raw_date.year, raw_date.month, raw_date.day).strftime('%Y-%m-%d')
            announcements_data.append({
                'date': date_str,
                'title': rec.get('Title', ''),
                'content': rec.get('Content', '')
            })
        
        print(f"--- {employee_id}のデータベース読み込み完了 ---")
        return {"attendance": attendance_data, "tasks": tasks_data, "announcements": announcements_data}

    def update_attendance(self, employee_id, date_str, day_data):
        subtasks_json = json.dumps(day_data.get('subtasks', []), ensure_ascii=False).replace("'", "''")
        work_type = (day_data.get('work_type', '') or '').replace("'", "''")
        check_in = (day_data.get('check_in', '') or '').replace("'", "''")
        check_out = (day_data.get('check_out', '') or '').replace("'", "''")
        rest_time = (day_data.get('rest_time', '01:00') or '').replace("'", "''")

        check_sql = f"SELECT ID FROM Attendance WHERE EmployeeID='{employee_id}' AND AttendanceDate=#{date_str}#"
        existing = self._query(check_sql)

        if existing:
            sql = f"""
                UPDATE Attendance SET
                    WorkType = '{work_type}',
                    CheckIn = '{check_in}',
                    CheckOut = '{check_out}',
                    RestTime = '{rest_time}',
                    Subtasks = '{subtasks_json}'
                WHERE EmployeeID='{employee_id}' AND AttendanceDate=#{date_str}#
            """
        else:
            sql = f"""
                INSERT INTO Attendance (EmployeeID, AttendanceDate, WorkType, CheckIn, CheckOut, RestTime, Subtasks)
                VALUES (
                    '{employee_id}',
                    #{date_str}#,
                    '{work_type or '出勤'}',
                    '{check_in}',
                    '{check_out}',
                    '{rest_time}',
                    '{subtasks_json}'
                )
            """
        self._execute(sql)
        print(f"勤怠データを更新しました: {employee_id} - {date_str}")

    def add_task(self, employee_id, category, task_name):
        safe_task_name = task_name.replace("'", "''")
        sql = f"INSERT INTO Tasks (EmployeeID, Category, TaskName) VALUES ('{employee_id}', '{category}', '{safe_task_name}')"
        self._execute(sql)
        print(f"タスクを追加しました: {employee_id} - [{category}] {task_name}")

    def delete_task(self, employee_id, category, task_name):
        safe_task_name = task_name.replace("'", "''")
        sql = f"DELETE FROM Tasks WHERE EmployeeID='{employee_id}' AND Category='{category}' AND TaskName='{safe_task_name}'"
        self._execute(sql)
        print(f"タスクを削除しました: {employee_id} - [{category}] {task_name}")

    def add_announcement(self, employee_id, title, content, date_str):
        safe_title = title.replace("'", "''")
        safe_content = content.replace("'", "''")
        sql = f"INSERT INTO Announcements (EmployeeID, AnnouncementDate, Title, Content) VALUES ('{employee_id}', #{date_str}#, '{safe_title}', '{safe_content}')"
        self._execute(sql)
        print(f"お知らせを追加しました: {employee_id} - {title}")

    def shutdown(self):
        if self.connection and self.connection.State == 1: # 1 == adStateOpen
            self.connection.Close()
        self.connection = None
        print("データベース接続を閉じました。")

# --- Backend Class ---
class Backend(QObject):
    dataLoaded = Signal(dict)
    dayDataChanged = Signal(str, dict)
    taskUpdated = Signal(dict)
    announcementUpdated = Signal(list)
    showEmployeeIdPrompt = Signal()

    def __init__(self):
        super().__init__()
        self.db_manager = DatabaseManager(DB_FILE_PATH)
        self.employee_id = None
        self.showEmployeeIdPrompt.emit()

    @Slot(str)
    def setEmployeeId(self, employee_id):
        self.employee_id = employee_id
        self.load_and_emit_employee_data()
        print(f"社員番号が設定されました: {self.employee_id}")

    def load_and_emit_employee_data(self):
        if not self.employee_id: return
        
        employee_data = self.db_manager.load_employee_data(self.employee_id)
        attendance_data = employee_data["attendance"]
        today = datetime.now()
        year, month = today.year, today.month

        # Get all holidays for the current month once
        month_holidays = {d.strftime("%Y-%m-%d") for d, n in jpholiday.month_holidays(year, month)}

        # Iterate through all days of the current month
        for day in range(1, calendar.monthrange(year, month)[1] + 1):
            current_date = datetime(year, month, day)
            date_str = current_date.strftime("%Y-%m-%d")
            day_of_week = current_date.weekday() # Monday is 0 and Sunday is 6

            # If there is no data for this day in the DB
            if date_str not in attendance_data:
                # Check if it's a weekend or a holiday
                if day_of_week >= 5 or date_str in month_holidays:
                    attendance_data[date_str] = {
                        'work_type': '休日',
                        'check_in': '',
                        'check_out': '',
                        'rest_time': '00:00',
                        'subtasks': []
                    }

        employee_data["holidays"] = list(month_holidays)
        self.dataLoaded.emit(employee_data)

    @Slot()
    def requestInitialData(self):
        if self.employee_id:
            self.load_and_emit_employee_data()
        else:
            print("社員番号が設定されていないため、初期データを要求できません。")

    def _get_day_data(self, date_str):
        data = self.db_manager.load_employee_data(self.employee_id)
        return data['attendance'].get(date_str, {'work_type': '出勤', 'check_in': '', 'check_out': '', 'rest_time': '01:00', 'subtasks': []})

    @Slot()
    def checkIn(self):
        if not self.employee_id: return print("社員番号が設定されていません。")
        now = datetime.now()
        today_str = now.strftime("%Y-%m-%d")
        check_in_time = round_up_time(now).strftime("%H:%M")
        
        day_data = self._get_day_data(today_str)
        day_data["check_in"] = check_in_time
        if not day_data.get("work_type"): day_data["work_type"] = "出勤"
        
        self.db_manager.update_attendance(self.employee_id, today_str, day_data)
        self.dayDataChanged.emit(today_str, day_data)
        print(f"✅ 出勤処理: {today_str} {check_in_time}")

    @Slot()
    def checkOut(self):
        if not self.employee_id: return print("社員番号が設定されていません。")
        now = datetime.now()
        today_str = now.strftime("%Y-%m-%d")
        check_out_time = round_down_time(now).strftime("%H:%M")

        day_data = self._get_day_data(today_str)
        day_data["check_out"] = check_out_time
        if not day_data.get("work_type"): day_data["work_type"] = "出勤"

        self.db_manager.update_attendance(self.employee_id, today_str, day_data)
        self.dayDataChanged.emit(today_str, day_data)
        print(f"✅ 退勤処理: {today_str} {check_out_time}")

    @Slot(str, dict)
    def updateDayData(self, date, new_data):
        if not self.employee_id: return print("社員番号が設定されていません。")
        
        self.db_manager.update_attendance(self.employee_id, date, new_data)
        self.dayDataChanged.emit(date, new_data)
        print(f"✅ データ更新と信号送信: {date}")

    @Slot(str, str)
    def defineTask(self, category, task_name):
        if not self.employee_id: return print("社員番号が設定されていません。")
        self.db_manager.add_task(self.employee_id, category, task_name)
        
        all_tasks = self.db_manager.load_employee_data(self.employee_id)["tasks"]
        self.taskUpdated.emit(all_tasks)
        print(f"✅ タスク追加: [{category}] {task_name}")

    @Slot(str, str)
    def deleteTask(self, category, task_name):
        if not self.employee_id: return print("社員番号が設定されていません。")
        self.db_manager.delete_task(self.employee_id, category, task_name)

        all_tasks = self.db_manager.load_employee_data(self.employee_id)["tasks"]
        self.taskUpdated.emit(all_tasks)
        print(f"✅ タスク削除: [{category}] {task_name}")

    @Slot(str, str)
    def addAnnouncement(self, title, content):
        if not self.employee_id: return print("社員番号が設定されていません。")
        date_str = datetime.now().strftime("%Y-%m-%d")
        self.db_manager.add_announcement(self.employee_id, title, content, date_str)

        all_announcements = self.db_manager.load_employee_data(self.employee_id)["announcements"]
        self.announcementUpdated.emit(all_announcements)
        print(f"✅ お知らせ追加: {title}")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("勤怠管理システム (Access DB - ADO)")
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