# 変更されたファイル全体コード (2025-08-05)

このマークダウンファイルは `attendance_app` プロジェクトの変更されたすべてのファイルの最終コードを含みます。

---

## 1. `app_access.py`

```python
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
            self._execute("""
                CREATE TABLE Comments (
                    ID AUTOINCREMENT PRIMARY KEY,
                    AnnouncementID LONG,
                    AuthorName TEXT(100),
                    CommentText MEMO,
                    CommentDate DATE
                );
            """)
            self._execute("""
                CREATE TABLE Users (
                    EmployeeID TEXT(50) PRIMARY KEY,
                    UserName TEXT(100)
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

        announcements_sql = f"SELECT ID, AnnouncementDate, Title, Content FROM Announcements WHERE EmployeeID='{employee_id}' ORDER BY AnnouncementDate DESC"
        announcement_records = self._query(announcements_sql)
        announcements_data = []
        for rec in announcement_records:
            raw_date = rec['AnnouncementDate']
            date_str = datetime(raw_date.year, raw_date.month, raw_date.day).strftime('%Y-%m-%d')
            announcements_data.append({
                'ID': rec.get('ID'),
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

    def get_user_name(self, employee_id):
        sql = f"SELECT UserName FROM Users WHERE EmployeeID='{employee_id}'"
        result = self._query(sql)
        return result[0]['UserName'] if result else None

    def set_user_name(self, employee_id, user_name):
        safe_name = user_name.replace("'", "''")
        check_sql = f"SELECT EmployeeID FROM Users WHERE EmployeeID='{employee_id}'"
        if self._query(check_sql):
            sql = f"UPDATE Users SET UserName='{safe_name}' WHERE EmployeeID='{employee_id}'"
        else:
            sql = f"INSERT INTO Users (EmployeeID, UserName) VALUES ('{employee_id}', '{safe_name}')"
        self._execute(sql)
        print(f"ユーザー名を設定しました: {employee_id} - {user_name}")

    def get_announcement_details(self, announcement_id):
        announcement_sql = f"SELECT * FROM Announcements WHERE ID={announcement_id}"
        announcement_result = self._query(announcement_sql)
        if not announcement_result: return None

        comments_sql = f"SELECT AuthorName, CommentText, CommentDate FROM Comments WHERE AnnouncementID={announcement_id} ORDER BY CommentDate ASC"
        comments_result = self._query(comments_sql)

        details = announcement_result[0]
        # Convert datetime objects to strings for JSON serialization
        if isinstance(details.get('AnnouncementDate'), datetime):
            details['AnnouncementDate'] = details['AnnouncementDate'].strftime('%Y-%m-%d')
        
        comments_data = []
        for comment in comments_result:
            if isinstance(comment.get('CommentDate'), datetime):
                comment['CommentDate'] = comment['CommentDate'].strftime('%Y-%m-%d %H:%M')
            comments_data.append(comment)
        details['Comments'] = comments_data
        return details

    def add_comment(self, announcement_id, author_name, comment_text, comment_date):
        safe_author = author_name.replace("'", "''")
        safe_comment = comment_text.replace("'", "''")
        sql = f"INSERT INTO Comments (AnnouncementID, AuthorName, CommentText, CommentDate) VALUES ({announcement_id}, '{safe_author}', '{safe_comment}', #{comment_date}#)"
        self._execute(sql)
        print(f"コメントを追加しました: AnnouncementID={announcement_id}")

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
    showAlert = Signal(str)
    announcementDetailsLoaded = Signal(dict)
    userNameRequired = Signal()

    def __init__(self):
        super().__init__()
        self.db_manager = DatabaseManager(DB_FILE_PATH)
        self.employee_id = None
        self.user_name = None
        self.showEmployeeIdPrompt.emit()

    @Slot(str)
    def setEmployeeId(self, employee_id):
        self.employee_id = employee_id
        self.user_name = self.db_manager.get_user_name(self.employee_id)
        self.load_and_emit_employee_data()
        print(f"社員番号が設定されました: {self.employee_id}")

    def load_and_emit_employee_data(self):
        if not self.employee_id: return
        
        employee_data = self.db_manager.load_employee_data(self.employee_id)
        attendance_data = employee_data["attendance"]
        today = datetime.now()
        year, month = today.year, today.month

        month_holidays = {d.strftime("%Y-%m-%d") for d, n in jpholiday.month_holidays(year, month)}

        for day in range(1, calendar.monthrange(year, month)[1] + 1):
            current_date = datetime(year, month, day)
            date_str = current_date.strftime("%Y-%m-%d")
            day_of_week = current_date.weekday()

            if date_str not in attendance_data:
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

    @Slot(int)
    def getAnnouncementDetails(self, announcement_id):
        if not self.employee_id: return
        details = self.db_manager.get_announcement_details(announcement_id)
        if details:
            self.announcementDetailsLoaded.emit(details)

    @Slot(str)
    def setUserName(self, user_name):
        if not self.employee_id: return
        self.db_manager.set_user_name(self.employee_id, user_name)
        self.user_name = user_name
        self.showAlert.emit(f"ようこそ、{user_name}さん！")

    @Slot(int, str)
    def addComment(self, announcement_id, comment_text):
        if not self.employee_id: return
        if not self.user_name:
            self.userNameRequired.emit()
            return
        
        comment_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.db_manager.add_comment(announcement_id, self.user_name, comment_text, comment_date)
        self.getAnnouncementDetails(announcement_id)


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
```

---

## 2. `templates/index.html`

```html
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>勤怠管理システム</title>
    <link rel="stylesheet" href="../static/style.css">
    <script src="qrc:///qtwebchannel/qwebchannel.js"></script>
</head>
<body>
    <div class="container">
        <header>
            <h1>勤怠管理システム</h1>
        </header>

        <main id="main-content" style="display: none;">
            <div class="top-row">
                <section class="card attendance-card">
                    <h2>クイック記録</h2>
                    <div class="button-group">
                        <button id="check-in">出勤</button>
                        <button id="check-out">退勤</button>
                    </div>
                </section>
                <section class="card announcements-card">
                    <div class="card-header">
                        <h2>お知らせ</h2>
                        <button id="open-announcement-modal" class="add-button">+</button>
                    </div>
                    <div id="announcements-list"></div>
                </section>
            </div>

            <section class="card monthly-record-card">
                <h2>月別記録</h2>
                <div id="monthly-summary"></div>
                <div class="calendar-nav">
                    <button id="prev-month">&lt; 前の月</button>
                    <span id="current-month-year"></span>
                    <button id="next-month">次の月 &gt;</button>
                </div>
                <div id="calendar-container">
                    <div id="calendar-grid"></div>
                </div>
            </section>

            <section class="card task-management-card">
                <h2>タスク管理</h2>
                <div class="input-group">
                    <select id="task-category">
                        <option value="顧客">顧客</option>
                        <option value="社内">社内</option>
                    </select>
                    <input type="text" id="task-name" placeholder="新しいタスク名">
                    <button id="define-task">タスク追加</button>
                </div>
                <div id="defined-tasks-list"></div>
            </section>
        </main>
    </div>

    <!-- Modals -->
    <div id="employee-id-modal" class="modal" style="display: flex;">
        <div class="modal-content">
            <h2>社員番号入力</h2>
            <div class="input-group">
                <input type="text" id="modal-employee-id" placeholder="社員番号を入力してください">
                <button id="submit-employee-id">確認</button>
            </div>
        </div>
    </div>

    <div id="user-name-modal" class="modal">
        <div class="modal-content">
            <h2>お名前を入力してください</h2>
            <p>コメント機能を使用するために、表示される名前を一度だけ入力してください。</p>
            <div class="input-group">
                <input type="text" id="modal-user-name" placeholder="表示名">
                <button id="submit-user-name">登録</button>
            </div>
        </div>
    </div>

    <div id="alert-modal" class="modal">
        <div class="modal-content">
            <span class="close-button">&times;</span>
            <p id="alert-message"></p>
        </div>
    </div>

    <div id="announcement-detail-modal" class="modal">
        <div class="modal-content wide">
            <span class="close-button">&times;</span>
            <h2 id="detail-title"></h2>
            <p id="detail-meta"></p>
            <hr>
            <div id="detail-content" class="modal-scroll-content"></div>
            <hr>
            <h3>コメント</h3>
            <div id="detail-comments" class="modal-scroll-content"></div>
            <div class="input-group">
                <textarea id="comment-text" placeholder="コメントを追加..."></textarea>
                <button id="submit-comment">コメントする</button>
            </div>
        </div>
    </div>

    <div id="announcement-create-modal" class="modal">
        <div class="modal-content">
            <span class="close-button">&times;</span>
            <h2>新しいお知らせ作成</h2>
            <div class="input-group">
                <input type="text" id="announcement-title" placeholder="タイトル">
                <textarea id="announcement-content" placeholder="内容"></textarea>
                <button id="add-announcement">お知らせ追加</button>
            </div>
        </div>
    </div>

    <script>
        // --- Global Variables ---
        let backend;
        let currentYear, currentMonth;
        let attendanceData = {};
        let allTasks = { "顧客": [], "社内": [] };
        let holidays = [];
        let currentAnnouncementId = null;
        const workTypes = ["出勤", "在宅", "有給", "祝日出勤", "休日", "午前有給", "午後有給", "欠勤"];
        const typesWithoutTime = ["有給", "休日", "欠勤"];

        // --- Utility Functions ---
        function toYYYYMMDD(date) {
            const y = date.getFullYear();
            const m = String(date.getMonth() + 1).padStart(2, '0');
            const d = String(date.getDate()).padStart(2, '0');
            return `${y}-${m}-${d}`;
        }

        function timeStrToMinutes(timeStr) {
            if (!timeStr) return 0;
            if (timeStr.includes(':')) {
                const [h, m] = timeStr.split(':').map(Number);
                return h * 60 + m;
            } else {
                return parseFloat(timeStr) * 60;
            }
        }

        function formatTimeInput(event) {
            let value = event.target.value.replace(/[^\d]/g, '');
            if (value.length === 4) {
                event.target.value = `${value.slice(0, 2)}:${value.slice(2, 4)}`;
            }
        }

        // --- Modal Management ---
        function setupModal(modalId, openTriggerId, closeSelector) {
            const modal = document.getElementById(modalId);
            if (openTriggerId) {
                document.getElementById(openTriggerId).addEventListener('click', () => modal.style.display = 'flex');
            }
            modal.querySelector(closeSelector).addEventListener('click', () => modal.style.display = 'none');
            window.addEventListener('click', (event) => {
                if (event.target == modal) {
                    modal.style.display = 'none';
                }
            });
            return modal;
        }

        // --- DOMContentLoaded ---
        window.addEventListener("DOMContentLoaded", () => {
            const now = new Date();
            currentYear = now.getFullYear();
            currentMonth = now.getMonth();

            // Setup Modals
            setupModal('announcement-create-modal', 'open-announcement-modal', '.close-button');
            setupModal('announcement-detail-modal', null, '.close-button');
            setupModal('alert-modal', null, '.close-button');

            new QWebChannel(qt.webChannelTransport, function (channel) {
                backend = channel.objects.backend;

                // Connect signals
                backend.dataLoaded.connect(initializeUI);
                backend.dayDataChanged.connect(updateDayOnCalendar);
                backend.taskUpdated.connect(renderTasks);
                backend.announcementUpdated.connect(renderAnnouncements);
                backend.showEmployeeIdPrompt.connect(() => document.getElementById('employee-id-modal').style.display = 'flex');
                backend.showAlert.connect(showAlert);
                backend.announcementDetailsLoaded.connect(showAnnouncementDetails);
                backend.userNameRequired.connect(() => document.getElementById('user-name-modal').style.display = 'flex');

                // Bind events
                document.getElementById("check-in").addEventListener("click", () => backend.checkIn());
                document.getElementById("check-out").addEventListener("click", () => backend.checkOut());
                document.getElementById('prev-month').addEventListener('click', () => changeMonth(-1));
                document.getElementById('next-month').addEventListener('click', () => changeMonth(1));
                document.getElementById('define-task').addEventListener('click', defineTask);
                document.getElementById('add-announcement').addEventListener('click', addAnnouncement);
                document.getElementById('submit-employee-id').addEventListener('click', submitEmployeeId);
                document.getElementById('modal-employee-id').addEventListener('keypress', (e) => { if (e.key === 'Enter') submitEmployeeId(); });
                document.getElementById('submit-user-name').addEventListener('click', submitUserName);
                document.getElementById('submit-comment').addEventListener('click', submitComment);
            });
        });

        // --- Backend Interaction ---
        function submitEmployeeId() {
            const employeeId = document.getElementById('modal-employee-id').value.trim();
            if (employeeId) {
                backend.setEmployeeId(employeeId);
                document.getElementById('employee-id-modal').style.display = 'none';
                document.getElementById('main-content').style.display = 'flex';
            } else {
                showAlert("社員番号を入力してください。");
            }
        }

        function submitUserName() {
            const userName = document.getElementById('modal-user-name').value.trim();
            if (userName) {
                backend.setUserName(userName);
                document.getElementById('user-name-modal').style.display = 'none';
            } else {
                showAlert("表示名を入力してください。");
            }
        }

        function submitComment() {
            const commentText = document.getElementById('comment-text').value.trim();
            if (commentText && currentAnnouncementId) {
                backend.addComment(currentAnnouncementId, commentText);
                document.getElementById('comment-text').value = '';
            }
        }

        function defineTask() {
            const category = document.getElementById('task-category').value;
            const nameInput = document.getElementById('task-name');
            const name = nameInput.value.trim();
            if (name) {
                backend.defineTask(category, name);
                nameInput.value = '';
            }
        }

        function addAnnouncement() {
            const modal = document.getElementById('announcement-create-modal');
            const title = modal.querySelector('#announcement-title').value.trim();
            const content = modal.querySelector('#announcement-content').value.trim();
            if (title && content) {
                backend.addAnnouncement(title, content);
                modal.querySelector('#announcement-title').value = '';
                modal.querySelector('#announcement-content').value = '';
                modal.style.display = 'none';
            }
        }

        function handleDayDataChange(event, dateStr) {
            const dayCell = document.querySelector(`.calendar-day[data-date='${dateStr}']`);
            if (event.target.classList.contains('work-type')) {
                const workType = event.target.value;
                if (workType === "午前有給" || workType === "午後有給") {
                    const restTimeInput = dayCell.querySelector('.rest-time');
                    if (restTimeInput) restTimeInput.value = "00:00";
                }
            }
            let new_data = {};
            dayCell.querySelectorAll('.day-input[data-field]').forEach(input => { new_data[input.dataset.field] = input.value; });
            const subtaskEntries = dayCell.querySelectorAll('.subtask-entry');
            new_data.subtasks = Array.from(subtaskEntries).map(entry => ({ name: entry.dataset.taskName, time: entry.querySelector('.subtask-time').value }));
            backend.updateDayData(dateStr, new_data);
        }

        // --- UI Rendering & Updates ---
        function initializeUI(data) {
            attendanceData = data.attendance || {};
            allTasks = data.tasks || { "顧客": [], "社内": [] };
            holidays = data.holidays || [];
            renderAnnouncements(data.announcements || []);
            renderTasks(allTasks);
            renderCalendar(currentYear, currentMonth, attendanceData);
            renderMonthlySummary(currentYear, currentMonth, attendanceData);
        }

        function showAlert(message) {
            document.getElementById('alert-message').textContent = message;
            document.getElementById('alert-modal').style.display = 'flex';
        }

        function showAnnouncementDetails(details) {
            if (!details) return;
            currentAnnouncementId = details.ID;
            document.getElementById('detail-title').textContent = details.Title;
            document.getElementById('detail-meta').textContent = `作成日: ${new Date(details.AnnouncementDate).toLocaleDateString()}`;
            document.getElementById('detail-content').innerHTML = details.Content.replace(/\n/g, '<br>');

            const commentsContainer = document.getElementById('detail-comments');
            commentsContainer.innerHTML = (details.Comments || []).map(c =>
                `<div class="comment">
                    <p><strong>${c.AuthorName}</strong> <span class="comment-date">(${new Date(c.CommentDate).toLocaleString()})</span></p>
                    <p>${c.CommentText}</p>
                </div>`
            ).join('');
            commentsContainer.scrollTop = commentsContainer.scrollHeight;

            document.getElementById('announcement-detail-modal').style.display = 'flex';
        }

        function renderAnnouncements(announcements) {
            const list = document.getElementById('announcements-list');
            list.innerHTML = announcements.map(a =>
                `<div class="announcement-item" data-id="${a.ID}">
                    <h4>${a.Title} (${new Date(a.date).toLocaleDateString()})</h4>
                </div>`
            ).join('');
            list.querySelectorAll('.announcement-item').forEach(item => {
                item.addEventListener('click', () => {
                    backend.getAnnouncementDetails(parseInt(item.dataset.id));
                });
            });
        }
        
        function changeMonth(direction) {
            currentMonth += direction;
            if (currentMonth < 0) { currentMonth = 11; currentYear--; }
            else if (currentMonth > 11) { currentMonth = 0; currentYear++; }
            backend.requestInitialData();
        }

        function renderCalendar(year, month, data) {
            attendanceData = data;
            const calendarGrid = document.getElementById('calendar-grid');
            document.getElementById('current-month-year').textContent = `${year}年 ${month + 1}月`;
            calendarGrid.innerHTML = '';
            const daysInMonth = new Date(year, month + 1, 0).getDate();
            for (let i = 1; i <= daysInMonth; i++) {
                const date = new Date(year, month, i);
                const dateStr = toYYYYMMDD(date);
                const dayData = attendanceData[dateStr] || {};
                const dayCell = createDayCell(date, dayData);
                calendarGrid.appendChild(dayCell);
            }
        }

        function createDayCell(date, dayData) {
            const dateStr = toYYYYMMDD(date);
            const dayCell = document.createElement('div');
            dayCell.className = 'calendar-day';
            dayCell.dataset.date = dateStr;

            const dayOfWeekNum = date.getDay();
            const isWeekend = dayOfWeekNum === 0 || dayOfWeekNum === 6;
            const isHoliday = holidays.includes(dateStr);

            if (isWeekend) dayCell.classList.add(dayOfWeekNum === 0 ? 'sunday' : 'saturday');
            if (isHoliday) dayCell.classList.add('holiday');

            let workType = dayData.work_type;
            if (!workType && (isWeekend || isHoliday)) workType = "休日";
            
            const isTimeHidden = typesWithoutTime.includes(workType);
            const subtasks = dayData.subtasks || [];
            const { workHours, overtimeHours } = calculateWorkHours(workType, dayData.check_in, dayData.check_out, dayData.rest_time);
            const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];

            const dailyTaskSummary = subtasks.reduce((acc, task) => {
                const taskCategory = Object.keys(allTasks).find(cat => allTasks[cat].includes(task.name));
                if (taskCategory) acc[taskCategory] = (acc[taskCategory] || 0) + timeStrToMinutes(task.time);
                return acc;
            }, {});

            dayCell.innerHTML = `
                <div class="day-header"><span class="day-number">${date.getDate()}</span><span class="day-of-week">${dayOfWeek}</span></div>
                <div class="day-content">
                    <select class="day-input work-type" data-field="work_type">${workTypes.map(wt => `<option value="${wt}" ${wt === workType ? 'selected' : ''}>${wt}</option>`).join('')}</select>
                    <div class="time-fields-container" ${isTimeHidden ? 'style="display: none;"' : ''}>
                        <div class="time-inputs">
                            <input type="text" class="day-input time-input" data-field="check_in" value="${dayData.check_in || ''}" placeholder="HH:MM"><span>-</span>
                            <input type="text" class="day-input time-input" data-field="check_out" value="${dayData.check_out || ''}" placeholder="HH:MM">
                        </div>
                        <div class="rest-time-input"><label>休憩:</label><input type="text" class="day-input rest-time time-input" data-field="rest_time" value="${dayData.rest_time || '01:00'}"></div>
                        <div class="calculated-times"><span>勤務: <strong>${workHours}</strong></span><span>残業: <strong>${overtimeHours}</strong></span></div>
                    </div>
                    <div class="daily-task-summary">${Object.entries(dailyTaskSummary).map(([cat, min]) => `${cat}: ${(min/60).toFixed(1)}h`).join(', ') || '-'}</div>
                    <div class="subtask-section">
                        <select class="day-input subtask-select"><option value="">+ タスク追加</option><optgroup label="顧客">${allTasks["顧客"].map(t => `<option value="${t}">${t}</option>`).join('')}</optgroup><optgroup label="社内">${allTasks["社内"].map(t => `<option value="${t}">${t}</option>`).join('')}</optgroup></select>
                        <div class="subtask-list">${subtasks.map(st => `<div class="subtask-entry" data-task-name="${st.name}"><span class="subtask-name">${st.name}</span><input type="text" class="day-input subtask-time" value="${st.time || '0.0'}" placeholder="0.0"><button class="delete-subtask">&times;</button></div>`).join('')}</div>
                    </div>
                </div>`;

            dayCell.querySelectorAll('.day-input, .subtask-time').forEach(input => input.addEventListener('change', (e) => handleDayDataChange(e, dateStr)));
            dayCell.querySelectorAll('.time-input').forEach(input => input.addEventListener('input', formatTimeInput));
            dayCell.querySelector('.subtask-select').addEventListener('change', (e) => addSubtask(e, dateStr));
            dayCell.querySelectorAll('.delete-subtask').forEach(button => button.addEventListener('click', (e) => removeSubtask(e, dateStr)));

            return dayCell;
        }

        function updateDayOnCalendar(dateStr, dayData) {
            attendanceData[dateStr] = dayData;
            const dayCell = document.querySelector(`.calendar-day[data-date='${dateStr}']`);
            if (dayCell) {
                const [year, month, day] = dateStr.split('-').map(Number);
                const date = new Date(year, month - 1, day);
                const newDayCell = createDayCell(date, dayData);
                dayCell.replaceWith(newDayCell);
            }
            renderMonthlySummary(currentYear, currentMonth, attendanceData);
        }

        function renderTasks(tasks) {
            allTasks = tasks;
            const listContainer = document.getElementById('defined-tasks-list');
            listContainer.innerHTML = '';
            for (const category in tasks) {
                const categoryDiv = document.createElement('div');
                categoryDiv.className = 'task-category';
                categoryDiv.innerHTML = `<h3>${category}</h3>`;
                const ul = document.createElement('ul');
                tasks[category].forEach(task => {
                    const li = document.createElement('li');
                    li.textContent = task;
                    const deleteBtn = document.createElement('button');
                    deleteBtn.textContent = '削除';
                    deleteBtn.onclick = () => backend.deleteTask(category, task);
                    li.appendChild(deleteBtn);
                    ul.appendChild(li);
                });
                categoryDiv.appendChild(ul);
                listContainer.appendChild(categoryDiv);
            }
            renderCalendar(currentYear, currentMonth, attendanceData);
        }

        function renderMonthlySummary(year, month, data) {
            const summaryContainer = document.getElementById('monthly-summary');
            const workTypeCounts = {};
            let totalOvertimeMinutes = 0;
            const taskCategoryMinutes = {};

            for (const dateStr in data) {
                if (dateStr.startsWith(`${year}-${String(month + 1).padStart(2, '0')}`)) {
                    const dayData = data[dateStr];
                    const workType = dayData.work_type;
                    if (workType) workTypeCounts[workType] = (workTypeCounts[workType] || 0) + 1;
                    const { overtimeHours } = calculateWorkHours(workType, dayData.check_in, dayData.check_out, dayData.rest_time);
                    if (overtimeHours !== '-') totalOvertimeMinutes += parseFloat(overtimeHours) * 60;
                    if (dayData.subtasks) {
                        dayData.subtasks.forEach(task => {
                            const taskCategory = Object.keys(allTasks).find(cat => allTasks[cat].includes(task.name));
                            if (taskCategory) taskCategoryMinutes[taskCategory] = (taskCategoryMinutes[taskCategory] || 0) + timeStrToMinutes(task.time);
                        });
                    }
                }
            }

            const workTypeHTML = Object.entries(workTypeCounts).map(([type, count]) => `<span>${type}: <strong>${count}日</strong></span>`).join('');
            const totalOvertimeHTML = `<span>総残業: <strong>${(totalOvertimeMinutes / 60).toFixed(2)}時間</strong></span>`;
            const taskTimeHTML = Object.entries(taskCategoryMinutes).map(([cat, min]) => `<span>${cat}: <strong>${(min/60).toFixed(2)}時間</strong></span>`).join('');

            summaryContainer.innerHTML = `<div class="summary-group"><h3>勤務日</h3>${workTypeHTML || '-'}</div><div class="summary-group"><h3>合計時間</h3>${totalOvertimeHTML}${taskTimeHTML}</div>`;
        }

        function addSubtask(event, dateStr) {
            const selectedTask = event.target.value;
            if (!selectedTask) return;
            let dayData = JSON.parse(JSON.stringify(attendanceData[dateStr] || {}));
            let subtasks = dayData.subtasks || [];
            if (!subtasks.some(st => st.name === selectedTask)) {
                subtasks.push({ name: selectedTask, time: "0.0" });
                dayData.subtasks = subtasks;
                updateDayOnCalendar(dateStr, dayData);
                backend.updateDayData(dateStr, dayData);
            }
            event.target.value = "";
        }

        function removeSubtask(event, dateStr) {
            const taskToRemove = event.target.closest('.subtask-entry').dataset.taskName;
            let dayData = JSON.parse(JSON.stringify(attendanceData[dateStr] || {}));
            let subtasks = dayData.subtasks || [];
            dayData.subtasks = subtasks.filter(st => st.name !== taskToRemove);
            updateDayOnCalendar(dateStr, dayData);
            backend.updateDayData(dateStr, dayData);
        }

        function calculateWorkHours(workType, checkIn, checkOut, restTimeStr) {
            if (!checkIn || !checkOut || typesWithoutTime.includes(workType)) return { workHours: "-", overtimeHours: "-" };
            try {
                const start = new Date(`1970-01-01T${checkIn}`);
                const end = new Date(`1970-01-01T${checkOut}`);
                const restMinutes = timeStrToMinutes(restTimeStr);
                let diffMinutes = (end - start) / 60000;
                if (diffMinutes < 0) diffMinutes += 24 * 60;
                const netWorkMinutes = Math.max(0, diffMinutes - restMinutes);
                const netWorkHours = (netWorkMinutes / 60).toFixed(2);
                let baseWorkHours = 8;
                if (workType === "午前有給" || workType === "午後有給") baseWorkHours = 4;
                let overtime = 0;
                if (workType === "祝日出勤") overtime = netWorkMinutes;
                else if (["出勤", "在宅", "午前有給", "午後有給"].includes(workType)) overtime = Math.max(0, netWorkMinutes - baseWorkHours * 60);
                const overtimeH = (overtime / 60).toFixed(2);
                return { workHours: netWorkHours, overtimeHours: overtimeH };
            } catch (e) { return { workHours: "Error", overtimeHours: "Error" }; }
        }
    </script>
</body>
</html>
```

---

## 3. `static/style.css`

```css
/* --- General & Layout --- */
body {
    background-color: #f0f2f5;
    color: #333;
    font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    margin: 0;
    padding: 20px;
    box-sizing: border-box;
}

.container {
    width: 100%;
    max-width: 95%;
    margin: 0 auto;
}

header h1 {
    text-align: center;
    font-size: 2.2rem;
    color: #1c2e3f;
    margin-bottom: 25px;
}

main {
    display: flex;
    flex-direction: column;
    gap: 25px;
}

.top-row {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 25px;
}

.card {
    background-color: #ffffff;
    border-radius: 12px;
    padding: 25px;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.07);
    display: flex;
    flex-direction: column;
}

.card-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 20px;
}

.card-header h2 {
    margin: 0;
    border: none;
    padding: 0;
}

h2 {
    font-size: 1.4rem;
    color: #1c2e3f;
    margin-top: 0;
    margin-bottom: 20px;
    padding-bottom: 10px;
    border-bottom: 1px solid #e9ecef;
}

/* --- Buttons & Inputs --- */
.button-group button, .input-group button, .add-button, .calendar-nav button {
    padding: 10px 15px;
    border: none;
    border-radius: 6px;
    font-size: 0.95rem;
    font-weight: 500;
    cursor: pointer;
    transition: background-color 0.2s ease, transform 0.1s ease;
    background-color: #007bff;
    color: white;
}

.button-group button:hover, .input-group button:hover, .add-button:hover, .calendar-nav button:hover {
    background-color: #0056b3;
}

#check-in { background-color: #28a745; }
#check-out { background-color: #dc3545; }
#check-in:hover { background-color: #218838; }
#check-out:hover { background-color: #c82333; }

.add-button { padding: 5px 10px; font-size: 1.2rem; line-height: 1; }

.input-group { display: flex; flex-direction: column; gap: 10px; }
.input-group input, .input-group select, .input-group textarea { width: 100%; padding: 10px; border: 1px solid #ced4da; border-radius: 6px; font-size: 0.9rem; box-sizing: border-box; }


/* --- Calendar & Summary --- */
.calendar-nav { display: flex; justify-content: center; align-items: center; margin-bottom: 15px; }
#current-month-year { font-size: 1.6rem; font-weight: 600; color: #1c2e3f; margin: 0 20px; }
#calendar-container { overflow-x: auto; padding-bottom: 15px; }
#calendar-grid { display: grid; grid-auto-flow: column; grid-auto-columns: minmax(250px, 1fr); gap: 15px; }

.calendar-day {
    background-color: #f8f9fa;
    border: 1px solid #e9ecef;
    border-radius: 8px;
    min-height: 360px;
    padding: 10px;
    font-size: 0.85rem;
    display: flex;
    flex-direction: column;
    gap: 10px;
}

/* Day specific colors */
.calendar-day.saturday { background-color: #e0f2f7; border-color: #a7d9ed; }
.calendar-day.sunday { background-color: #ffe0e0; border-color: #ffb3b3; }
.calendar-day.holiday { background-color: #fff3e0; border-color: #ffcc80; }

.day-header { display: flex; justify-content: space-between; font-weight: 600; border-bottom: 1px solid #e0e0e0; padding-bottom: 8px; }
.day-number { font-size: 1.1rem; }
.day-of-week { color: #6c757d; }

.day-input { width: 100%; padding: 8px; border: 1px solid #ced4da; border-radius: 4px; font-size: 0.85rem; box-sizing: border-box; }
.time-inputs { display: flex; gap: 5px; align-items: center; }
.time-inputs input { flex: 1; min-width: 0; }
.rest-time-input { display: flex; gap: 5px; align-items: center; }
.calculated-times, .daily-task-summary { font-size: 0.85rem; display: flex; justify-content: space-between; padding-top: 8px; margin-top: 5px; border-top: 1px dashed #e0e0e0; }
.daily-task-summary { font-weight: bold; color: #0056b3; }

#monthly-summary { background-color: #e9ecef; padding: 15px; border-radius: 8px; margin-bottom: 20px; display: flex; justify-content: space-around; flex-wrap: wrap; gap: 15px; }
.summary-group { display: flex; flex-direction: column; gap: 5px; align-items: center; }
.summary-group h3 { font-size: 1rem; color: #495057; margin: 0 0 5px 0; }
.summary-group span { background: #fff; padding: 3px 8px; border-radius: 10px; font-size: 0.85rem; }

/* --- Subtasks & Announcements --- */
.subtask-section { margin-top: auto; }
.subtask-list { display: flex; flex-direction: column; gap: 5px; margin-top: 5px; }
.subtask-entry { display: flex; align-items: center; gap: 5px; }
.subtask-name { flex-grow: 1; font-size: 0.8rem; }
.subtask-time { width: 70px; min-width: 0; }
.delete-subtask { background: none; border: none; color: #dc3545; cursor: pointer; font-size: 1.2rem; line-height: 1; padding: 0 5px; }

#announcements-list { flex-grow: 1; max-height: 150px; overflow-y: auto; }
.announcement { padding: 10px; border-bottom: 1px solid #e9ecef; }
.announcement:last-child { border-bottom: none; }
.announcement h4 { margin: 0 0 5px 0; font-size: 1rem; }
.announcement p { margin: 0; font-size: 0.9rem; }

#defined-tasks-list { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
#defined-tasks-list .task-category h3 { font-size: 1.1rem; margin-bottom: 8px; color: #007bff; }
#defined-tasks-list ul { list-style: none; padding: 0; margin: 0; }
#defined-tasks-list li { display: flex; justify-content: space-between; align-items: center; padding: 8px; background-color: #f8f9fa; border-radius: 4px; margin-bottom: 5px; }
#defined-tasks-list button { background-color: #dc3545; color: white; padding: 3px 8px; font-size: 0.8rem; }

/* --- Modal --- */
.modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; overflow: auto; background-color: rgba(0,0,0,0.5); justify-content: center; align-items: center; }
.modal-content { background-color: #fff; margin: auto; padding: 30px; border-radius: 12px; box-shadow: 0 5px 15px rgba(0,0,0,0.3); width: 90%; max-width: 500px; position: relative; }
.modal-content.wide { max-width: 800px; }
.modal-scroll-content { max-height: 300px; overflow-y: auto; margin: 15px 0; padding: 10px; border: 1px solid #e9ecef; border-radius: 6px; }
.modal-content .input-group { margin-top: 20px; }
.close-button { color: #aaa; position: absolute; top: 10px; right: 20px; font-size: 28px; font-weight: bold; cursor: pointer; }
.close-button:hover, .close-button:focus { color: black; }

/* --- Announcements & Comments --- */
.announcement-item { padding: 10px; border-bottom: 1px solid #e9ecef; cursor: pointer; transition: background-color 0.2s; }
.announcement-item:last-child { border-bottom: none; }
.announcement-item:hover { background-color: #f8f9fa; }
.announcement-item h4 { margin: 0; font-size: 1rem; }

.comment { border-bottom: 1px solid #e9ecef; padding: 10px; }
.comment:last-child { border-bottom: none; }
.comment p { margin: 0 0 5px 0; }
.comment-date { font-size: 0.8rem; color: #6c757d; }

/* --- Select Option Hover Fix --- */
select option:hover {
    background-color: #007bff;
    color: #ffffff;
}
```