# 코드 변경 내역 비교 (2025-08-05)

이 문서는 최근 요청하신 기능 추가에 따른 코드 변경 사항을 **변경 전(Before)**과 **변경 후(After)**로 나누어 상세히 비교 설명합니다.

## 🚀 추가된 주요 기능

1.  **백엔드 알림**: 백엔드에서 프론트엔드로 모달 알림창을 띄웁니다.
2.  **공지사항 상세 및 댓글**: 공지사항 클릭 시 상세 내용과 댓글을 보고 작성할 수 있습니다.
3.  **사용자 이름 설정**: 최초 1회 사용자 이름을 입력받아 댓글 작성자로 사용합니다.
4.  **휴일 자동 설정**: DB에 데이터가 없는 주말/공휴일은 "休日"로 자동 표시합니다.

---

## 📄 변경된 파일 목록

-   `app_access.py` (백엔드)
-   `templates/index.html` (프론트엔드)
-   `static/style.css` (스타일)

---

## ✨ `app_access.py` 변경 상세

### 1. 모듈 임포트

-   **변경 전**
    ```python
    import jpholiday
    ```
-   **변경 후**
    ```python
    import jpholiday
    import calendar
    ```
-   **사유**: 월의 마지막 날을 계산하여 모든 날짜를 순회하기 위해 `calendar` 모듈을 추가했습니다.

### 2. 데이터베이스 테이블 생성 (`_create_tables`)

-   **변경 전**
    ```python
    self._execute("""
        CREATE TABLE Announcements (
            ID AUTOINCREMENT PRIMARY KEY,
            # ...
        );
    """)
    ```
-   **변경 후**
    ```python
    self._execute(""" # Announcements 테이블 (기존과 동일)
        # ...
    """)
    self._execute(""" # Comments 테이블 추가
        CREATE TABLE Comments (
            ID AUTOINCREMENT PRIMARY KEY,
            AnnouncementID LONG,
            AuthorName TEXT(100),
            CommentText MEMO,
            CommentDate DATE
        );
    """)
    self._execute(""" # Users 테이블 추가
        CREATE TABLE Users (
            EmployeeID TEXT(50) PRIMARY KEY,
            UserName TEXT(100)
        );
    """)
    ```
-   **사유**: 댓글과 사용자 이름을 영구적으로 저장하기 위해 `Comments`와 `Users` 테이블을 새로 추가했습니다.

### 3. `DatabaseManager` 클래스

-   **변경 내용**: `add_announcement` 함수 이후에 사용자, 공지사항 상세, 댓글 관련 CRUD 함수 4개를 새로 추가했습니다.
    ```python
    # ... add_announcement(...) 함수 끝

    # vvvvvv 추가된 함수들 vvvvvv
    def get_user_name(self, employee_id):
        # ...

    def set_user_name(self, employee_id, user_name):
        # ...

    def get_announcement_details(self, announcement_id):
        # ...

    def add_comment(self, announcement_id, author_name, comment_text, comment_date):
        # ...
    # ^^^^^^ 추가된 함수들 ^^^^^^

    def shutdown(self):
        # ...
    ```
-   **사유**: 새로 추가된 테이블과 기능에 필요한 데이터베이스 작업을 수행하기 위함입니다.

### 4. `Backend` 클래스

-   **변경 내용**: Signal 추가, `user_name` 속성 추가, Slot 함수 추가 등 클래스 전반이 수정되었습니다.

-   **변경 전**
    ```python
    class Backend(QObject):
        dataLoaded = Signal(dict)
        # ...
        showEmployeeIdPrompt = Signal()

        def __init__(self):
            # ...
            self.employee_id = None

        @Slot(str)
        def setEmployeeId(self, employee_id):
            self.employee_id = employee_id
            self.load_and_emit_employee_data()
            # ...

        def load_and_emit_employee_data(self):
            # ... DB에서 데이터 로드 후 바로 emit
        
        # ... (기존 Slot 함수들)
    ```

-   **변경 후**
    ```python
    class Backend(QObject):
        # ... (기존 Signal)
        # vvvvvv 추가된 Signal vvvvvv
        showAlert = Signal(str)
        announcementDetailsLoaded = Signal(dict)
        userNameRequired = Signal()
        # ^^^^^^ 추가된 Signal ^^^^^^

        def __init__(self):
            # ...
            self.user_name = None # 사용자 이름 속성 추가

        @Slot(str)
        def setEmployeeId(self, employee_id):
            self.employee_id = employee_id
            self.user_name = self.db_manager.get_user_name(self.employee_id) # 사용자 이름 조회
            self.load_and_emit_employee_data()
            # ...

        def load_and_emit_employee_data(self):
            # ... (DB 데이터 로드)
            # vvvvvv 휴일 자동 설정 로직 추가 vvvvvv
            for day in range(1, calendar.monthrange(year, month)[1] + 1):
                # ... (날짜 순회)
                if date_str not in attendance_data:
                    if day_of_week >= 5 or date_str in month_holidays:
                        attendance_data[date_str] = {'work_type': '休日', # ... }
            # ^^^^^^ 휴일 자동 설정 로직 ^^^^^^
            self.dataLoaded.emit(employee_data)
        
        # ... (기존 Slot 함수들)

        # vvvvvv 추가된 Slot vvvvvv
        @Slot(int)
        def getAnnouncementDetails(self, announcement_id):
            # ...

        @Slot(str)
        def setUserName(self, user_name):
            # ...

        @Slot(int, str)
        def addComment(self, announcement_id, comment_text):
            # ...
        # ^^^^^^ 추가된 Slot ^^^^^^
    ```
-   **사유**: 프론트엔드와 새로운 기능(알림, 댓글 등)을 연동하고, 휴일 자동 설정 로직을 수행하기 위해 백엔드 클래스를 대폭 확장했습니다.

---

## ✨ `templates/index.html` 변경 상세

이 파일은 UI와 프론트엔드 로직을 모두 포함하고 있어, 변경 범위가 매우 넓습니다. 핵심적인 변경 영역은 다음과 같습니다.

### 1. 새로운 모달(Modal) UI 추가

-   **변경 내용**: 기존 `employee-id-modal` 외에 4개의 모달이 `<body>` 내에 추가되었습니다.
    ```html
    <!-- 기존 모달 -->
    <div id="employee-id-modal" class="modal" style="display: flex;"> ... </div>

    <!-- vvvvvv 추가된 모달들 vvvvvv -->
    <div id="user-name-modal" class="modal"> ... </div>
    <div id="alert-modal" class="modal"> ... </div>
    <div id="announcement-detail-modal" class="modal"> ... </div>
    <div id="announcement-create-modal" class="modal"> ... </div>
    <!-- ^^^^^^ 추가된 모달들 ^^^^^^ -->
    ```

### 2. `<script>` 태그 로직 전체 변경

-   **변경 내용**: 기존 스크립트가 새로운 기능(모달 관리, Signal/Slot 연동, 댓글 처리 등)을 모두 처리하기 위해 전체적으로 재구성되고 확장되었습니다. 변경 전후를 직접 비교하기보다는, 추가된 주요 함수 블록을 설명하는 것이 더 명확합니다.

-   **주요 추가 함수 및 로직**:
    -   `setupModal(...)`: 여러 모달을 쉽게 관리하기 위한 헬퍼 함수.
    -   `backend.showAlert.connect(showAlert)`: 백엔드 알림 신호를 받아 `showAlert` 함수를 실행.
    -   `backend.announcementDetailsLoaded.connect(showAnnouncementDetails)`: 공지 상세 정보 신호를 받아 `showAnnouncementDetails` 함수 실행.
    -   `backend.userNameRequired.connect(...)`: 사용자 이름 요청 신호를 받아 이름 입력 모달을 표시.
    -   `submitUserName()`, `submitComment()`: 새로 추가된 버튼의 이벤트 리스너 함수.
    -   `showAnnouncementDetails(details)`: 백엔드에서 받은 데이터로 공지 상세 모달의 내용을 채우는 함수.
    -   `renderAnnouncements(announcements)`: 공지 목록의 각 항목에 `click` 이벤트를 추가하여 상세 정보 요청을 보내도록 수정됨.

---

## ✨ `static/style.css` 변경 상세

### 1. 모달 및 댓글 스타일 추가

-   **변경 전**
    ```css
    /* --- Modal --- */
    .modal { /* ... */ }
    .modal-content { /* ... */ }
    /* ... */
    ```
-   **변경 후**
    ```css
    /* --- Modal --- */
    .modal { /* ... */ }
    .modal-content { /* ... */ }
    .modal-content.wide { max-width: 800px; } /* 넓은 모달용 스타일 추가 */
    .modal-scroll-content { /* 스크롤 가능 영역 스타일 추가 */ }
    /* ... */

    /* --- Announcements & Comments --- */
    .announcement-item { /* 클릭 가능한 공지 아이템 스타일 추가 */ }
    .comment { /* 댓글 스타일 추가 */ }
    ```
-   **사유**: 새로 추가된 넓은 모달(공지 상세)과 스크롤이 필요한 콘텐츠 영역, 그리고 댓글 목록의 디자인을 위해 새로운 CSS 클래스를 추가했습니다.
