# 코드 변경 내역 (2025-08-05)

## 변경된 파일
- `C:\Users\Chan\Desktop\attendance_app\templates\index.html`

## 변경 내용
캘린더에서 근무 유형을 **"오전유급"** 또는 **"오후유급"**으로 선택했을 때, 해당 날짜의 **휴식 시간이 자동으로 "00:00"으로 설정**되도록 JavaScript 코드를 수정했습니다.

### 변경 전 (Before)
```javascript
function handleDayDataChange(event, dateStr) {
    const dayCell = document.querySelector(`.calendar-day[data-date='${dateStr}']`);
    let new_data = {};
    dayCell.querySelectorAll('.day-input[data-field]').forEach(input => { new_data[input.dataset.field] = input.value; });
    const subtaskEntries = dayCell.querySelectorAll('.subtask-entry');
    new_data.subtasks = Array.from(subtaskEntries).map(entry => ({ name: entry.dataset.taskName, time: entry.querySelector('.subtask-time').value }));
    backend.updateDayData(dateStr, new_data);
}
```

### 변경 후 (After)
```javascript
function handleDayDataChange(event, dateStr) {
    const dayCell = document.querySelector(`.calendar-day[data-date='${dateStr}']`);

    // If the work type was changed to AM/PM leave, set rest time to 00:00
    if (event.target.classList.contains('work-type')) {
        const workType = event.target.value;
        if (workType === "午前有給" || workType === "午後有給") {
            const restTimeInput = dayCell.querySelector('.rest-time');
            if (restTimeInput) {
                restTimeInput.value = "00:00";
            }
        }
    }

    let new_data = {};
    dayCell.querySelectorAll('.day-input[data-field]').forEach(input => { new_data[input.dataset.field] = input.value; });
    const subtaskEntries = dayCell.querySelectorAll('.subtask-entry');
    new_data.subtasks = Array.from(subtaskEntries).map(entry => ({ name: entry.dataset.taskName, time: entry.querySelector('.subtask-time').value }));
    backend.updateDayData(dateStr, new_data);
}
```

## 기대 효과
사용자가 캘린더에서 특정 날짜의 근무 유형을 "오전유급" 또는 "오후유급"으로 변경하면, 휴식 시간 입력 필드의 값이 즉시 "00:00"으로 자동 변경되어 사용자의 편의성이 향상됩니다.

```