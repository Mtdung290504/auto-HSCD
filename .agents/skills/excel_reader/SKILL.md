---
name: excel_reader
description: Triggers when the user asks to read, analyze, process, or inspect an Excel file (.xlsx or .xls) to extract data.
---

# Excel Reader Skill

This skill allows you to automatically read and parse any Excel spreadsheet file (.xlsx or .xls) by converting all of its sheets into clean, readable Markdown tables.

## How to use:
Run the internal python parser script using the `run_command` tool:
```bash
# Mặc định chỉ hiển thị các tab (sheet) đang hiển thị (Visible)
python "c:/z-Information Security Level Profile/.agents/skills/excel_reader/scripts/excel_parser.py" "<path_to_excel_file>"

# Để hiển thị cả các tab bị ẩn (Hidden)
python "c:/z-Information Security Level Profile/.agents/skills/excel_reader/scripts/excel_parser.py" "<path_to_excel_file>" --all
```

## Lưu ý về Encoding (Tiếng Việt):
Khi chạy lệnh trên môi trường Windows PowerShell, luôn set biến môi trường UTF-8 trước khi gọi python để tránh lỗi `UnicodeEncodeError`:
```powershell
$env:PYTHONIOENCODING="utf-8"
```
Đoạn mã python đã được tích hợp sẵn `sys.stdout.reconfigure(encoding='utf-8')` để bảo đảm an toàn đầu ra tiếng Việt.
