---
name: excel_reader
description: "Use this skill whenever the user asks to read, open, analyze, summarize, or inspect an Excel file (.xlsx). Triggers include: any mention of an Excel file, spreadsheet, workbook, '.xlsx', 'file excel', 'bảng tính', 'file .xlsx', or a request to understand what's inside an Excel file. Converts every visible sheet into a clean Markdown table so the content can be read directly. Do NOT use for .xls (legacy binary Excel), .csv, or for writing/editing Excel files."
---

# Skill Đọc File Excel

Đọc một file `.xlsx` nhỏ và chuyển từng sheet thành bảng Markdown, in ra stdout. Được thiết kế cho file nhỏ, cần đọc và hiểu toàn bộ nội dung — script không phân trang hay cắt bớt dữ liệu.

## Cách dùng

Chạy script, truyền vào đường dẫn tới file Excel. Đường dẫn script bên dưới là đường dẫn tương đối so với thư mục chứa skill này (dù skill được cài ở đâu) — hãy resolve theo vị trí đó, đừng hard-code đường dẫn tuyệt đối.

```bash
python ".agents/skills/excel_reader/scripts/excel_parser.py" "<đường_dẫn_file_excel>"
```

Mặc định, các sheet ẩn sẽ bị bỏ qua. Để hiển thị luôn cả sheet ẩn:

```bash
python ".agents/skills/excel_reader/scripts/excel_parser.py" "<đường_dẫn_file_excel>" --all
```

## Định dạng output

Stdout là Markdown, có cấu trúc như sau:

```
# EXCEL DATA DUMP: <tên_file>

## Sheet: <tên sheet>
| Header1 | Header2 | ... |
| --- | --- | --- |
| val | val | ... |

## Sheet: <sheet khác> (Hidden)
...
```

Một số lưu ý khi đọc output:

- `(Hidden)` sau tên sheet nghĩa là tab đó bị ẩn trong file gốc (chỉ xuất hiện khi chạy với `--all`).
- Header cột hiển thị dạng `Col 3` nghĩa là ô tiêu đề đó trống trong file gốc — cột vẫn có dữ liệu nhưng không có tên.
- Các ô bị merge (gộp ô) đã được "trải" ra, mỗi ô bên dưới vùng merge đều hiện đúng giá trị đã gộp, thay vì hiện trống.
- Ngày tháng được định dạng `YYYY-MM-DD` (hoặc `YYYY-MM-DD HH:MM:SS` nếu có kèm giờ).
- `*Sheet is empty.*` / `*Sheet has only empty cells.*` nghĩa là sheet đó không có dữ liệu sử dụng được.
- Các dòng/cột trống ở cuối bảng được tự động cắt bỏ.

## Lỗi

- Không tìm thấy file hoặc không đọc được (file hỏng, sai định dạng) → script in `Error: ...` và thoát với exit code 1. Không có output nào được in ra trong trường hợp này.
- Nếu một sheet bị lỗi khi parse, sheet đó sẽ in thông báo `*Error reading this sheet: ...*` ngay tại vị trí của nó, nhưng các sheet còn lại vẫn được in bình thường.

## Giới hạn

- Không hỗ trợ `.xls` (định dạng nhị phân cũ) — chỉ hỗ trợ `.xlsx`.
- Không giới hạn số dòng/cột — script dành cho file nhỏ, đủ để nằm gọn trong context. Với bảng tính rất lớn, script sẽ in ra toàn bộ mà không cắt bớt.
