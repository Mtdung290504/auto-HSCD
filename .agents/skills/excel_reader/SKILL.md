---
name: excel_reader
description: "Use this skill whenever the user asks to read, open, analyze, summarize, or inspect an Excel file (.xlsx). Triggers include: any mention of an Excel file, spreadsheet, workbook, '.xlsx', 'file excel', 'bảng tính', 'file .xlsx', or a request to understand what's inside an Excel file. Converts every visible sheet into a clean Markdown table so the content can be read directly. Do NOT use for .xls (legacy binary Excel), .csv, or for writing/editing Excel files."
---

# Skill Đọc File Excel

Đọc một file `.xlsx` nhỏ và chuyển từng sheet thành bảng Markdown, in ra stdout. Được thiết kế cho file nhỏ, cần đọc và hiểu toàn bộ nội dung — script không phân trang hay cắt bớt dữ liệu.

## Cách dùng

Chạy script, truyền vào đường dẫn tới file Excel. Đường dẫn script bên dưới là đường dẫn tương đối so với thư mục chứa skill này.

1. **Chuyển đổi sang JSON (xuất ra file) - Mặc định**:

```bash
python ".agents/skills/excel_reader/scripts/excel_parser.py" "<đường_dẫn_file_excel>" "<đường_dẫn_file_ra.json>" [--all]
```

2. **Chuyển đổi sang Markdown (in ra stdout)**:

```bash
python ".agents/skills/excel_reader/scripts/excel_parser.py" "<đường_dẫn_file_excel>" --md [--all]
```

## Định dạng output

### 1. Định dạng JSON (Mặc định)

Lưu dữ liệu ra tệp JSON dưới dạng grid mảng hai chiều:

```json
{
  "sheets": {
    "Tên Sheet": {
      "hidden": false,
      "grid": [
        ["Giá trị ô chủ", [row_idx, col_idx]],
        ["Giá trị ô khác", ""]
      ]
    }
  }
}
```

- **Tối ưu hóa ô gộp (Merged Cells)**: Để tránh trùng lặp dữ liệu chuỗi dài, chỉ có ô góc trên bên trái (ô chủ) lưu chuỗi giá trị thực. Tất cả các ô phụ còn lại trong vùng merge sẽ lưu một mảng tọa độ 0-based `[row, col]` chỉ đến ô chủ.
- Khi phân tích bằng Python, bất kỳ ô nào thuộc kiểu `list`/`array` trong grid đều là ô tham chiếu trỏ đến ô chủ. Cấu trúc tham chiếu này có thể được đọc dễ dàng:
  ```python
  cell = grid[r][c]
  if isinstance(cell, list):
      ref_row, ref_col = cell
      val = grid[ref_row][ref_col]
  ```
- `*Sheet is empty.*` / `*Sheet has only empty cells.*` nghĩa là sheet đó không có dữ liệu sử dụng được.
- Các dòng/cột trống ở cuối bảng được tự động cắt bỏ.

### 2. Định dạng Markdown (khi dùng `--md`)

Stdout là Markdown, có cấu trúc như sau:

```
# EXCEL DATA DUMP: <tên_file>

## Sheet: <tên sheet>
| Header1 | Header2 | ... |
| --- | --- | --- |
| val | val | ... |
```

- Các ô bị merge (gộp ô) được tự động giải quyết trực tiếp để hiển thị đầy đủ chuỗi ký tự của ô chủ (ô góc trên bên trái).
- Ngày tháng được định dạng `YYYY-MM-DD` (hoặc `YYYY-MM-DD HH:MM:SS` nếu có kèm giờ).

## Lỗi

- Không tìm thấy file hoặc không đọc được (file hỏng, sai định dạng) → script in `Error: ...` và thoát với exit code 1. Không có output nào được in ra trong trường hợp này.
- Nếu một sheet bị lỗi khi parse, sheet đó sẽ in thông báo `*Error reading this sheet: ...*` ngay tại vị trí của nó, nhưng các sheet còn lại vẫn được in bình thường.

## Giới hạn

- Không hỗ trợ `.xls` (định dạng nhị phân cũ) — chỉ hỗ trợ `.xlsx`.
- Không giới hạn số dòng/cột — script dành cho file nhỏ, đủ để nằm gọn trong context. Với bảng tính rất lớn, script sẽ in ra toàn bộ mà không cắt bớt.
