---
name: word_editor
description: Công cụ chỉnh sửa văn bản Word thông qua Microsoft Word COM API. Dùng skill này bất cứ khi nào cần điền dữ liệu vào file .docx. Cung cấp hai thao tác: (1) --read để đọc cấu trúc đoạn văn và bảng biểu của file, (2) --apply để thực thi thay thế văn bản từ file changes.json.
---

# word_editor

Script duy nhất dùng lại cho mọi phiên: `.agents/skills/word_editor/scripts/word_editor.py`

Agent sinh ra `changes.json` phù hợp với từng phiên, script chỉ là executor.

---

## Quy trình sử dụng

### Bước 1 — Đọc cấu trúc file

```powershell
python ".agents/skills/word_editor/scripts/word_editor.py" --read "sessions/<tên_phiên>/temp/filled_temp.docx"
```

Output JSON gồm:

- `paragraphs`: danh sách `{index, text}` — **chỉ dùng để xác định nội dung `anchor`/`find`, không dùng index để targeting**
- `tables`: danh sách bảng với từng ô `{row, col, text}` — index 0-based, dùng được cho `table_cell`

> **Giới hạn quan trọng**: `paragraph.index` từ `--read` (python-docx) KHÔNG tương đương với index COM trong `--apply` khi tài liệu có bảng biểu. Do đó `paragraph_index` **không được hỗ trợ** trong `--apply`. Chỉ dùng `anchor`.

### Bước 2 — Sinh `changes.json`

Lưu tại: `sessions/<tên_phiên>/changes.json`

```json
{
  "target_file": "sessions/<tên_phiên>/temp/filled_temp.docx",
  "show_ui": false,
  "replacements": [
    {
      "comment": "Mô tả mục đích — bắt buộc",
      "scope": "paragraph",
      "anchor": "Chuỗi nhận diện ngữ cảnh duy nhất của đoạn văn",
      "find": "Chuỗi cần tìm — phải khớp chính xác với text trong output --read",
      "replace": "Chuỗi thay thế",
      "max_replacements": 1
    },
    {
      "comment": "Ô bảng theo tọa độ — chính xác nhất",
      "scope": "table_cell",
      "table_index": 0,
      "row_index": 1,
      "col_index": 1,
      "find": "Hải Châu",
      "replace": "Khâm Đức",
      "max_replacements": 1
    }
  ]
}
```

### Bước 3 — Thực thi

```powershell
python ".agents/skills/word_editor/scripts/word_editor.py" --apply "sessions/<tên_phiên>/changes.json"
```

### Bước 4 — Tạo Track Changes

Gọi Skill `doc_diff` để tạo `.check.docx`.

---

## Quy tắc viết `changes.json`

**Thứ tự ưu tiên scope:**

```
table_cell  >  paragraph + anchor
```

**Bắt buộc:**

- `find` phải là chuỗi thực tế xác nhận từ output `--read`, không được tự bịa.
- `anchor` phải đủ phân biệt — không xuất hiện ở nhiều đoạn văn khác nhau.
- `max_replacements: 1` là mặc định (thay đúng 1 lần). Đặt `0` để thay tất cả trong Range. Đặt số nguyên N để thay tối đa N lần — count thực tế được in ra sau mỗi rule.
- Mỗi rule phải có `comment`.
- Không dùng `find` quá ngắn (dưới 3 ký tự).
- Tránh dùng ký tự `^` trong `find`/`replace` — Word sẽ diễn giải như ký tự đặc biệt.

**Giới hạn kỹ thuật:**

- `replace` dài hơn 250 ký tự: script tự động fallback sang `TypeText` thay vì Word Replace engine.
- Ô bảng merge: tọa độ `row/col` có thể không khớp với COM nếu có merged cells — kiểm tra kỹ bằng `--read`.
- `show_ui: true`: Word và document giữ mở sau khi thực thi để kiểm tra thủ công.
