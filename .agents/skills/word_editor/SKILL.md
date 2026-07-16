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
- `paragraphs`: danh sách `{index, text}` — **index 0-based**
- `tables`: danh sách bảng với từng ô `{row, col, text}` — **index 0-based**

Dùng output này để xác định chính xác `find`, `anchor`, `paragraph_index`, `table_index/row_index/col_index` trước khi viết `changes.json`.

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
      "find": "Chuỗi cần tìm — phải khớp chính xác kết quả --read",
      "replace": "Chuỗi thay thế",
      "max_replacements": 1
    },
    {
      "comment": "Ô bảng — chính xác nhất",
      "scope": "table_cell",
      "table_index": 0,
      "row_index": 1,
      "col_index": 1,
      "find": "Hải Châu",
      "replace": "Khâm Đức",
      "max_replacements": 1
    },
    {
      "comment": "Fallback khi không có anchor phân biệt",
      "scope": "paragraph",
      "paragraph_index": 9,
      "find": "Chuỗi cần tìm",
      "replace": "Chuỗi thay thế",
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
table_cell  >  paragraph + anchor  >  paragraph + paragraph_index
```

**Bắt buộc:**
- `find` phải là chuỗi thực tế xác nhận từ output `--read`, không được tự bịa.
- `anchor` phải đủ phân biệt — không xuất hiện ở nhiều đoạn văn khác nhau.
- `max_replacements: 1` là mặc định — chỉ tăng khi có lý do rõ ràng.
- Mỗi rule phải có `comment`.
- Không dùng `find` quá ngắn (dưới 3 ký tự).

**Lưu ý kỹ thuật:**
- Replacement được thực thi trong **Range cụ thể** của từng đoạn/ô — không phải toàn file.
- `MatchCase: true` — phân biệt hoa thường.
- Nếu lỗi xảy ra trước `doc.Save()`, file gốc không bị thay đổi.
