---
name: word_editor
description: Tool for editing Word documents via the Microsoft Word COM API. Use this skill whenever you need to fill data into a .docx file. Provides two operations: (1) --read to read the paragraph and table structure of the file, (2) --apply to execute text replacements from a changes.json file.
---

# word_editor

Script: `.agents/skills/word_editor/scripts/word_editor.py`

Agent sinh `changes.json` phù hợp với từng phiên; script chỉ là executor.

---

## Cơ chế kỹ thuật

`--read` và `--apply` đều dùng Word COM (`doc.Paragraphs`, `doc.Tables`) — cùng một engine, cùng một cách đánh số. Index xuất ra từ `--read` khớp trực tiếp với index dùng trong `--apply`, không cần đối chiếu qua nguồn thứ hai.

`--read` chậm hơn đọc file trực tiếp vì phải khởi động Word, nhưng đổi lại loại bỏ hoàn toàn rủi ro lệch index giữa đọc và ghi.

**Phạm vi:** chỉ đọc/ghi nội dung trong main document body (đoạn văn + bảng, kể cả bảng lồng bảng). Không xử lý textbox, header/footer, footnote/endnote.

---

## `--read`: cấu trúc output

```powershell
python word_editor.py --read "sessions/<tên_phiên>/temp/filled_temp.docx"
```

```json
{
  "paragraph_count": 120,
  "table_count": 3,
  "paragraphs": [{ "index": 0, "text": "..." }],
  "tables": [
    {
      "index": 0,
      "rows": 4,
      "cols": 3,
      "cells": [
        { "row": 0, "col": 0, "text": "...", "merged_or_unreadable": false },
        { "row": 0, "col": 1, "text": null, "merged_or_unreadable": true }
      ]
    }
  ]
}
```

`merged_or_unreadable: true` nghĩa là tọa độ đó rơi vào vùng bị merge vào ô khác — không phải ô độc lập trong Word. **Không được** dùng tọa độ này làm `table_cell` trong `changes.json`; `--apply` sẽ báo lỗi và bỏ qua rule nếu target vào đó.

---

## `changes.json`: scope và field

**`scope: "table_cell"`** — dùng cho nội dung trong bảng, chính xác nhất vì định vị bằng tọa độ:

```json
{
  "comment": "Mô tả mục đích — bắt buộc",
  "scope": "table_cell",
  "table_index": 0,
  "row_index": 1,
  "col_index": 1,
  "find": "Hải Châu",
  "replace": "Khâm Đức",
  "max_replacements": 1
}
```

**`scope: "paragraph"`** — dùng cho đoạn văn ngoài bảng. Hỗ trợ hai cách định vị, chọn một:

- `anchor`: tìm đoạn văn chứa cả `anchor` và `find`. Ưu tiên dùng khi có, vì bền hơn trước các chỉnh sửa xảy ra giữa lúc `--read` và `--apply`. **Chỉ lấy kết quả khớp đầu tiên** — nếu `anchor` xuất hiện ở nhiều đoạn, phải viết `anchor` đủ dài/đặc trưng để không trùng.
- `paragraph_index`: dùng khi nội dung lặp lại nhiều nơi khiến `anchor` không đủ phân biệt. Lấy trực tiếp từ field `index` trong output `--read`. `--apply` sẽ đối chiếu `find` có thực sự nằm trong đoạn đó không trước khi thay — nếu không khớp (file đã đổi từ lúc `--read`), rule bị bỏ qua có báo lỗi, không thay nhầm.

```json
{
  "comment": "Mô tả mục đích — bắt buộc",
  "scope": "paragraph",
  "anchor": "Chuỗi ngữ cảnh duy nhất trong đoạn văn",
  "find": "Chuỗi cần tìm — khớp chính xác với text trong output --read",
  "replace": "Chuỗi thay thế",
  "max_replacements": 1
}
```

**Thứ tự ưu tiên chọn scope:** `table_cell` > `paragraph` + `anchor` > `paragraph` + `paragraph_index`.

---

## Field `max_replacements`

Số lần thay tối đa trong phạm vi Range đã xác định (không phải toàn file):

- `1` (mặc định): thay đúng 1 lần.
- `0`: thay tất cả các lần khớp trong Range.
- `N` (số nguyên > 1): thay tối đa N lần.

Count thực tế được in ra sau mỗi rule (`(Nx thực tế)`), dùng để xác nhận không có sai lệch giữa kỳ vọng và thực thi.

---

## Giới hạn kỹ thuật

- `find`/`replace` chứa ký tự `^`: Word Find engine diễn giải `^p`, `^t`... như mã đặc biệt bất kể chế độ MatchWildcards. Script in `[WARN]` khi phát hiện — tự kiểm tra kết quả nếu bắt buộc phải dùng.
- `replace` dài hơn 250 ký tự: không bị giới hạn (script dùng `Selection.TypeText`, không dùng Word Replace engine trực tiếp).
- `show_ui: true`: giữ Word mở sau khi thực thi (thành công hoặc lỗi) để kiểm tra thủ công. Mặc định `false`.
- Merged cell trong bảng: xem field `merged_or_unreadable` ở output `--read`.
