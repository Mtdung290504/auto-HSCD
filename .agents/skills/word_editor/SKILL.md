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

`Paragraphs` và `Tables` là hai collection độc lập trong Word COM, mỗi loại đánh số riêng theo đối tượng của nó. Một paragraph đứng trước/sau một bảng về mặt vị trí vật lý không có nghĩa là index của chúng so sánh được với nhau — `paragraph_index` và `table_index` chỉ có ý nghĩa trong phạm vi collection của chính nó. Không có field nào cho biết "đoạn văn nào đứng ngay trước/sau bảng này"; nếu cần quan hệ đó, phải tự suy luận qua nội dung `text` của cả hai danh sách, không dựa vào so sánh số index giữa hai loại.

**Phạm vi:** chỉ đọc/ghi nội dung trong main document body (đoạn văn + bảng, kể cả bảng lồng bảng). Không xử lý textbox, header/footer, footnote/endnote.

---

## `--read`: cấu trúc output

Output là JSON nén (không indent) để giảm token — parse trực tiếp, không cần định dạng đẹp.

```powershell
python word_editor.py --read "sessions/<tên_phiên>/temp/filled_temp.docx"
```

```json
{
  "paragraphs": ["Đoạn văn thứ 0...", "Đoạn văn thứ 1...", "..."],
  "tables": [
    {
      "rows": 4,
      "cols": 3,
      "cells": [[0, 0, "text A"], [0, 1, "text B"], ["..."]],
      "unreadable_cells": [[0, 2]]
    }
  ]
}
```

- `paragraphs`: mảng chuỗi thuần. Vị trí trong mảng **chính là** `paragraph_index` — không có field index riêng (`paragraphs[5]` → `paragraph_index: 5`).
- `tables[i].cells`: mảng bộ ba `[row, col, text]`, 0-based, chỉ gồm các cell đọc được.
- `tables[i].unreadable_cells`: mảng cặp `[row, col]` — **chỉ xuất hiện** nếu bảng đó có ít nhất một cell không đọc được. Không có field này nghĩa là mọi cell trong bảng đều đọc bình thường.
- Vị trí trong mảng `tables` **chính là** `table_index`.

Tọa độ xuất hiện trong `unreadable_cells` nghĩa là nó rơi vào vùng bị merge vào ô khác — không phải ô độc lập trong Word. **Không được** dùng tọa độ này làm `table_cell` trong `changes.json`; `--apply` sẽ báo lỗi và bỏ qua rule nếu target vào đó.

---

## `changes.json`: scope và field

**`scope: "table_cell"`** — dùng cho nội dung trong bảng, chính xác nhất vì định vị bằng tọa độ:

```json
{
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
- `paragraph_index`: dùng khi nội dung lặp lại nhiều nơi khiến `anchor` không đủ phân biệt. Lấy trực tiếp từ vị trí trong mảng `paragraphs` của output `--read`. `--apply` sẽ đối chiếu `find` có thực sự nằm trong đoạn đó không trước khi thay — nếu không khớp (file đã đổi từ lúc `--read`), rule bị bỏ qua có báo lỗi, không thay nhầm.

```json
{
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
- Merged cell trong bảng: xem field `unreadable_cells` ở output `--read`.
