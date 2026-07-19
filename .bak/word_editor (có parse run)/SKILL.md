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

## `--read`: Trích xuất cấu trúc văn bản

Lưu cấu trúc văn bản của tệp Word ra file JSON và Markdown chỉ định (lược bỏ phần mở rộng đuôi tệp khi truyền vào đối số):

```powershell
python word_editor.py --read "sessions/<tên_phiên>/temp/filled_temp.docx" "sessions/<tên_phiên>/temp/word_structure"
```

Lệnh này sinh ra đồng thời hai file:

1. `word_structure.json`: Chứa cấu trúc JSON cho script ánh xạ tự động. **Không chứa comment** — script mapping không cần đọc comment, chỉ cần `text` gốc.
2. `word_structure.md`: Chứa outline dạng Markdown có cấu trúc phân cấp (các đoạn văn được ghi trực tiếp trên từng dòng và các bảng biểu dạng Markdown Table được đánh số chỉ số hàng/cột trực quan để xem nhanh), **có kèm `<run:...>` cho đoạn text có style khác biệt và comment marker nếu vị trí đó có comment**.

Agent nên đọc `.md` để hiểu nội dung, định vị vị trí cần điền (kể cả dựa vào style/màu chữ), và đọc yêu cầu từ comment (nếu có). File `.json` chỉ dùng khi viết script ánh xạ để sinh `changes.json`.

### Cấu trúc file JSON sinh ra:

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

## Run-level style trong file MD

Word chia mỗi paragraph/cell thành nhiều Run — mỗi Run mang một bộ định dạng riêng (font, màu, bold, italic...). Baseline mặc định của Word là: chữ đen, không bold, không italic, không underline, không highlight.

Nếu một Run có style khác baseline, file `.md` bọc đoạn text đó bằng `<run:...>...</run>`:

```
Kết nối và định tuyến động với <run:#FF0000>các Router của ISP</run>, dùng để kiểm soát truy cập
<run:b,i>Ghi chú quan trọng</run>
```

Style list trong `<run:...>` dùng dấu phẩy, gồm:

| Ký hiệu   | Ý nghĩa                   |
| --------- | ------------------------- |
| `b`       | bold                      |
| `i`       | italic                    |
| `u`       | underline                 |
| `hl`      | highlight (tô nền)        |
| `#RRGGBB` | màu chữ khác đen mặc định |

Run có style đúng baseline (đen, không bold/italic/underline/highlight) giữ nguyên text thuần, không bọc gì cả — phần lớn nội dung trong 1 văn bản hành chính là baseline, nên chỉ những đoạn thật sự đáng chú ý mới xuất hiện tag.

Theo Template Hints, chữ tô màu đỏ (`<run:#...>` với sắc đỏ) thường là gợi ý vị trí cần điền — nhưng chỉ là tín hiệu hỗ trợ, luôn đối chiếu với dữ liệu người dùng cung cấp trước khi quyết định thay.

**Không có trong `.json`** — Run tag chỉ dành cho Agent đọc hiểu. Script ánh xạ không cần và không được dùng thông tin style để quyết định nội dung `find`/`replace`.

**Khi ghi đè (`--apply`):** style tại vị trí đang thay được giữ nguyên tự động — Word áp style hiện có của Range lên text mới khi gán `.Text`, không tạo style riêng cho nội dung mới. Nếu `find` vắt qua ranh giới nhiều `<run>` có style khác nhau, Word có thể áp toàn bộ text mới theo style của Run đầu tiên trong vùng thay vì giữ đúng pattern nhiều style ban đầu — cứ tuân theo Minimum Edit Principle như bình thường, không cần mở rộng phạm vi `find` để né trường hợp này.

Nếu nội dung gốc của tài liệu tình cờ chứa chuỗi `<run:` hoặc `</run>`, script tự escape (chèn khoảng trắng vào giữa) khi xuất `.md`.

---

## Comment trong file MD

Nếu một paragraph hoặc cell có comment của người review đính kèm trong file Word gốc, file `.md` đánh dấu bằng cặp marker mang ID tăng dần (1, 2, 3... theo thứ tự comment xuất hiện trong document):

- `%%COMMENT:<id>%%<nội dung>` — luôn xuất hiện, nối vào cuối nội dung của paragraph/cell nơi comment **kết thúc**, kèm theo nội dung yêu cầu của người review.
- `%%COMMENT_START:<id>%%` — chỉ xuất hiện thêm, ở đầu nội dung, khi comment đó trải qua **nhiều hơn một** paragraph/cell. Nếu comment gọn trong đúng một paragraph/cell, không có marker mở — chỉ cần `%%COMMENT:<id>%%<nội dung>`.

Cùng một `<id>` ở `%%COMMENT_START%%` và `%%COMMENT%%` thuộc về cùng một comment — dùng ID để ghép đúng cặp mở/đóng khi có nhiều comment lồng nhau hoặc chồng phạm vi lên nhau. Mọi paragraph/cell nằm giữa `%%COMMENT_START:<id>%%` và `%%COMMENT:<id>%%` đều thuộc phạm vi của comment đó.

Ví dụ — comment gọn trong 1 cell:

```
| 3 | Firewall Fortigate FG 100F (SL: 02) %%COMMENT:1%%Xác nhận lại vị trí lắp đặt theo khảo sát thực tế | Vùng mạng nội bộ | ... |
```

Ví dụ — comment trải qua 2 paragraph liên tiếp:

```
%%COMMENT_START:2%%UBND XÃ KHÂM ĐỨC
Hệ thống thông tin mạng Internet phục vụ công tác quản lý %%COMMENT:2%%Xác nhận lại toàn bộ tên đơn vị và mục đích sử dụng theo khảo sát thực tế
```

**Quan trọng:** phần comment (`%%COMMENT...`) không bao giờ được đưa vào `changes.json` làm giá trị `find` — chỉ nội dung gốc (phần đứng trước `%%COMMENT_START` hoặc `%%COMMENT`) mới là text thật của tài liệu.

Nếu nội dung gốc của tài liệu tình cờ chứa chuỗi `%%COMMENT`, script tự escape (chèn khoảng trắng vào giữa) khi xuất `.md` để tránh nhầm lẫn với marker thật.

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

**Tự động thêm hàng khi thiếu chỗ:** Đây là ngoại lệ bất khả kháng, trong system prompt nêu không được phép thay đổi layout bảng, tuy nhiên nếu có nhiều thiết bị hơn, được phép thêm row. Cụ thể nếu `row_index` vượt quá số hàng hiện có của bảng, `--apply` tự gọi `Rows.Add()` để bảng đủ chỗ trước khi ghi, không báo lỗi và không cần rule riêng để "tạo hàng" — chỉ cần trỏ thẳng `row_index` tới vị trí mong muốn. Ví dụ bảng đang có 4 hàng (index 0-3) nhưng khảo sát có 10 thiết bị: cứ tạo rule với `row_index` chạy tới 9, `--apply` sẽ tự thêm 6 hàng còn thiếu. Số hàng đã thêm được log ra (`[INFO] Table[...]: Đã chèn thêm N hàng mới`).

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
- `replace` dài hơn 250 ký tự: không bị giới hạn (script dùng Range gán trực tiếp, không dùng Word Replace engine).
- `show_ui: true`: giữ Word mở sau khi thực thi (thành công hoặc lỗi) để kiểm tra thủ công. Mặc định `false`.
- Merged cell trong bảng: xem field `unreadable_cells` ở output `--read`.
- Comment: chỉ đọc được trong `--read` (xuất ra `.md`, không có trong `.json`). `--apply` không tương tác với comment — không tự động resolve/xóa comment sau khi điền dữ liệu.
- Run-level style (`<run:...>`): chỉ đọc được trong `--read` (xuất ra `.md`, không có trong `.json`). `Font.Color` dạng "tự động"/theme color không xác định được mã hex cụ thể — coi như baseline, không hiển thị tag màu.
