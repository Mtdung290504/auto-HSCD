# Mô tả

Prompt và skill tự động hóa điền hồ sơ cấp độ ATTT xuất ra dạng Word Track Changes để accept hoặc reject từng thay đổi cụ thể của Agent.

# Phụ thuộc

- Máy Windows có MS Word (để chạy COM API so sánh văn bản)
- Môi trường chạy Agent: Antigravity IDE
- Trình thông dịch Python 3.8+ (đã được cài đặt và cấu hình PATH)
- Thư viện python:
  - `python-docx` (để xử lý nội dung file Word `.docx`)
  - `openpyxl` (để đọc dữ liệu cấu hình hệ thống từ file Excel `.xlsx`)
  - `pywin32` (cung cấp thư viện `win32com` để gọi API Microsoft Word so sánh và tạo Track Changes)

# Cấu trúc thư mục

```
.agents/
  rules/
    docx_automation_rules.md   ← Quy tắc điều khiển toàn bộ hành vi Agent
  skills/
    word_editor/               ← Đọc cấu trúc & ghi nội dung vào .docx qua Word COM
    excel_parser/              ← Đọc file .xlsx xuất JSON + Markdown
    doc_diff/                  ← So sánh 2 file .docx, tạo Track Changes gốc của Word
Output/                        ← Kết quả cuối cùng (<tên gốc>.check.docx)
Templates/                     ← Mẫu tài liệu & file Excel khảo sát (Read-Only)
Sessions/                      ← File tạm theo từng phiên làm việc (Sessions/<tên_phiên>/temp/)
```

# Prompt kick-off

Sử dụng các prompt sau để yêu cầu Agent bắt đầu làm việc:

- Chạy file đơn (khuyến khích — file nào dài, nhiều thông tin nên chạy đơn):

  ```
  Hãy xử lý hồ sơ ATTT.

  Điền thông tin từ Nguồn dữ liệu vào Templates\<Thư mục đơn vị>\<tên file>.docx

  Nguồn dữ liệu:
  Templates\<tên file khảo sát>.xlsx

  Thực hiện toàn bộ quy trình theo System Prompt.
  ```

- Chạy toàn bộ file trong một thư mục (không khuyến khích — khó kiểm soát lỗi, nên chỉ dùng khi hồ sơ đơn giản):

  ```
  Hãy xử lý hồ sơ ATTT.

  Điền toàn bộ file trong thư mục:
  Templates\<Thư mục đơn vị>

  Nguồn dữ liệu:
  Templates\<tên file khảo sát>.xlsx

  Thực hiện toàn bộ quy trình theo System Prompt.
  ```

# Ghi chú (cập nhật)

- Mỗi lần chạy Agent tạo một phiên làm việc (`Sessions/<tên_phiên>/`) riêng, chứa toàn bộ file tạm. Nếu thất bại, giữ nguyên để debug; nếu thành công, Agent tự xóa.
- Thư mục `Templates\` là Read-Only — không chỉnh sửa file gốc trực tiếp.
- Kết quả luôn nằm trong `Output\` với tên `<tên gốc>.check.docx`.
- Nếu hồ sơ có nhiều file, nên chạy từng file lẻ để dễ kiểm soát chất lượng từng tài liệu.

# Dự kiến phát triển

- Thường thì xử lý cả bộ hồ sơ, nhưng đưa cho agent làm hết 1 lần thì khó kiểm soát lỗi quá.
- Sẽ lên kế hoạch & sửa trong prompt để agent xử lý tốt trên từng tài liệu lẻ và output các tài liệu để ở cùng 1 nơi mà không bị mâu thuẫn các file temp dù có không xóa temp nhằm mục đích debug.
