# Mô tả

Prompt và skill tự động hóa điền hồ sơ cấp độ ATTT xuất ra dạng Word Track Changes để accept hoặc reject từng thay đổi cụ thể của Agent.

# Lưu ý quan trọng

Agent không sửa trang bìa vì nội dung nhập nhằng, người điền tự sửa

# Phụ thuộc

- Máy Windows có MS Word (để chạy COM API so sánh văn bản)
- Môi trường chạy Agent: Antigravity IDE
- Trình thông dịch Python 3.8+ (đã được cài đặt và cấu hình PATH)
- Thư viện python:
  - `python-docx` (để xử lý nội dung file Word `.docx`)
  - `openpyxl` (để đọc dữ liệu cấu hình hệ thống từ file Excel `.xlsx`)
  - `pywin32` (cung cấp thư viện `win32com` để gọi API Microsoft Word so sánh và tạo Track Changes)

# TODO

Sau khi clone, tạo thư mục Output\ và Templates\, trong đó:

- Template chứa đống mẫu, excel khảo sát
- Output chứa kết quả từ Agent

# Prompt kick-off

- Sử dụng các prompt sau để yêu cầu Agent bắt đầu làm việc:
  - Chạy file đơn (khuyến khích file nào dài, nhiều thì chạy đơn):

    ```
    Hãy xử lý hồ sơ ATTT.

    Điền thông tin từ Nguồn dữ liệu vào Templates\UBND Phuong Hai Chau\2. HSDXCD UBNDPHC_HTTTmangInternet.docx

    Nguồn dữ liệu:
    Templates\Mạng nội bộ UBND Xã Khâm Đức.xlsx

    Thực hiện toàn bộ quy trình theo System Prompt.
    ```

  - Chạy toàn bộ file trong thư mục gì đó (không khuyến khích, nên chạy đơn, hoặc nhiều file có chỉ định, chạy hết thư mục 1 lần hơi...)

    ```
    Hãy xử lý hồ sơ ATTT.

    Điền toàn bộ file trong thư mục:
    Templates\UBND Phuong Hai Chau

    Nguồn dữ liệu:
    Templates\Mạng nội bộ UBND Xã Khâm Đức.xlsx

    Thực hiện toàn bộ quy trình theo System Prompt.
    ```
