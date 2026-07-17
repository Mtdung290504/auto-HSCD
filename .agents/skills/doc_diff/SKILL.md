---
name: doc_diff
description: Generate a native Microsoft Word Track Changes document by comparing an original document with an edited document. Use this skill whenever a reviewable Word document with Track Changes is required.
---

# So Sánh Tài Liệu Word

Skill này tạo file Track Changes gốc của Microsoft Word bằng cách so sánh tài liệu gốc với tài liệu đã chỉnh sửa, sử dụng tính năng Compare có sẵn của Microsoft Word.

## Đầu vào

Script cần:

- Tài liệu gốc (`.docx`)
- Tài liệu đã chỉnh sửa (`.docx`)
- Đường dẫn file output (`.check.docx`)

## Thực thi

Chạy script so sánh:

```bash
python ".agents/skills/doc_diff/scripts/doc_diff.py" "<file_docx_gốc>" "<file_docx_đã_sửa>" "<đường_dẫn_file_output_check_docx>"
```

## Output

Script tạo ra:

- một file `.check.docx`;
- Track Changes gốc của Microsoft Word;
- toàn bộ insertion, deletion và thay đổi định dạng do Microsoft Word phát hiện và xử lý.
