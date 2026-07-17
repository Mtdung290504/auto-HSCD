"""
word_editor.py - Universal Word document editor
Part of the word_editor skill. Cách dùng, cấu trúc changes.json, và giới
hạn kỹ thuật: xem SKILL.md.

Usage:
  python word_editor.py --read  <file.docx>       # Đọc cấu trúc file
  python word_editor.py --apply <changes.json>    # Thực thi thay thế
"""

import sys
import os
import json

sys.stdout.reconfigure(encoding="utf-8")

# Giới hạn của Word Find/Replace engine — chỉ dùng để in cảnh báo thông tin,
# KHÔNG còn là giới hạn cứng vì _replace_in_range dùng TypeText, không dùng
# Find.Execute(Replace=...).
WORD_REPLACE_MAX_LEN = 250


# ===========================================================
# MODE: --read — đọc cấu trúc file qua Word COM
# ===========================================================
def cmd_read(file_path):
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        print("[INFO] Khởi động Microsoft Word để đọc (COM)...", file=sys.stderr)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(file_path, ReadOnly=True)

        result = {
            "file": file_path,
            "paragraph_count": doc.Paragraphs.Count,
            "table_count": doc.Tables.Count,
            "note": (
                "paragraph.index và table.index được đọc bằng Word COM — "
                "khớp trực tiếp với index dùng trong --apply. "
                "Ưu tiên anchor cho nội dung có thể lặp lại; dùng "
                "paragraph_index chỉ khi anchor không đủ phân biệt."
            ),
            "warning_scope": (
                "Không bao gồm text trong textbox, header/footer, "
                "footnote/endnote — ngoài phạm vi công việc, các story "
                "này nằm ngoài doc.Paragraphs/doc.Tables main story."
            ),
            "paragraphs": [],
            "tables": [],
        }

        # ---- Paragraphs (1-based COM -> xuất 0-based cho JSON) ----
        para_count = doc.Paragraphs.Count
        for i in range(1, para_count + 1):
            try:
                text = doc.Paragraphs(i).Range.Text
                # Word trả về ký tự pilcrow/control ở cuối Range.Text, strip cho sạch
                text = text.replace("\r", "").replace("\x07", "").strip()
            except Exception as e:
                text = f"[ERROR đọc paragraph {i}: {e}]"
            result["paragraphs"].append({"index": i - 1, "text": text})

        # ---- Tables (1-based COM -> xuất 0-based cho JSON) ----
        table_count = doc.Tables.Count
        for t in range(1, table_count + 1):
            table = doc.Tables(t)
            try:
                n_rows = table.Rows.Count
                n_cols = table.Columns.Count
            except Exception as e:
                result["tables"].append(
                    {"index": t - 1, "error": f"Không đọc được kích thước bảng: {e}"}
                )
                continue

            table_data = {"index": t - 1, "rows": n_rows, "cols": n_cols, "cells": []}

            for r in range(1, n_rows + 1):
                for c in range(1, n_cols + 1):
                    try:
                        cell = table.Cell(r, c)
                        cell_text = (
                            cell.Range.Text.replace("\r", "")
                            .replace("\x07", "")
                            .strip()
                        )
                        table_data["cells"].append(
                            {
                                "row": r - 1,
                                "col": c - 1,
                                "text": cell_text,
                                "merged_or_unreadable": False,
                            }
                        )
                    except Exception:
                        # Cell(r,c) ném lỗi -> tọa độ này rơi vào vùng merge
                        # hoặc không tồn tại độc lập. Đánh dấu rõ, KHÔNG suy đoán text.
                        table_data["cells"].append(
                            {
                                "row": r - 1,
                                "col": c - 1,
                                "text": None,
                                "merged_or_unreadable": True,
                            }
                        )

            result["tables"].append(table_data)

        print(json.dumps(result, ensure_ascii=False, indent=2))

    finally:
        if doc is not None:
            doc.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


# ===========================================================
# MODE: --apply — thực thi thay thế trong Range cụ thể qua Word COM
# ===========================================================
def cmd_apply(changes_file):
    import win32com.client
    import pythoncom

    with open(changes_file, "r", encoding="utf-8") as f:
        changes = json.load(f)

    target_file = os.path.abspath(changes["target_file"])
    replacements = changes.get("replacements", [])
    show_ui = changes.get("show_ui", False)

    if not os.path.exists(target_file):
        print(f"[ERROR] File không tồn tại: {target_file}")
        sys.exit(1)

    if not replacements:
        print("[WARN] Không có replacement nào trong changes.json.")
        return

    pythoncom.CoInitialize()
    word = None
    doc = None
    saved = False

    try:
        print("[INFO] Khởi động Microsoft Word...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = show_ui
        if show_ui:
            word.WindowState = 1  # wdWindowStateNormal

        print(f"[INFO] Mở file: {target_file}")
        doc = word.Documents.Open(target_file)

        total_replaced = 0

        for i, rule in enumerate(replacements):
            scope = rule.get("scope", "paragraph")
            find_text = rule.get("find", "")
            replace_text = rule.get("replace", "")
            max_replacements = rule.get("max_replacements", 1)
            comment = rule.get("comment", f"Rule {i}")

            if not find_text:
                print(f"  [SKIP] Rule {i}: 'find' rỗng. | {comment}")
                continue

            if "^" in find_text or "^" in replace_text:
                print(
                    f"  [WARN] Rule {i}: Chuỗi chứa '^' — Word sẽ diễn giải như ký tự đặc biệt. Kiểm tra kết quả. | {comment}"
                )

            target_range = None

            # --------------------------------------------------
            # Scope: table_cell — chính xác nhất cho nội dung dạng bảng
            # --------------------------------------------------
            if scope == "table_cell":
                t_idx = rule.get("table_index", 0) + 1
                r_idx = rule.get("row_index", 0) + 1
                c_idx = rule.get("col_index", 0) + 1
                try:
                    target_range = doc.Tables(t_idx).Cell(r_idx, c_idx).Range
                except Exception as e:
                    print(
                        f"  [ERROR] Rule {i}: Không lấy được Table[{t_idx-1}] R{r_idx-1}C{c_idx-1} "
                        f"(có thể là merged cell): {e} | {comment}"
                    )
                    continue

            # --------------------------------------------------
            # Scope: paragraph — anchor ưu tiên, fallback paragraph_index
            # --------------------------------------------------
            elif scope == "paragraph":
                anchor = rule.get("anchor")
                p_idx = rule.get("paragraph_index")

                if anchor:
                    found_para = None
                    for para in doc.Paragraphs:
                        para_text = para.Range.Text
                        if anchor in para_text and find_text in para_text:
                            found_para = para
                            break
                    if found_para is None:
                        print(
                            f"  [MISS] Rule {i}: Không tìm thấy anchor='{anchor[:50]}' | {comment}"
                        )
                        continue
                    target_range = found_para.Range

                elif p_idx is not None:
                    try:
                        target_range = doc.Paragraphs(p_idx + 1).Range
                    except Exception as e:
                        print(
                            f"  [ERROR] Rule {i}: paragraph_index={p_idx} không hợp lệ: {e} | {comment}"
                        )
                        continue
                    # find_text vẫn phải khớp trong đoạn này, không sửa mù theo index
                    if find_text not in target_range.Text:
                        print(
                            f"  [MISS] Rule {i}: paragraph_index={p_idx} không chứa find_text='{find_text[:50]}' "
                            f"— khả năng file đã bị thay đổi kể từ lúc --read. | {comment}"
                        )
                        continue

                else:
                    print(
                        f"  [ERROR] Rule {i}: Scope 'paragraph' cần 'anchor' hoặc 'paragraph_index'. | {comment}"
                    )
                    continue

            else:
                print(
                    f"  [ERROR] Rule {i}: Scope không hợp lệ '{scope}'. Dùng 'paragraph' hoặc 'table_cell'. | {comment}"
                )
                continue

            n = _replace_in_range(
                word_range=target_range,
                find_text=find_text,
                replace_text=replace_text,
                max_replacements=max_replacements,
                word_app=word,
            )

            if n > 0:
                limit_str = "ALL" if max_replacements == 0 else str(max_replacements)
                print(
                    f"  [OK]   Rule {i}: [max={limit_str}] '{find_text}' → '{replace_text}' ({n}x thực tế) | {comment}"
                )
                total_replaced += n
            else:
                print(
                    f"  [MISS] Rule {i}: '{find_text}' không tìm thấy trong target range | {comment}"
                )

        print(
            f"\n[INFO] Tổng replacements: {total_replaced}/{len(replacements)} rules khớp"
        )
        print("[INFO] Đang lưu file...")
        doc.Save()
        saved = True
        print(f"[SUCCESS] Đã lưu: {target_file}")

    except Exception as e:
        print(f"[FATAL] {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)

    finally:
        if doc is not None:
            if show_ui and not saved:
                print("[INFO] Lỗi trước khi save — giữ Word mở để kiểm tra.")
            elif show_ui and saved:
                print("[INFO] Word UI vẫn mở. Đóng thủ công khi xong.")
            else:
                doc.Close(False)
        if word is not None and not show_ui:
            word.Quit()
        pythoncom.CoUninitialize()


def _replace_in_range(
    word_range, find_text, replace_text, max_replacements=1, word_app=None
):
    if len(replace_text) > WORD_REPLACE_MAX_LEN:
        print(
            f"    [INFO] replace_text dài ({len(replace_text)} ký tự) — dùng TypeText, không bị giới hạn 255 ký tự."
        )

    find_obj = word_range.Find
    find_obj.ClearFormatting()
    find_obj.Text = find_text
    find_obj.Forward = True
    find_obj.Wrap = 0  # wdFindStop — không thoát khỏi Range
    find_obj.MatchCase = True
    find_obj.MatchWholeWord = False
    find_obj.MatchWildcards = False

    count = 0
    unlimited = max_replacements == 0

    while find_obj.Execute():
        word_app.Selection.TypeText(replace_text)
        count += 1
        if not unlimited and count >= max_replacements:
            break

    return count


def main():
    args = sys.argv[1:]

    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)

    mode = args[0]

    if mode == "--read":
        if len(args) < 2:
            print("[ERROR] --read cần đường dẫn file .docx")
            sys.exit(1)
        file_path = os.path.abspath(args[1])
        if not os.path.exists(file_path):
            print(f"[ERROR] File không tồn tại: {file_path}")
            sys.exit(1)
        cmd_read(file_path)

    elif mode == "--apply":
        if len(args) < 2:
            print("[ERROR] --apply cần đường dẫn changes.json")
            sys.exit(1)
        changes_file = os.path.abspath(args[1])
        if not os.path.exists(changes_file):
            print(f"[ERROR] File không tồn tại: {changes_file}")
            sys.exit(1)
        cmd_apply(changes_file)

    else:
        print(f"[ERROR] Mode không hợp lệ '{mode}'. Dùng --read hoặc --apply.")
        sys.exit(1)


if __name__ == "__main__":
    main()
