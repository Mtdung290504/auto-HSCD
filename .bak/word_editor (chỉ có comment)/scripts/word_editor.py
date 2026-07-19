"""
word_editor.py - Universal Word document editor
Part of the word_editor skill. Cách dùng, cấu trúc changes.json, và giới
hạn kỹ thuật: xem SKILL.md.

Usage:
  python word_editor.py --read  <file.docx> <output.json> # Đọc cấu trúc file và ghi ra JSON
  python word_editor.py --apply <changes.json>            # Thực thi thay thế
"""

import sys
import os
import json

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

# Giới hạn của Word Find/Replace engine — chỉ dùng để in cảnh báo thông tin,
# KHÔNG còn là giới hạn cứng vì _replace_in_range dùng Range.Text, không dùng
# Find.Execute(Replace=...).
WORD_REPLACE_MAX_LEN = 250

# Marker dùng để đánh dấu comment trong file .md — xem SKILL.md mục "Comment
# trong file MD" để biết quy ước đầy đủ.
COMMENT_MARK = "%%COMMENT"  # tiền tố chung, dùng để escape nội dung gốc
COMMENT_START_TPL = "%%COMMENT_START:{id}%%"
COMMENT_END_TPL = "%%COMMENT:{id}%%"


def _escape_comment_marks(text):
    """Nếu nội dung gốc tài liệu tình cờ chứa chuỗi marker, phá vỡ nó bằng
    cách chèn khoảng trắng, để không bị hiểu nhầm là comment marker thật khi
    agent đọc lại file .md."""
    return text.replace(COMMENT_MARK, "%% COMMENT")


def _clean_raw_text(text):
    """Xử lý ký tự điều khiển chung cho mọi text đọc từ Word COM."""
    return (text or "").replace("\r", "").replace("\x07", "").replace("\x0b", " ")


# ===========================================================
# MODE: --read — helpers cho comment
# ===========================================================
def _extract_comments(doc):
    """Trả về danh sách comment gốc: [{"id", "text", "start", "end"}, ...],
    id đánh số 1..N theo thứ tự xuất hiện trong document (không dùng
    Comment.Index trực tiếp để tránh phụ thuộc thứ tự nội bộ của Word)."""
    comments = []
    for idx, cmt in enumerate(doc.Comments, start=1):
        try:
            text = cmt.Range.Text.replace("\r", " ").replace("\x07", " ").strip()
            comments.append(
                {
                    "id": idx,
                    "text": text,
                    "start": cmt.Scope.Start,
                    "end": cmt.Scope.End,
                }
            )
        except Exception as e:
            print(f"[WARN] Không đọc được comment #{idx}: {e}", file=sys.stderr)
    return comments


def _comments_touching_range(comments, range_start, range_end):
    """Trả về list comment có Scope giao với [range_start, range_end), kèm
    is_start/is_end đánh dấu đoạn/ô này có phải là điểm bắt đầu/kết thúc của
    scope comment đó hay không (dùng để quyết định chèn START marker)."""
    touching = []
    for c in comments:
        if c["start"] < range_end and c["end"] > range_start:
            touching.append(
                {
                    "id": c["id"],
                    "text": c["text"],
                    "is_start": c["start"] >= range_start,
                    "is_end": c["end"] <= range_end,
                }
            )
    return touching


def _apply_comment_markers(clean_text, touching):
    """Chèn %%COMMENT_START:id%% / %%COMMENT:id%%<text> vào nội dung đã clean
    của một paragraph/cell, theo danh sách comment liên quan tới nó."""
    if not touching:
        return clean_text

    prefix = ""
    suffix = ""
    for c in touching:
        if not c["is_start"]:
            # Comment đã bắt đầu từ paragraph/cell trước đó — không lặp lại
            # marker mở ở đây.
            pass
        elif not c["is_end"]:
            # Bắt đầu tại đây nhưng còn tiếp tục ở paragraph/cell sau.
            prefix += COMMENT_START_TPL.format(id=c["id"])
        # Nếu is_start và is_end đều True: comment gọn trong đúng đơn vị này,
        # không cần marker mở, chỉ cần marker đóng bên dưới.

        if c["is_end"]:
            suffix += COMMENT_END_TPL.format(id=c["id"]) + c["text"]

    return prefix + clean_text + suffix


# ===========================================================
# MODE: --read — render bảng Markdown
# ===========================================================
def render_word_table_markdown(t_idx, table_data, styled_cells):
    """styled_cells: dict {(row, col): text_da_qua_comment_marker} — text đã
    escape | \\r \\n từ trước, không xử lý lại ở đây."""
    rows = table_data["rows"]
    cols = table_data["cols"]

    grid = [["" for _ in range(cols)] for _ in range(rows)]
    for r, c, _ in table_data["cells"]:
        grid[r][c] = styled_cells.get((r, c), "")

    for r, c in table_data.get("unreadable_cells", []):
        grid[r][c] = "[Merged]"

    lines = []
    lines.append(f"## Table {t_idx} ({rows} rows x {cols} cols)")

    headers = ["Row"] + [f"Col {c}" for c in range(cols)]
    lines.append("| " + " | ".join(headers) + " |")
    lines.append("| " + " | ".join(["---"] * len(headers)) + " |")

    for r in range(rows):
        row_vals = [str(r)] + grid[r]
        lines.append("| " + " | ".join(row_vals) + " |")

    lines.append("")
    return "\n".join(lines)


# ===========================================================
# MODE: --read — đọc cấu trúc file qua Word COM
# ===========================================================
def cmd_read(file_path, output_json_path):
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

        comments = _extract_comments(doc)

        result = {
            "file": file_path,
            # paragraphs: mảng chuỗi thuần, index = vị trí trong mảng = paragraph_index dùng cho --apply
            "paragraphs": [],
            # tables: mỗi table gồm kích thước + mảng cell nén [row, col, text].
            # Cell bị merge/unreadable KHÔNG xuất hiện trong "cells", mà liệt kê
            # riêng trong "unreadable_cells" ([row, col]) — không lặp flag false
            # cho từng cell bình thường.
            "tables": [],
        }
        # Bản MD dùng text đã chèn %%COMMENT marker — xây song song với
        # result (plain text, dùng cho JSON) để không phải mở lại Range COM
        # lần 2. KHÔNG ghi comment vào result/JSON.
        md_paragraphs = []
        md_table_cells = (
            []
        )  # list song song với result["tables"], mỗi phần tử là dict {(r,c): styled_text}

        # ---- Paragraphs (1-based COM -> vị trí mảng 0-based = paragraph_index) ----
        para_count = doc.Paragraphs.Count
        for i in range(1, para_count + 1):
            try:
                para_range = doc.Paragraphs(i).Range
                text = _clean_raw_text(para_range.Text).strip()

                styled_text = _escape_comment_marks(text)
                touching = _comments_touching_range(
                    comments, para_range.Start, para_range.End
                )
                styled_text = _apply_comment_markers(styled_text, touching)
            except Exception as e:
                text = f"[ERROR đọc paragraph {i}: {e}]"
                styled_text = text
            result["paragraphs"].append(text)
            md_paragraphs.append(styled_text)

        # ---- Tables (1-based COM -> index 0-based = table_index) ----
        table_count = doc.Tables.Count
        for t in range(1, table_count + 1):
            table = doc.Tables(t)
            try:
                n_rows = table.Rows.Count
                n_cols = table.Columns.Count
            except Exception as e:
                result["tables"].append(
                    {"error": f"Không đọc được kích thước bảng {t-1}: {e}"}
                )
                md_table_cells.append({})
                continue

            cells = []
            unreadable = []
            styled_cells = {}

            for r in range(1, n_rows + 1):
                for c in range(1, n_cols + 1):
                    try:
                        cell = table.Cell(r, c)
                        cell_range = cell.Range
                        cell_text = _clean_raw_text(cell_range.Text).strip()
                        cells.append([r - 1, c - 1, cell_text])

                        styled_text = cell_text.replace("|", "\\|")
                        styled_text = _escape_comment_marks(styled_text)
                        touching = _comments_touching_range(
                            comments, cell_range.Start, cell_range.End
                        )
                        styled_text = _apply_comment_markers(styled_text, touching)
                        styled_cells[(r - 1, c - 1)] = styled_text
                    except Exception:
                        # Cell(r,c) ném lỗi -> tọa độ này rơi vào vùng merge
                        # hoặc không tồn tại độc lập.
                        unreadable.append([r - 1, c - 1])

            table_entry = {"rows": n_rows, "cols": n_cols, "cells": cells}
            if unreadable:
                table_entry["unreadable_cells"] = unreadable
            result["tables"].append(table_entry)
            md_table_cells.append(styled_cells)

        # Xử lý phần mở rộng
        base_path, _ = os.path.splitext(output_json_path)
        json_path = base_path + ".json"
        md_path = base_path + ".md"

        out_dir = os.path.dirname(json_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # 1. Ghi file JSON — KHÔNG chứa comment, chỉ dành cho script mapping.
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, separators=(",", ":"))
        print(f"[SUCCESS] Saved Word structure JSON to: {json_path}")

        # 2. Ghi file MD — có chèn comment marker, dành cho Agent đọc.
        md_lines = []
        md_lines.append(f"# WORD STRUCTURE DUMP: {os.path.basename(file_path)}\n")

        md_lines.append("# PARAGRAPHS\n")
        for styled_text in md_paragraphs:
            md_lines.append(styled_text)
        md_lines.append("\n")

        md_lines.append("# TABLES\n")
        for i, table_data in enumerate(result["tables"]):
            if "error" in table_data:
                md_lines.append(
                    f"## Table {i}\nError reading table: {table_data['error']}\n"
                )
                continue
            md_lines.append(
                render_word_table_markdown(i, table_data, md_table_cells[i])
            )

        if comments:
            md_lines.append(
                f"\n<!-- {len(comments)} comment(s) found — xem quy ước %%COMMENT trong SKILL.md -->"
            )

        md_content = "\n".join(md_lines)
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)
        print(f"[SUCCESS] Saved Word structure MD to: {md_path}")

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

            if "^" in find_text or "^" in replace_text:
                print(
                    f"  [WARN] Rule {i}: Chuỗi chứa '^' — Word sẽ diễn giải như ký tự đặc biệt. Kiểm tra kết quả."
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
                    table = doc.Tables(t_idx)
                    added_rows = 0
                    while r_idx > table.Rows.Count:
                        table.Rows.Add()
                        added_rows += 1
                    if added_rows > 0:
                        print(
                            f"  [INFO] Table[{t_idx-1}]: Đã chèn thêm {added_rows} hàng mới (Tổng số hàng hiện tại: {table.Rows.Count})"
                        )
                    target_range = table.Cell(r_idx, c_idx).Range
                except Exception as e:
                    print(
                        f"  [ERROR] Rule {i}: Không lấy được Table[{t_idx-1}] R{r_idx-1}C{c_idx-1}: {e}"
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
                        if anchor in para_text and (
                            not find_text or find_text in para_text
                        ):
                            found_para = para
                            break
                    if found_para is None:
                        print(
                            f"  [MISS] Rule {i}: Không tìm thấy anchor='{anchor[:50]}'"
                        )
                        continue
                    target_range = found_para.Range

                elif p_idx is not None:
                    try:
                        target_range = doc.Paragraphs(p_idx + 1).Range
                    except Exception as e:
                        print(
                            f"  [ERROR] Rule {i}: paragraph_index={p_idx} không hợp lệ: {e}"
                        )
                        continue
                    # find_text vẫn phải khớp trong đoạn này nếu find_text không rỗng
                    if find_text and find_text not in target_range.Text:
                        print(
                            f"  [MISS] Rule {i}: paragraph_index={p_idx} không chứa find_text='{find_text[:50]}' "
                            f"— khả năng file đã bị thay đổi kể từ lúc --read."
                        )
                        continue

                else:
                    print(
                        f"  [ERROR] Rule {i}: Scope 'paragraph' cần 'anchor' hoặc 'paragraph_index'."
                    )
                    continue

            else:
                print(
                    f"  [ERROR] Rule {i}: Scope không hợp lệ '{scope}'. Dùng 'paragraph' hoặc 'table_cell'."
                )
                continue

            if not find_text:
                try:
                    target_range.Text = replace_text
                    print(f"  [OK]   Rule {i}: [direct write] → '{replace_text}'")
                    total_replaced += 1
                except Exception as e:
                    print(f"  [ERROR] Rule {i}: Ghi trực tiếp thất bại: {e}")
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
                    f"  [OK]   Rule {i}: [max={limit_str}] '{find_text}' → '{replace_text}' ({n}x thực tế)"
                )
                total_replaced += n
            else:
                print(
                    f"  [MISS] Rule {i}: '{find_text}' không tìm thấy trong target range"
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
            f"    [INFO] replace_text dài ({len(replace_text)} ký tự) — dùng Range trực tiếp, không bị giới hạn 255 ký tự."
        )

    search_range = word_range.Duplicate
    find_obj = search_range.Find
    find_obj.ClearFormatting()
    find_obj.Text = find_text
    find_obj.Forward = True
    find_obj.Wrap = 0  # wdFindStop
    find_obj.MatchCase = True
    find_obj.MatchWholeWord = False
    find_obj.MatchWildcards = False

    count = 0
    unlimited = max_replacements == 0

    while find_obj.Execute():
        search_range.Text = replace_text
        count += 1
        search_range.Collapse(0)  # wdCollapseEnd = 0
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
        if len(args) < 3:
            print(
                "[ERROR] --read cần đường dẫn file .docx và đường dẫn file output JSON"
            )
            print("Usage: python word_editor.py --read <file.docx> <output_json_path>")
            sys.exit(1)
        file_path = os.path.abspath(args[1])
        output_json_path = os.path.abspath(args[2])
        if not os.path.exists(file_path):
            print(f"[ERROR] File không tồn tại: {file_path}")
            sys.exit(1)
        cmd_read(file_path, output_json_path)

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
