import sys
import os
from datetime import datetime, date, time
import openpyxl

# Set UTF-8 encoding for console output on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def clean_cell_value(val):
    if val is None:
        return ""

    # Format dates/times so the model sees a plain readable string,
    # not "2026-07-17 00:00:00" from str(datetime).
    if isinstance(val, datetime):
        if val.time() == time(0, 0):
            val_str = val.strftime("%Y-%m-%d")
        else:
            val_str = val.strftime("%Y-%m-%d %H:%M:%S")
    elif isinstance(val, date):
        val_str = val.strftime("%Y-%m-%d")
    elif isinstance(val, time):
        val_str = val.strftime("%H:%M:%S")
    elif isinstance(val, float) and val.is_integer():
        # Avoid "1.0" for what is really the integer 1
        val_str = str(int(val))
    else:
        val_str = str(val)

    # Escape/strip characters that would break a markdown table row
    val_str = (
        val_str.replace("|", "\\|")
        .replace("\r\n", " ")
        .replace("\n", " ")
        .replace("\r", " ")
    )
    return val_str.strip()


def build_merge_value_map(sheet):
    """Merged cells only carry a value in their top-left cell; openpyxl
    returns None for the rest (and disallows writing to them directly).
    Build a lookup so every cell in a merge range resolves to that value."""
    merge_map = {}
    for merged_range in sheet.merged_cells.ranges:
        top_left_val = sheet.cell(merged_range.min_row, merged_range.min_col).value
        for r in range(merged_range.min_row, merged_range.max_row + 1):
            for c in range(merged_range.min_col, merged_range.max_col + 1):
                merge_map[(r, c)] = top_left_val
    return merge_map


def sheet_to_grid(sheet):
    merge_map = build_merge_value_map(sheet)

    grid = []
    for r in range(1, sheet.max_row + 1):
        row_data = []
        for c in range(1, sheet.max_column + 1):
            val = merge_map.get((r, c), sheet.cell(row=r, column=c).value)
            row_data.append(clean_cell_value(val))
        grid.append(row_data)

    # Strip trailing empty rows
    while grid and all(v == "" for v in grid[-1]):
        grid.pop()
    if not grid:
        return []

    # Pad every row to the same width BEFORE filtering columns, otherwise
    # column filtering misaligns rows of different lengths.
    max_cols = max(len(row) for row in grid)
    grid = [row + [""] * (max_cols - len(row)) for row in grid]

    # Find columns that are empty across all rows, and drop them
    col_is_empty = [all(row[c] == "" for row in grid) for c in range(max_cols)]
    clean_grid = [
        [row[c] for c in range(max_cols) if not col_is_empty[c]] for row in grid
    ]

    return clean_grid


def render_sheet_markdown(sheet):
    lines = []
    status = " (Hidden)" if sheet.sheet_state != "visible" else ""
    lines.append(f"## Sheet: {sheet.title}{status}")

    if sheet.max_row == 0 or sheet.max_column == 0:
        lines.append("*Sheet is empty.*\n")
        return "\n".join(lines)

    grid = sheet_to_grid(sheet)
    if not grid or not grid[0]:
        lines.append("*Sheet has only empty cells.*\n")
        return "\n".join(lines)

    headers = [h if h else f"Col {i+1}" for i, h in enumerate(grid[0])]
    lines.append("| " + " | ".join(headers) + " |")
    lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
    for row in grid[1:]:
        lines.append("| " + " | ".join(row) + " |")
    lines.append("")

    return "\n".join(lines)


def excel_to_markdown(filepath, include_hidden=False):
    if not os.path.exists(filepath):
        print(f"Error: File not found: {filepath}")
        sys.exit(1)

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        sys.exit(1)

    print(f"# EXCEL DATA DUMP: {os.path.basename(filepath)}\n")

    any_sheet_shown = False
    for sheet in wb.worksheets:
        is_hidden = sheet.sheet_state != "visible"
        if is_hidden and not include_hidden:
            continue
        any_sheet_shown = True
        try:
            print(render_sheet_markdown(sheet))
        except Exception as e:
            print(f"## Sheet: {sheet.title}\n*Error reading this sheet: {e}*\n")

    if not any_sheet_shown:
        print("*No visible sheets found. Re-run with --all to include hidden sheets.*")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python excel_parser.py <path_to_excel_file> [--all]")
        sys.exit(1)

    include_hidden = "--all" in sys.argv
    excel_to_markdown(sys.argv[1], include_hidden)
