import sys
import os
import openpyxl
import re

# Set UTF-8 encoding for console output on Windows
sys.stdout.reconfigure(encoding='utf-8')

def clean_cell_value(val):
    if val is None:
        return ""
    # Strip double quotes or backticks to avoid breaking markdown format
    val_str = str(val).replace("|", "\\|").replace("\n", " ")
    return val_str.strip()

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
    
    for sheet in wb.worksheets:
        # Check sheet visibility
        is_hidden = sheet.sheet_state != 'visible'
        if is_hidden and not include_hidden:
            continue
            
        status = " (Hidden)" if is_hidden else ""
        print(f"## Sheet: {sheet.title}{status}")
        
        # Check if sheet is empty
        if sheet.max_row == 0 or sheet.max_column == 0:
            print("*Sheet is empty.*\n")
            continue
            
        # Parse sheet grid
        grid = []
        for r in range(1, sheet.max_row + 1):
            row_data = []
            for c in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=r, column=c)
                row_data.append(clean_cell_value(cell.value))
            grid.append(row_data)
            
        # Strip trailing empty rows and columns to clean up output
        while grid and all(val == "" for val in grid[-1]):
            grid.pop()
            
        if not grid:
            print("*Sheet has only empty cells.*\n")
            continue
            
        # Find maximum columns used in non-empty rows
        max_cols = max(len(row) for row in grid)
        
        # Strip trailing empty columns
        col_is_empty = [True] * max_cols
        for row in grid:
            for c in range(len(row)):
                if row[c] != "":
                    col_is_empty[c] = False
                    
        # Filter grid based on active columns
        clean_grid = []
        for row in grid:
            clean_row = [row[c] for c in range(len(row)) if not col_is_empty[c]]
            clean_grid.append(clean_row)
            
        if not clean_grid or not clean_grid[0]:
            print("*Sheet has only empty cells after cleanup.*\n")
            continue
            
        # Render as Markdown Table
        headers = clean_grid[0]
        # Fill empty headers
        headers = [h if h else f"Col {i+1}" for i, h in enumerate(headers)]
        
        print("| " + " | ".join(headers) + " |")
        print("| " + " | ".join(["---"] * len(headers)) + " |")
        
        for row in clean_grid[1:]:
            # Ensure row length matches header length
            row_vals = row + [""] * (len(headers) - len(row))
            row_vals = row_vals[:len(headers)]
            print("| " + " | ".join(row_vals) + " |")
            
        print("\n")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python excel_parser.py <path_to_excel_file> [--all]")
        sys.exit(1)
        
    include_hidden = "--all" in sys.argv
    excel_to_markdown(sys.argv[1], include_hidden)
