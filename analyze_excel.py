# -*- coding: utf-8 -*-
import xlrd
import sys

file_path = r'd:\Projects\phan mem quet mat\000001_02Báo cáo 28.02.2026.xls'
out_file = r'd:\Projects\phan mem quet mat\excel_analysis.log'
wb = xlrd.open_workbook(file_path)
sheet = wb.sheet_by_index(0)

with open(out_file, 'w', encoding='utf-8') as f:
    f.write(f"Sheet: '{sheet.name}', Rows: {sheet.nrows}, Cols: {sheet.ncols}\n")
    
    # Print rows 0-15 in detail
    f.write("\n--- Rows 0-15 ---\n")
    for r in range(min(16, sheet.nrows)):
        row_data = []
        for c in range(sheet.ncols):
            val = sheet.cell_value(r, c)
            ctype = sheet.cell_type(r, c)
            if ctype != xlrd.XL_CELL_EMPTY and val != '':
                row_data.append(f"C{c}={repr(val)[:60]}")
        f.write(f"Row {r}: {row_data}\n")
    
    # Find all person rows
    f.write("\n--- All person rows ---\n")
    for r in range(sheet.nrows):
        c0 = sheet.cell_value(r, 0)
        c1 = sheet.cell_value(r, 1) if sheet.ncols > 1 else ''
        if isinstance(c0, (int, float)) and c0 > 1000 and isinstance(c1, str) and len(c1) > 3:
            row_full = [sheet.cell_value(r, c) for c in range(sheet.ncols)]
            f.write(f"Row {r}: ID={c0}, Name={c1}, Full={row_full[:8]}\n")
    
    # Detailed first person block (header + ~5 data rows)
    f.write("\n--- First person detailed block ---\n")
    first_found = False
    for r in range(sheet.nrows):
        c0 = sheet.cell_value(r, 0)
        c1 = sheet.cell_value(r, 1) if sheet.ncols > 1 else ''
        if isinstance(c0, (int, float)) and c0 > 1000 and isinstance(c1, str) and len(c1) > 3:
            if not first_found:
                first_found = True
                start = max(0, r-3)
                end = min(sheet.nrows, r+35)
                for rr in range(start, end):
                    row_data = [sheet.cell_value(rr, c) for c in range(sheet.ncols)]
                    f.write(f"Row {rr}: {row_data}\n")
                break

print(f"Analysis written to {out_file}")
