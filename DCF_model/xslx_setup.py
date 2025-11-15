import xlsxwriter
import os
from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol
#from openpyxl import Workbook 
#from openpyxl.styles import Font, PatternFill, Alignment, Border, Side 
#from openpyxl.worksheet.views import SheetView
import pandas as pd
import math

def row_col_fmt(wb, ws):    
    ws.hide_gridlines(2)                                                                            # Hide gridlines (0=show, 1=hide on screen, 2=hide everywhere)
    fmt_default = wb.add_format({"font_name": "Arial", "font_size": 11})
    ws.set_column("A:XFD", None, fmt_default)
    ws.set_default_row(12.75)
    ws.set_row(1, 23.25)
    ws.set_row(2, 18.75)
    ws.set_row(3, 3)
    ws.set_column("A:A", 3)
    ws.set_column("B:C", 1)
    ws.set_column("D:D", 11)
    ws.set_column("E:E", 13.5)
    ws.set_column("F:F", 9.75)
    ws.set_column("G:G", 1)
    ws.set_column("H:W", 8.5)
    
def center_across_range(ws, first_cell, last_cell, text, fmt):
    """Center `text` across first_cell:last_cell (one row) using center_across."""
    r1, c1 = xl_cell_to_rowcol(first_cell)
    r2, c2 = xl_cell_to_rowcol(last_cell)
    if r1 != r2:
        raise ValueError("Range must be a single row for Center Across Selection.")
    ws.write(r1, c1, text, fmt)
    for c in range(c1 + 1, c2 + 1):
        ws.write_blank(r1, c, None, fmt)

'''
def write_dataframe(ws, df, start_row, start_col, *, write_index=False, header_format=None, cell_format=None):
    """
    Write a pandas DataFrame to an XlsxWriter worksheet without using ExcelWriter.
    - Converts NaN -> None (blank in Excel)
    - Writes headers (optionally)
    """
    # Header
    col = start_col
    if write_index:
        ws.write(start_row, col, df.index.name or "", header_format)
        col += 1
    for name in df.columns:
        ws.write(start_row, col, name, header_format)
        col += 1

    # Rows
    r = start_row + 1
    # Pre-convert to Python types, replace NaN with None
    values = df.where(df.notna(), None).to_numpy().tolist()
    for idx, row_vals in zip(df.index, values):
        c = start_col
        if write_index:
            ws.write(r, c, idx, cell_format)
            c += 1
        ws.write_row(r, c, row_vals, cell_format)
        r += 1
'''

def write_dataframe(ws, df, start_row, start_col, *, write_index=False, header_format=None, cell_format=None):
    # Pre-convert to Python types, replace NaN with None
    values = df.where(df.notna(), 0).to_numpy().tolist()
    for idx, row_vals in zip(df.index, values):
        c = start_col
        if write_index:
            ws.write(start_row, c, idx, cell_format)
            c += 1
        ws.write_row(start_row, c, row_vals, cell_format)
        start_row += 1