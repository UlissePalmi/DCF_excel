"""
Reusable sheet structure builders for Excel output.

This module contains functions for building common sheet structures
(headers, sections, etc.) that can be used across multiple sheet types.
"""


def write_header(ws, workbook, company_name: str, subtitle: str, start_row: int = 0,
                 col_start: int = 1, col_end: int = 25) -> int:
    """
    Write header section (3 rows: title, subtitle, border).

    Args:
        ws: xlsxwriter worksheet
        workbook: xlsxwriter workbook (used to create formats)
        company_name: Name of company to display in title
        subtitle: Subtitle text to display in second row
        start_row: Starting row index (0-based)
        col_start: First column index for the header span (default 1 = col B)
        col_end: Last column index for the header span (default 25 = col Z)

    Returns:
        Next available row index
    """
    fmt_title = workbook.add_format({'font_size': 18, 'bold': True, 'align': 'center_across'})
    fmt_subtitle = workbook.add_format({'font_size': 14, 'bold': True, 'align': 'center_across'})
    fmt_border = workbook.add_format({'bottom': 2})

    # Set row heights for header
    ws.set_row(start_row, 23.25)        # Title row
    ws.set_row(start_row + 1, 18)       # Subtitle row
    ws.set_row(start_row + 2, 3)        # Border row

    # Title row - center across col_start to col_end
    ws.write(start_row, col_start, company_name, fmt_title)
    for col in range(col_start + 1, col_end + 1):
        ws.write(start_row, col, '', fmt_title)

    # Subtitle row
    ws.write(start_row + 1, col_start, subtitle, fmt_subtitle)
    for col in range(col_start + 1, col_end + 1):
        ws.write(start_row + 1, col, '', fmt_subtitle)

    # Border row
    for col in range(col_start, col_end + 1):
        ws.write(start_row + 2, col, '', fmt_border)

    return start_row + 3
