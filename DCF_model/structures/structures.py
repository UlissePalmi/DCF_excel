"""
Reusable sheet structure builders for Excel output.

This module contains functions for building common sheet structures
(headers, sections, etc.) that can be used across multiple sheet types.
"""


def write_header(ws, company_name: str, subtitle: str, start_row: int = 0, fmt_title=None, fmt_subtitle=None, fmt_border=None) -> int:
    """
    Write header section (3 rows: title, subtitle, border).

    Args:
        ws: xlsxwriter worksheet
        company_name: Name of company to display in title
        subtitle: Subtitle text to display in second row
        start_row: Starting row index (0-based)
        fmt_title: Format for title row
        fmt_subtitle: Format for subtitle row
        fmt_border: Format for border row

    Returns:
        Next available row index
    """
    # Set row heights for header
    ws.set_row(start_row, 23.25)        # Title row
    ws.set_row(start_row + 1, 18)       # Subtitle row
    ws.set_row(start_row + 2, 3)        # Border row

    # Title section - center across B to Z
    ws.write(start_row, 1, company_name, fmt_title)
    for col in range(2, 26):  # C to Z (columns 3-26)
        ws.write(start_row, col, '', fmt_title)

    ws.write(start_row + 1, 1, subtitle, fmt_subtitle)
    for col in range(2, 26):  # C to Z
        ws.write(start_row + 1, col, '', fmt_subtitle)

    # Border row
    for col in range(1, 26):  # B to Z
        ws.write(start_row + 2, col, '', fmt_border)

    return start_row + 3
