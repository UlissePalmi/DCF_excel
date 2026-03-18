"""
Format definitions for the Summary sheet.

This module centralizes all format objects used in the Summary sheet builder,
keeping format definitions separate from sheet layout and data logic.
"""


def index_to_column(idx: int) -> str:
    """Convert 0-based column index to Excel column letter(s).

    Examples: 0='A', 25='Z', 26='AA', 27='AB', etc.
    """
    result = ""
    idx += 1  # Convert to 1-based
    while idx > 0:
        idx -= 1
        result = chr(ord('A') + (idx % 26)) + result
        idx //= 26
    return result


def create_summary_formats(workbook):
    """
    Create all format objects for the Summary sheet.

    Args:
        workbook: xlsxwriter.Workbook instance

    Returns:
        dict: Dictionary of format objects keyed by format name
    """
    formats = {}


    formats['fmt_section'] = workbook.add_format({
        'font_size': 10, 'bold': True,
        'align': 'center_across'
    })

    formats['fmt_section_left'] = workbook.add_format({
        'font_size': 10, 'bold': True
    })

    # Header formats
    formats['fmt_projected'] = workbook.add_format({
        'font_size': 10, 'bold': True, 'italic': True,
        'align': 'center_across', 'bottom': 1, 'top': 1
    })

    formats['fmt_hdr_white_border'] = workbook.add_format({
        'font_size': 10, 'bold': True,
        'align': 'vcenter', 'bottom': 1
    })

    # Label formats
    formats['fmt_label'] = workbook.add_format({
        'font_size': 10
    })

    formats['fmt_label_border'] = workbook.add_format({
        'font_size': 10, 'bottom': 1
    })

    formats['fmt_sub_label'] = workbook.add_format({
        'font_size': 10
    })

    # Data value formats
    formats['fmt_impl_white'] = workbook.add_format({
        'font_size': 10, 'bold': True
    })

    # Border formats for the outside box
    formats['fmt_border_thin'] = workbook.add_format({'bottom': 1})
    formats['fmt_border_top'] = workbook.add_format({'top': 1})
    formats['fmt_border_bottom'] = workbook.add_format({'bottom': 1})
    formats['fmt_border_left'] = workbook.add_format({'left': 1})
    formats['fmt_border_right'] = workbook.add_format({'right': 1})
    formats['fmt_border_top_left'] = workbook.add_format({'top': 1, 'left': 1})
    formats['fmt_border_top_right'] = workbook.add_format({'top': 1, 'right': 1})
    formats['fmt_border_bottom_left'] = workbook.add_format({'bottom': 1, 'left': 1})
    formats['fmt_border_bottom_right'] = workbook.add_format({'bottom': 1, 'right': 1})

    # Separator formats
    formats['fmt_dashed_border'] = workbook.add_format({'bottom': 4})

    return formats

def write_summary_content(builder, start_row: int, start_col: int, end_proj: int) -> None:
    """Write income statement and valuation summary content.

    Args:
        builder: SummarySheetBuilder instance with ws and format attributes
        start_row: Starting row index (0-based)
        start_col: Starting column index (0-based, where B=1)
        end_proj: Column index where projected data ends
    """
    ws = builder.ws

    # Income Statement Items section
    ws.write(start_row + 5, start_col + 2, 'Income Statement Items', builder.fmt_section_left)

    ws.write(start_row + 7, start_col + 3, 'Net Revenue', builder.fmt_label)
    ws.write(start_row + 8, start_col + 3, '   Growth', builder.fmt_sub_label)
    ws.write(start_row + 11, start_col + 3, 'EBITDA', builder.fmt_label)
    ws.write(start_row + 12, start_col + 3, '   Margin', builder.fmt_sub_label)
    ws.write(start_row + 13, start_col + 3, '   Growth', builder.fmt_sub_label)
    ws.write(start_row + 16, start_col + 3, 'Net Income', builder.fmt_label)
    ws.write(start_row + 17, start_col + 3, '   Margin', builder.fmt_sub_label)
    ws.write(start_row + 18, start_col + 3, '   Growth', builder.fmt_sub_label)
    ws.write(start_row + 21, start_col + 3, 'NOPAT', builder.fmt_sub_label)
    ws.write(start_row + 22, start_col + 3, '   Margin', builder.fmt_sub_label)
    ws.write(start_row + 23, start_col + 3, '   Growth', builder.fmt_sub_label)
    ws.write(start_row + 26, start_col + 3, 'D&A', builder.fmt_sub_label)
    ws.write(start_row + 27, start_col + 3, 'Capex', builder.fmt_sub_label)
    ws.write(start_row + 28, start_col + 3, 'NWC', builder.fmt_sub_label)
    ws.write(start_row + 31, start_col + 3, 'Unlevered FCFF', builder.fmt_label)
    ws.write(start_row + 34, start_col + 3, '   Discount Period', builder.fmt_label)
    ws.write(start_row + 35, start_col + 3, '   Discount Factor', builder.fmt_label)
    ws.write(start_row + 36, start_col + 3, 'Present Value of FCF', builder.fmt_impl_white)

    # Right column (V) valuation summary
    ws.write(start_row + 5, end_proj + 1, 'Discount Rate', builder.fmt_label)
    ws.write(start_row + 7, end_proj + 1, 'Terminal Growth Rate', builder.fmt_label)
    ws.write(start_row + 8, end_proj + 1, 'Terminal Value', builder.fmt_label)
    ws.write(start_row + 11, end_proj + 1, 'Cumulative PV of FCF', builder.fmt_label)
    ws.write(start_row + 16, end_proj + 1, 'PV of Terminal Value', builder.fmt_label)
    ws.write(start_row + 21, end_proj + 1, 'Enterprise Value', builder.fmt_label)
    ws.write(start_row + 22, end_proj + 1, 'Net Cash', builder.fmt_label)
    ws.write(start_row + 23, end_proj + 1, 'Equity Value', builder.fmt_label)
    ws.write(start_row + 27, end_proj + 1, 'Shares (MM)', builder.fmt_label)
    ws.write(start_row + 31, end_proj + 1, 'Implied Shared Price', builder.fmt_impl_white)

def draw_border_box(builder, start_row: int, start_col: int, end_proj: int, draw_closing_border: bool = True) -> None:
    """Draw the border box around the content section.

    Args:
        builder: SummarySheetBuilder instance with ws and format attributes
        start_row: Starting row index (0-based)
        start_col: Starting column index (0-based, where B=1)
        end_proj: Column index where projected data ends
        draw_closing_border: Whether to draw the closing border row (default True)
    """
    ws = builder.ws

    # Corner borders
    ws.write(start_row + 1, start_col + 1, '', builder.fmt_border_top_left)
    ws.write(start_row + 1, end_proj + 4, '', builder.fmt_border_top_right)
    ws.write(start_row + 37, start_col + 1, '', builder.fmt_border_bottom_left)
    ws.write(start_row + 37, end_proj + 4, '', builder.fmt_border_bottom_right)

    # Top border - fixed left section
    for col in range(start_col + 2, start_col + 9):
        ws.write(start_row + 1, col, '', builder.fmt_border_top)

    # Projected years label
    ws.write(start_row + 1, start_col + 9, 'Projected', builder.fmt_projected)
    for col in range(start_col + 10, end_proj):
        ws.write(start_row + 1, col, '', builder.fmt_projected)

    # Top border - fixed right section
    for col in range(end_proj, end_proj + 4):
        ws.write(start_row + 1, col, '', builder.fmt_border_top)

    # Left and Right side borders
    for offset in range(2, 37):
        ws.write(start_row + offset, start_col + 1, '', builder.fmt_border_left)
        ws.write(start_row + offset, end_proj + 4, '', builder.fmt_border_right)

    # Bottom border
    for col in range(start_col + 2, end_proj + 4):
        ws.write(start_row + 37, col, '', builder.fmt_border_bottom)

    # Closing border row (only for left-side content)
    if draw_closing_border:
        for col in range(start_col, end_proj + 6):
            ws.write(start_row + 38, col, '', builder.fmt_border_thin)


def write_year_headers(builder, start_row: int, start_col: int, end_proj: int) -> None:
    """Write year header row with historical and projected year formulas.

    Args:
        builder: SummarySheetBuilder instance with ws and format attributes
        start_row: Starting row index (0-based)
        start_col: Starting column index (0-based, where B=1)
        end_proj: Column index where projected data ends
    """
    ws = builder.ws
    year_row = start_row + 3  # 1-based Excel row number for year headers

    # Fixed left section headers
    ws.write(start_row + 2, start_col + 2, '($ Millions)', builder.fmt_hdr_white_border)
    ws.write(start_row + 2, start_col + 3, '', builder.fmt_label_border)
    ws.write(start_row + 2, start_col + 4, 'Trend', builder.fmt_hdr_white_border)
    ws.write(start_row + 2, start_col + 5, '', builder.fmt_label_border)

    # Historical years (first 3 data columns): each references next column - 1
    for col_idx in range(start_col + 6, start_col + 9):
        # Column at col_idx references next column (col_idx + 1) with -1
        next_col_letter = index_to_column(col_idx + 1)
        formula = f'={next_col_letter}{year_row}-1'
        ws.write(start_row + 2, col_idx, formula, builder.fmt_label_border)

    # First projected year (4th data column): references assumptions
    ws.write(start_row + 2, start_col + 9, '=Assumptions!A1', builder.fmt_label_border)

    # Remaining projected years: each references previous column + 1
    for col_idx in range(start_col + 10, end_proj):
        prev_col_letter = index_to_column(col_idx - 1)
        formula = f'={prev_col_letter}{year_row}+1'
        ws.write(start_row + 2, col_idx, formula, builder.fmt_label_border)


def adjust_column_widths(builder, start_col: int, end_proj: int) -> None:
    """Adjust column widths for the content section.

    Args:
        builder: SummarySheetBuilder instance with ws attribute
        start_col: Starting column index (0-based, where B=1)
        end_proj: Column index where projected data ends
    """
    ws = builder.ws

    # Adjust column widths based on start_col
    col_widths = {}
    col_widths[index_to_column(start_col)] = 4      # B equivalent (or offset from start_col)
    col_widths[index_to_column(start_col + 1)] = 1  # C equivalent
    col_widths[index_to_column(start_col + 2)] = 1  # D equivalent
    col_widths[index_to_column(start_col + 3)] = 20  # E equivalent
    col_widths[index_to_column(start_col + 5)] = 1  # G equivalent

    # Right side columns based on end_proj
    right_cols = [index_to_column(end_proj + i) for i in range(5)]
    col_widths[right_cols[0]] = 1   # U equivalent
    col_widths[right_cols[1]] = 10  # V equivalent
    col_widths[right_cols[4]] = 1   # Y equivalent (5th column)

    for col, width in col_widths.items():
        ws.set_column(col + ':' + col, width)


def adjust_row_heights(builder, start_row: int) -> None:
    """Adjust row heights for the content section.

    Args:
        builder: SummarySheetBuilder instance with ws attribute
        start_row: Starting row index (0-based)
    """
    ws = builder.ws

    # Set row heights for content section
    row_heights_offsets = {
        3: 3, 4: 3, 6: 3, 9: 3, 10: 3, 14: 3, 15: 3, 19: 3,
        20: 3, 24: 3, 25: 3, 29: 3, 30: 3, 32: 3, 33: 3, 37: 3,
    }
    for offset, height in row_heights_offsets.items():
        ws.set_row(start_row + offset, height)