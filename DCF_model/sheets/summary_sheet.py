"""
Summary sheet builder - creates key metrics summary page.
"""

import xlsxwriter
from .summary_formats import create_summary_formats, draw_border_box, write_summary_content
from structures.structures import write_header

# Fixed columns before the projected years in the header span (B through the last non-projected col)
# col_end = HEADER_FIXED_COLS + num_projected_years  →  with 10 proj years: 15 + 10 = 25 (col Z)
HEADER_COL_START = 1   # col B
HEADER_FIXED_COLS = 15

class SummarySheetBuilder:
    """Builds a formatted Summary sheet with model highlights."""

    def __init__(self, workbook: xlsxwriter.Workbook, company_name: str, num_projected_years: int = 0):
        self.workbook = workbook
        self.company_name = company_name
        self.num_projected_years = num_projected_years
        self.ws = None
        self._calculate_columns()

    def _calculate_columns(self) -> None:
        """Generate dynamic column ranges based on num_projected_years."""
        HIST_YEARS = 3
        total_data_cols = HIST_YEARS + self.num_projected_years

        # Data columns start at H (index 7)
        self.data_start_idx = 7

        # Generate LEFT_COLS dynamically (e.g., H, I, J, ... T)
        self.LEFT_COLS = [chr(ord('A') + i) for i in range(self.data_start_idx, self.data_start_idx + total_data_cols)]

        # Projected columns skip the first 3 historical years
        self.PROJ_L = self.LEFT_COLS[HIST_YEARS:]

        # Right side is offset by 26 columns (one alphabet)
        right_offset = 26
        self.RIGHT_COLS = [chr(ord('A') + self.data_start_idx + right_offset + i) for i in range(total_data_cols)]
        self.PROJ_R = self.RIGHT_COLS[HIST_YEARS:]

    @property
    def data_end_idx(self) -> int:
        """Index of the last data column."""
        return self.data_start_idx + len(self.LEFT_COLS) - 1

    @property
    def right_start_idx(self) -> int:
        """Index of the first fixed right column (after data)."""
        return self.data_end_idx + 1

    @property
    def right_end_idx(self) -> int:
        """Index of the last fixed right column."""
        return self.right_start_idx + 4  # 5 columns: U-Y


    def build(self) -> None:
        """Create and populate the Summary sheet with 3 header/content pairs."""
        self.ws = self.workbook.add_worksheet("Summary")
        self.ws.hide_gridlines(2)
        self._setup_layout()
        self._load_formats()

        row = 0

        # First header/content pair
        row = self._write_header(start_row=row, subtitle='Base Case DCF')
        row = self._write_content(start_row=row, section_title='BASE CASE')

        # Second header/content pair
        row = self._write_header(start_row=row, subtitle='Best Case DCF')
        row = self._write_content(start_row=row, section_title='BEST CASE')

        # Third header/content pair
        row = self._write_header(start_row=row, subtitle='Worse Case DCF')
        row = self._write_content(start_row=row, section_title='WORSE CASE')

    def _setup_layout(self) -> None:
        """Apply custom column widths based on dynamic column layout."""
        col_widths = {
            'A': 2,      # Outer left spacing
            'B': 4,      # Outer left spacing
            'C': 1,      # Left border
            'D': 1,
            'E': 20,
            'G': 1,
        }

        # Add right side fixed columns (5 columns after data)
        data_end_col = self.LEFT_COLS[-1]
        data_end_idx = ord(data_end_col) - ord('A')

        right_cols = [chr(ord('A') + data_end_idx + 1 + i) for i in range(5)]
        col_widths[right_cols[0]] = 1  # U equivalent
        col_widths[right_cols[1]] = 10  # V equivalent
        col_widths[right_cols[4]] = 1  # Y equivalent (5th column)

        for col, width in col_widths.items():
            self.ws.set_column(col + ':' + col, width)


    def _load_formats(self) -> None:
        """Load format objects from summary_formats module."""
        formats = create_summary_formats(self.workbook)
        for name, fmt in formats.items():
            setattr(self, name, fmt)

    def _write_header(self, start_row: int = 0, subtitle: str = 'Base Case DCF') -> int:
        """
        Write header section (3 rows: title, subtitle, border).

        Args:
            start_row: Starting row index (0-based)
            subtitle: Subtitle text to display (default: 'Base Case DCF')

        Returns:
            Next available row index
        """
        return write_header(
            self.ws,
            self.workbook,
            self.company_name,
            subtitle,
            start_row=start_row,
            col_start=HEADER_COL_START,
            col_end= self.num_projected_years + 15,
        ) + 1  # +1 for blank gap after header

    def _write_content(self, start_row, section_title, start_col: int = 1) -> int:
        """
        Write content section (rows starting at start_row).

        Args:
            start_row: Starting row index (0-based). Should be 4 (row 5) or later.
            section_title: Title text for the section header
            start_col: Starting column index (default 1 = column B)

        Returns:
            Next available row index
        """
        end_proj = start_col + self.num_projected_years + 9

        # Set row heights for content section
        row_heights_offsets = {
            3: 3, 4: 3, 6: 3, 9: 3, 10: 3, 14: 3, 15: 3, 19: 3,
            20: 3, 24: 3, 25: 3, 29: 3, 30: 3, 32: 3, 33: 3, 37: 3,
        }
        for offset, height in row_heights_offsets.items():
            self.ws.set_row(start_row + offset, height)

        # Section header
        self.ws.write(start_row, start_col + 1, f"SUMMARY VALUES - {section_title}", self.fmt_section)
        for col in range(start_col + 2, end_proj + 5):
            self.ws.write(start_row, col, '', self.fmt_section)

        
        # Column headers (year headers row, using dynamic row calculation)
        year_row = start_row + 3  # 1-based Excel row number for year headers

        # Fixed left section headers
        self.ws.write(start_row + 2, start_col + 2, '($ Millions)', self.fmt_hdr_white_border)
        self.ws.write(start_row + 2, start_col + 3, '', self.fmt_label_border)
        self.ws.write(start_row + 2, start_col + 4, 'Trend', self.fmt_hdr_white_border)
        self.ws.write(start_row + 2, start_col + 5, '', self.fmt_label_border)

        # Historical years (first 3 data columns): each references next column - 1
        for i, col_idx in enumerate(range(start_col + 6, start_col + 9)):
            # Column at i references column at i+1 with -1
            next_col_letter = self.LEFT_COLS[i + 1]
            formula = f'={next_col_letter}{year_row}-1'
            self.ws.write(start_row + 2, col_idx, formula, self.fmt_label_border)

        # First projected year (4th data column): references assumptions
        self.ws.write(start_row + 2, start_col + 9, '=Assumptions!A1', self.fmt_label_border)

        # Remaining projected years: each references previous column + 1
        for col_idx in range(start_col + 10, end_proj):
            prev_col_letter = chr(ord('A') + col_idx - 1)
            formula = f'={prev_col_letter}{year_row}+1'
            self.ws.write(start_row + 2, col_idx, formula, self.fmt_label_border)

        # Dotted border separators across data columns
        # Offsets from start_row: 9, 14, 19, 24, 29, 32
        for offset in [9, 14, 19, 24, 29, 32]:
            for col in range(start_col + 6, end_proj):
                self.ws.write_blank(start_row + offset, col, None, self.fmt_dashed_border)

        # Write summary content (income statement and valuation summary)
        write_summary_content(self, start_row, start_col, end_proj)

        # Draw border box around content
        draw_border_box(self, start_row, start_col, end_proj)

        return start_row + 40  # +40 for content rows (39) + blank gap (1)


