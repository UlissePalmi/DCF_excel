"""
Summary sheet builder - creates key metrics summary page.
"""

import xlsxwriter
from .summary_formats import create_summary_formats
from structures.structures import write_header

# Year columns (left mirror and right source)
LEFT_COLS = ['H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']
RIGHT_COLS = ['AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT']
PROJ_L = ['K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']
PROJ_R = ['AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT']


class SummarySheetBuilder:
    """Builds a formatted Summary sheet with model highlights."""

    def __init__(self, workbook: xlsxwriter.Workbook, company_name: str, all_years: list, num_projected_years: int = 0):
        self.workbook = workbook
        self.company_name = company_name
        self.all_years = all_years
        self.num_projected_years = num_projected_years
        self.ws = None

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
        """Apply custom column widths and row heights."""
        col_widths = {
            'A': 2, 'B': 4, 'C': 1, 'D': 1, 'E': 20, 'G': 1, 'U': 1, 'V': 10, 'Y': 1,
            'AC': 1, 'AD': 1, 'AE': 20, 'AG': 1, 'AU': 1, 'AV': 10, 'AY': 1,
        }

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
            self.company_name,
            subtitle,
            start_row=start_row,
            fmt_title=self.fmt_title,
            fmt_subtitle=self.fmt_subtitle,
            fmt_border=self.fmt_border
        ) + 1  # +1 for blank gap after header

    def _write_content(self, start_row, section_title) -> int:
        """
        Write content section (rows starting at start_row).

        Args:
            start_row: Starting row index (0-based). Should be 4 (row 5) or later.
            section_title: Title text for the section header

        Returns:
            Next available row index
        """
        # Set row heights for content section
        row_heights_offsets = {
            3: 3, 4: 3, 6: 3, 9: 3, 10: 3, 14: 3, 15: 3, 19: 3,
            20: 3, 24: 3, 25: 3, 29: 3, 30: 3, 32: 3, 33: 3, 37: 3,
        }
        for offset, height in row_heights_offsets.items():
            self.ws.set_row(start_row + offset, height)

        # Section header
        self.ws.write(start_row, 2, f"SUMMARY VALUES - {section_title}", self.fmt_section)
        for col in range(3, 25):  # D to Y (columns 4-25, indices 3-24)
            self.ws.write(start_row, col, '', self.fmt_section)

        # Projected header - center across K to T with top border
        self.ws.write(start_row + 1, 10, 'Projected', self.fmt_hdr_white_center_top)
        for col in range(11, 20):  # L to T (columns 11-19)
            self.ws.write(start_row + 1, col, '', self.fmt_hdr_white_center_top)

        # Column headers (year headers row, using dynamic row calculation)
        year_row = start_row + 3  # 1-based Excel row number for year headers

        self.ws.write(start_row + 2, 3, '($ Millions)', self.fmt_hdr_white_border)
        self.ws.write(start_row + 2, 4, '', self.fmt_label_border)
        self.ws.write(start_row + 2, 5, 'Trend', self.fmt_hdr_white_border)
        self.ws.write(start_row + 2, 6, '', self.fmt_label_border)
        self.ws.write(start_row + 2, 7, f'=I{year_row}-1', self.fmt_label_border)
        self.ws.write(start_row + 2, 8, f'=J{year_row}-1', self.fmt_label_border)
        self.ws.write(start_row + 2, 9, f'=K{year_row}-1', self.fmt_label_border)
        self.ws.write(start_row + 2, 10, '=Assumptions!A1', self.fmt_label_border)

        # Year columns from L to T - each references previous column + 1
        for col in range(11, 20):  # L to T (columns 11-19)
            prev_col_letter = chr(ord('A') + col - 1)  # Previous column letter
            formula = f'={prev_col_letter}{year_row}+1'
            self.ws.write(start_row + 2, col, formula, self.fmt_label_border)

        # Dotted border separators (H to T) - at rows 14, 19, 24, 29, 34, 37
        # Offsets from start_row: 10, 15, 20, 25, 30, 33
        for offset in [10, 15, 20, 25, 30, 33]:
            for col in range(7, 20):  # H to T (columns 7-19)
                self.ws.write_blank(start_row + offset, col, None, self.fmt_dashed_border)

        # Income Statement Items section
        self.ws.write(start_row + 5, 3, 'Income Statement Items', self.fmt_section_left)

        self.ws.write(start_row + 7, 4, 'Net Revenue', self.fmt_label)
        self.ws.write(start_row + 8, 4, '   Growth', self.fmt_sub_label)
        self.ws.write(start_row + 11, 4, 'EBITDA', self.fmt_label)
        self.ws.write(start_row + 12, 4, '   Margin', self.fmt_sub_label)
        self.ws.write(start_row + 13, 4, '   Growth', self.fmt_sub_label)
        self.ws.write(start_row + 16, 4, 'Net Income', self.fmt_label)
        self.ws.write(start_row + 17, 4, '   Margin', self.fmt_sub_label)
        self.ws.write(start_row + 18, 4, '   Growth', self.fmt_sub_label)
        self.ws.write(start_row + 21, 4, 'NOPAT', self.fmt_sub_label)
        self.ws.write(start_row + 22, 4, '   Margin', self.fmt_sub_label)
        self.ws.write(start_row + 23, 4, '   Growth', self.fmt_sub_label)
        self.ws.write(start_row + 26, 4, 'D&A', self.fmt_sub_label)
        self.ws.write(start_row + 27, 4, 'Capex', self.fmt_sub_label)
        self.ws.write(start_row + 28, 4, 'NWC', self.fmt_sub_label)
        self.ws.write(start_row + 31, 4, 'Unlevered FCFF', self.fmt_label)
        self.ws.write(start_row + 34, 4, '   Discount Period', self.fmt_label)
        self.ws.write(start_row + 35, 4, '   Discount Factor', self.fmt_label)
        self.ws.write(start_row + 36, 4, 'Present Value of FCF', self.fmt_impl_white)

        # Right column (V) valuation summary
        self.ws.write(start_row + 5, 21, 'Discount Rate', self.fmt_label)
        self.ws.write(start_row + 7, 21, 'Terminal Growth Rate', self.fmt_label)
        self.ws.write(start_row + 8, 21, 'Terminal Value', self.fmt_label)
        self.ws.write(start_row + 11, 21, 'Cumulative PV of FCF', self.fmt_label)
        self.ws.write(start_row + 16, 21, 'PV of Terminal Value', self.fmt_label)
        self.ws.write(start_row + 21, 21, 'Enterprise Value', self.fmt_label)
        self.ws.write(start_row + 22, 21, 'Net Cash', self.fmt_label)
        self.ws.write(start_row + 23, 21, 'Equity Value', self.fmt_label)
        self.ws.write(start_row + 27, 21, 'Shares (MM)', self.fmt_label)
        self.ws.write(start_row + 31, 21, 'Implied Shared Price', self.fmt_impl_white)

        # Draw outside border box
        # Top border (at start_row + 1, columns C-Y)
        self.ws.write(start_row + 1, 2, '', self.fmt_border_top_left)
        for col in range(3, 10):  # D to J
            self.ws.write(start_row + 1, col, '', self.fmt_border_top)
        # K-T already have top border from Projected header format
        for col in range(20, 25):  # U to Y
            self.ws.write(start_row + 1, col, '', self.fmt_border_top)
        self.ws.write(start_row + 1, 24, '', self.fmt_border_top_right)

        # Left and right borders (columns C and Y, rows from start_row+2 to start_row+36)
        for offset in range(2, 37):
            self.ws.write(start_row + offset, 2, '', self.fmt_border_left)  # Column C
            self.ws.write(start_row + offset, 24, '', self.fmt_border_right)  # Column Y

        # Bottom border (at start_row + 37, columns C-Y)
        self.ws.write(start_row + 37, 2, '', self.fmt_border_bottom_left)
        for col in range(3, 25):  # D to Y
            self.ws.write(start_row + 37, col, '', self.fmt_border_bottom)
        self.ws.write(start_row + 37, 24, '', self.fmt_border_bottom_right)

        # Closing border row (B to Z)
        for col in range(1, 26):  # B to Z
            self.ws.write(start_row + 38, col, '', self.fmt_border_thin)

        return start_row + 40  # +40 for content rows (39) + blank gap (1)


