"""
Summary sheet builder - creates key metrics summary page.
"""

import xlsxwriter
from .summary_formats import create_summary_formats, draw_border_box, write_summary_content, write_year_headers, adjust_column_widths, adjust_row_heights
from structures.structures import write_header

class SummarySheetBuilder:
    """Builds a formatted Summary sheet with model highlights."""

    def __init__(self, workbook: xlsxwriter.Workbook, company_name: str, num_projected_years: int = 0):
        self.workbook = workbook
        self.company_name = company_name
        self.num_projected_years = num_projected_years
        self.ws = None

    def build(self) -> None:
        """Create and populate the Summary sheet with 3 header/content pairs."""
        self.ws = self.workbook.add_worksheet("Summary")
        self.ws.hide_gridlines(2)
        self.ws.set_column('A:A', 2)
        self._load_formats()

        row = 0
        # Calculate right-side start column (1 column after left content ends)
        start_col_left = 1
        start_col_right = start_col_left + self.num_projected_years + 16

        # First header/content pair (Base Case - with side-by-side layout)
        row = self._write_header(start_row=row, subtitle='Base Case DCF')
        self._write_content(start_row=row, section_title='BASE CASE', start_col=start_col_left, adj_col_widths=True)
        self._write_content(start_row=row, section_title='linked', start_col=start_col_right, adj_col_widths=True)
        row += 40

        # Second header/content pair
        row = self._write_header(start_row=row, subtitle='Best Case DCF')
        row = self._write_content(start_row=row, section_title='BEST CASE')

        # Third header/content pair
        row = self._write_header(start_row=row, subtitle='Worse Case DCF')
        row = self._write_content(start_row=row, section_title='WORSE CASE')

    def _load_formats(self) -> None:
        """Load format objects from summary_formats module."""
        formats = create_summary_formats(self.workbook)
        for name, fmt in formats.items():
            setattr(self, name, fmt)

    def _write_header(self, start_row: int, subtitle: str) -> int:
        """
        Write header section (3 rows: title, subtitle, border).

        Args:
            start_row: Starting row index (0-based)
            subtitle: Subtitle text to display

        Returns:
            Next available row index
        """
        return write_header(
            self.ws,
            self.workbook,
            self.company_name,
            subtitle,
            start_row=start_row,
            col_start=1,
            col_end= self.num_projected_years + 15,
        ) + 1

    def _write_content(self, start_row, section_title, start_col: int = 1, adj_col_widths: bool = False) -> int:
        """
        Write content section (rows starting at start_row).

        Args:
            start_row: Starting row index (0-based). Should be 4 (row 5) or later.
            section_title: Title text for the section header
            start_col: Starting column index (default 1 = column B)
            adjust_column_widths: Whether to adjust column widths (default False)

        Returns:
            Next available row index
        """
        end_proj = start_col + self.num_projected_years + 9

        # Adjust row heights (for all content sections)
        adjust_row_heights(self, start_row)

        # Adjust column widths (only for initial calls)
        if adj_col_widths:
            adjust_column_widths(self, start_col, end_proj)

        # Section header
        self.ws.write(start_row, start_col + 1, f"SUMMARY VALUES - {section_title}", self.fmt_section)
        for col in range(start_col + 2, end_proj + 5):
            self.ws.write(start_row, col, '', self.fmt_section)


        # Write year headers (historical and projected)
        write_year_headers(self, start_row, start_col, end_proj)

        # Dotted border separators across data columns
        # Offsets from start_row: 9, 14, 19, 24, 29, 32
        for offset in [9, 14, 19, 24, 29, 32]:
            for col in range(start_col + 6, end_proj):
                self.ws.write_blank(start_row + offset, col, None, self.fmt_dashed_border)

        # Write summary content (income statement and valuation summary)
        write_summary_content(self, start_row, start_col, end_proj)

        # Draw border box around content (closing border only on left side)
        draw_border_box(self, start_row, start_col, end_proj, draw_closing_border=(start_col == 1))

        return start_row + 40  # +40 for content rows (39) + blank gap (1)


