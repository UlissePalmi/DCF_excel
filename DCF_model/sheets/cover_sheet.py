"""Cover sheet builder - creates the title/cover page for the DCF model."""

import xlsxwriter

class CoverSheetBuilder:
    """Builds a formatted Cover sheet with company info and metadata."""

    # Layout configuration
    TITLE_ROW = 3          # Row 4 (0-based)
    DATE_ROW = 4           # Row 5
    TICKER_ROW = 5         # Row 6
    PREPARED_BY_ROW = 23   # Row 24

    COMPANY_COL = 1        # Column B
    DATE_START_COL = 2     # Column C
    DATE_END_COL = 4       # Column E
    TICKER_COL = 2         # Column C
    PREPARED_BY_COL = 3    # Column D

    def __init__(self, workbook: xlsxwriter.Workbook, company_name: str, ticker: str):
        self.workbook = workbook
        self.company_name = company_name
        self.ticker = ticker
        self.ws = None

    def build(self) -> None:
        """Create and populate the Cover sheet."""
        self.ws = self.workbook.add_worksheet("Cover")
        self.ws.hide_gridlines(2)
        self._setup_layout()
        self._write_content()

    def _setup_layout(self) -> None:
        """Apply column widths and row heights."""
        self.ws.set_column(0, 0, 2.0)    # Column A
        self.ws.set_column(1, 1, 1.0)    # Column B
        self.ws.set_column(2, 2, 3.0)    # Column C
        self.ws.set_row(self.TITLE_ROW, 23.25)  # Title row height

    def _write_content(self) -> None:
        """Write content to the cover sheet."""
        # Create formats
        fmt_title = self.workbook.add_format({
            "font_name": "Arial",
            "font_size": 18,
            "bold": True,
            "align": "left",
            "valign": "top",
        })

        fmt_date = self.workbook.add_format({
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "num_format": "mm/dd/yyyy",
            "align": "left",
        })

        fmt_text = self.workbook.add_format({
            "font_name": "Calibri",
            "font_size": 11,
        })

        # Write title
        self.ws.write(self.TITLE_ROW, self.COMPANY_COL, self.company_name, fmt_title)

        # Write date
        self.ws.merge_range(
            self.DATE_ROW, self.DATE_START_COL,
            self.DATE_ROW, self.DATE_END_COL,
            "=TODAY()",
            fmt_date
        )

        # Write ticker
        self.ws.write(self.TICKER_ROW, self.TICKER_COL, self.ticker, fmt_text)

        # Write prepared by names (4 rows)
        for i in range(4):
            self.ws.write(
                self.PREPARED_BY_ROW + i,
                self.PREPARED_BY_COL,
                "Ulisse Palmiero" if i == 0 else "",
                fmt_text
            )
