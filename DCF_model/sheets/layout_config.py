"""
Layout configuration for all Excel sheets.

Defines column widths, row heights, and gridline settings for consistent formatting.
"""


class LayoutConfig:
    """Central configuration for sheet layout properties."""

    # Default settings for all sheets
    DEFAULT_COLUMN_WIDTH = 8.5
    DEFAULT_ROW_HEIGHT = 12.75
    HIDE_GRIDLINES = True

    # Column width definitions (column_index -> width)
    COLUMN_WIDTHS = {
        # Column A (0)
        0: 1.0,
        # Column B (1)
        1: 1.71,
        # Column C (2)
        2: 11.0,
        # Column D (3)
        3: 13.0,
        # Column E onwards (4+): 8.5 each (for data years)
    }

    # Row height definitions (row_index -> height)
    ROW_HEIGHTS = {
        # Title/header rows (slightly taller)
        0: 18.75,    # Row 1
        1: 18.75,    # Row 2
        3: 12.75,    # Row 4
        4: 12.75,    # Row 5
        6: 12.75,    # Row 7
        7: 12.75,    # Row 8
        9: 18.75,    # Row 10 - section header
    }

    @staticmethod
    def apply_to_sheet(ws, hide_gridlines: bool = True, column_count: int = 20) -> None:
        """
        Apply layout configuration to a worksheet.

        Args:
            ws: xlsxwriter worksheet
            hide_gridlines: Whether to hide gridlines
            column_count: Number of columns to configure (default 20 for year data)
        """
        # Hide gridlines if requested
        if hide_gridlines:
            ws.hide_gridlines(2)

        # Set all column widths (0-based indexing)
        for col_idx in range(column_count):
            if col_idx in LayoutConfig.COLUMN_WIDTHS:
                width = LayoutConfig.COLUMN_WIDTHS[col_idx]
            else:
                # Default width for data columns (E onwards)
                width = LayoutConfig.DEFAULT_COLUMN_WIDTH

            ws.set_column(col_idx, col_idx, width)

        # Set specific row heights
        for row_idx, height in LayoutConfig.ROW_HEIGHTS.items():
            ws.set_row(row_idx, height)

    @staticmethod
    def apply_header_rows(ws, header_rows: list[int]) -> None:
        """
        Apply taller height to specific header rows.

        Args:
            ws: xlsxwriter worksheet
            header_rows: List of row indices (0-based) to set as headers
        """
        for row_idx in header_rows:
            ws.set_row(row_idx, 18.75)

    @staticmethod
    def apply_spacer_rows(ws, spacer_rows: list[int], height: float = 3.0) -> None:
        """
        Apply separator row heights (thin spacers).

        Args:
            ws: xlsxwriter worksheet
            spacer_rows: List of row indices (0-based) for spacers
            height: Height for spacer rows (default 3pt)
        """
        for row_idx in spacer_rows:
            ws.set_row(row_idx, height)
