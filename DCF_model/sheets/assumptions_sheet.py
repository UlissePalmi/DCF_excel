"""
Assumptions sheet builder - creates a simple sheet with model assumptions.
"""

import xlsxwriter


class AssumptionsSheetBuilder:
    """Builds a simple Assumptions sheet with model year."""

    def __init__(self, workbook: xlsxwriter.Workbook):
        self.workbook = workbook
        self.ws = None

    def build(self) -> None:
        """Create and populate the Assumptions sheet."""
        self.ws = self.workbook.add_worksheet("Assumptions")
        self.ws.write('A1', 2026)
