"""
Summary sheet builder - creates key metrics summary page.
"""

import xlsxwriter
from .layout_config import LayoutConfig

# Year columns (left mirror and right source)
LEFT_COLS = ['H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']
RIGHT_COLS = ['AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT']
PROJ_L = ['K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']
PROJ_R = ['AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT']

# Number formats
FMT_DOLLAR = '"$"#,##0_);\\("$"#,##0\\)'
FMT_DOLLAR2 = '"$"#,##0.00_);\\("$"#,##0.00\\)'
FMT_PCT = '0.0%;\\(0.0%\\)'
FMT_PCT_S = '0.0%'
FMT_YEAR = '0\\A'
FMT_DEC1 = '#,##0.0_);\\(#,##0.0\\)'
FMT_DEC1S = '#,##0.0_);\\(#,##0.0\\)'
FMT_DEC2 = '#,##0.00_);\\(#,##0.00\\)'
FMT_DEC2S = '#,##0.00_);(#,##0.00)'
FMT_SHARES = '#,##0.0'


class SummarySheetBuilder:
    """Builds a formatted Summary sheet with model highlights."""

    def __init__(self, workbook: xlsxwriter.Workbook, company_name: str, all_years: list):
        self.workbook = workbook
        self.company_name = company_name
        self.all_years = all_years
        self.ws = None

    def build(self) -> None:
        """Create and populate the Summary sheet."""
        self.ws = self.workbook.add_worksheet("Summary")
        n_years = len(self.all_years)
        LayoutConfig.apply_to_sheet(self.ws, hide_gridlines=True, column_count=4 + n_years)
        self._setup_layout()
        self._create_formats()
        self._write_base_case()

    def _setup_layout(self) -> None:
        """Apply custom column widths and row heights."""
        col_widths = {
            'A': 2, 'B': 4, 'C': 1, 'D': 1, 'E': 20, 'G': 1, 'U': 1, 'V': 10, 'Y': 1,
            'AC': 1, 'AD': 1, 'AE': 20, 'AG': 1, 'AU': 1, 'AV': 10, 'AY': 1,
        }

        for col, width in col_widths.items():
            self.ws.set_column(col + ':' + col, width)

        row_heights = {
            1: 23.25, 2: 18, 3: 3, 8: 3, 9: 3, 11: 3, 14: 3, 15: 3,
            19: 3, 20: 3, 24: 3, 25: 3, 29: 3, 30: 3, 34: 3, 35: 3,
            37: 3, 38: 3, 42: 3, 45: 23.25, 46: 18, 47: 3, 49: 18.75,
            52: 3, 53: 3, 55: 3, 58: 3, 59: 3, 63: 3, 64: 3, 68: 3,
            69: 3, 73: 3, 74: 3, 78: 3, 79: 3, 81: 3, 82: 3, 86: 3,
            89: 23.25, 90: 18, 91: 3, 93: 18.75, 96: 3, 97: 3, 99: 3,
            102: 3, 103: 3, 107: 3, 108: 3, 112: 3, 113: 3, 117: 3,
            118: 3, 122: 3, 123: 3, 125: 3, 126: 3, 130: 3,
        }

        for row, height in row_heights.items():
            self.ws.set_row(row - 1, height)

    def _create_formats(self) -> None:
        """Create all format objects for the Summary sheet."""
        NAVY = '002F6C'

        self.fmt_title = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 18, 'bold': True,
            'align': 'center_across'
        })

        self.fmt_subtitle = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 14, 'bold': True,
            'align': 'center_across'
        })

        self.fmt_section = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10, 'bold': True
        })

        self.fmt_hdr_white = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10, 'bold': True,
            'font_color': 'FFFFFF', 'bg_color': NAVY, 'align': 'vcenter'
        })

        self.fmt_hdr_white_center = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10, 'bold': True,
            'font_color': 'FFFFFF', 'bg_color': NAVY, 'align': 'center_across'
        })

        self.fmt_hdr_white_year = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10, 'bold': True,
            'font_color': 'FFFFFF', 'bg_color': NAVY, 'align': 'vcenter',
            'num_format': FMT_YEAR
        })

        self.fmt_label = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10
        })

        self.fmt_sub_label = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 9
        })

        self.fmt_sub_lbl8 = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 8
        })

        self.fmt_navy = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10, 'font_color': NAVY
        })

        self.fmt_navy_dollar = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10, 'font_color': NAVY,
            'num_format': FMT_DOLLAR
        })

        self.fmt_navy_pct = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10, 'font_color': NAVY,
            'num_format': FMT_PCT_S
        })

        self.fmt_navy_dec1 = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10, 'font_color': NAVY,
            'num_format': FMT_DEC1
        })

        self.fmt_navy_dec2 = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10, 'font_color': NAVY,
            'num_format': FMT_DEC2
        })

        self.fmt_navy_shares = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10, 'font_color': NAVY,
            'num_format': FMT_SHARES
        })

        self.fmt_navy9 = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 9, 'font_color': NAVY
        })

        self.fmt_navy9_pct = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 9, 'font_color': NAVY,
            'num_format': FMT_PCT
        })

        self.fmt_impl_white = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10, 'bold': True,
            'font_color': 'FFFFFF', 'bg_color': NAVY
        })

        self.fmt_impl_white_dollar2 = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10, 'bold': True,
            'font_color': 'FFFFFF', 'bg_color': NAVY,
            'num_format': FMT_DOLLAR2
        })

        self.fmt_impl_dark = self.workbook.add_format({
            'font_name': 'Arial', 'font_size': 10, 'bold': True,
            'font_color': '404040'
        })

        self.fmt_pv_bold = self.workbook.add_format({
            'font_name': 'Calibri', 'font_size': 11, 'bold': True
        })

        self.fmt_border = self.workbook.add_format({
            'bottom': 2
        })

        self.fmt_border_navy = self.workbook.add_format({
            'bottom': 2, 'font_color': NAVY
        })

    def _write_base_case(self) -> None:
        """Write the Base Case section of the Summary sheet."""
        # Title section - center across B to Z
        self.ws.write('B1', self.company_name, self.fmt_title)
        for col in range(2, 26):  # C to Z (columns 3-26)
            self.ws.write(0, col, '', self.fmt_title)

        self.ws.write('B2', 'Base Case DCF', self.fmt_subtitle)
        for col in range(2, 26):  # C to Z
            self.ws.write(1, col, '', self.fmt_subtitle)

        # Row 3 border
        for col in range(1, 26):  # B to Z
            self.ws.write(2, col, '', self.fmt_border)

        self.ws.write('C5', 'SUMMARY VALUES - BASE CASE', self.fmt_section)

        # Projected header - center across K to T
        self.ws.write('K6', 'Projected', self.fmt_hdr_white_center)
        for col in range(11, 20):  # L to S (columns 12-19, indices 11-19)
            self.ws.write(5, col, '', self.fmt_hdr_white_center)
        self.ws.write('D7', '($ Millions)', self.fmt_hdr_white)
        self.ws.write('F7', 'Trend', self.fmt_hdr_white)

        # Income Statement Items
        self.ws.write('D10', 'Income Statement Items', self.fmt_section)
        
        self.ws.write('E12', 'Net Revenue', self.fmt_label)
        self.ws.write('E13', '   Growth', self.fmt_sub_label)
        self.ws.write('E16', 'EBITDA', self.fmt_label)
        self.ws.write('E17', '   Margin', self.fmt_sub_label)
        self.ws.write('E18', '   Growth', self.fmt_sub_label)
        self.ws.write('E21', 'Net Income', self.fmt_label)
        self.ws.write('E22', '   Margin', self.fmt_sub_label)
        self.ws.write('E23', '   Growth', self.fmt_sub_label)
        self.ws.write('E26', 'NOPAT', self.fmt_sub_label)
        self.ws.write('E27', '   Margin', self.fmt_sub_label)
        self.ws.write('E28', '   Growth', self.fmt_sub_label)
        self.ws.write('E31', 'D&A', self.fmt_sub_label)
        self.ws.write('E32', 'Capex', self.fmt_sub_label)
        self.ws.write('E33', 'NWC', self.fmt_sub_label)
        self.ws.write('E36', 'Unlevered FCFF', self.fmt_label)
        self.ws.write('E39', '   Discount Period', self.fmt_label)
        self.ws.write('E40', '   Discount Factor', self.fmt_label)
        self.ws.write('E41', 'Present Value of FCF', self.fmt_pv_bold)
        self.ws.write('V10', 'Discount Rate', self.fmt_navy)
        self.ws.write('V12', 'Terminal Growth Rate', self.fmt_navy)
        self.ws.write('V13', 'Terminal Value', self.fmt_navy)
        self.ws.write('V16', 'Cumulative PV of FCF', self.fmt_navy)
        self.ws.write('V21', 'PV of Terminal Value', self.fmt_navy)
        self.ws.write('V27', 'Net Cash', self.fmt_navy)
        self.ws.write('V28', 'Equity Value', self.fmt_navy)
        self.ws.write('V32', 'Shares (MM)', self.fmt_navy)
        self.ws.write('V36', 'Implied Shared Price', self.fmt_impl_white)


