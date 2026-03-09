"""
CapIQ data loader for the DUOL DCF Model.

Reads an S&P Capital IQ .xls export (Duolingo financials) and maps CapIQ row
labels to the schedule field keys used by BaseSchedule subclasses.

Returns empty dicts ({}) for fields not present in the CapIQ data so that
schedules whose data is unavailable stay intact but simply show no values.
"""

import re
import xlrd


HISTORICAL_YEARS = [2019, 2020, 2021, 2022, 2023, 2024, 2025]
PROJECTED_YEARS: list[int] = [2026, 2027, 2028, 2029, 2030, 2031, 2032]
ALL_YEARS = HISTORICAL_YEARS + PROJECTED_YEARS


def _extract_year(v) -> int | None:
    """Return a 4-digit year (2000-2099) from strings like '12 months\\nDec-31-2019A'."""
    m = re.search(r'\b(20\d{2})\b', str(v))
    return int(m.group(1)) if m else None


class CapIQLoader:
    """
    Reads an S&P Capital IQ .xls export and exposes data via field(schedule_name, key).

    Interface is compatible with MultiSheetLoader:
        loader.field("Income Statement", "revenue")  -> {2019: 70760.0, ...}
        loader.field("Oil Revenue Schedule", "rev_total")  -> {}  (no data for DUOL)

    Units: all values are kept in the original CapIQ units for the financial
    statement sheets (Thousands of USD). Key Stats values (Millions) are
    multiplied x1000 to match.
    """

    HISTORICAL_YEARS = HISTORICAL_YEARS
    PROJECTED_YEARS  = PROJECTED_YEARS
    ALL_YEARS        = ALL_YEARS

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.company_name: str = ""
        self.ticker: str = ""
        # {capiq_sheet_name: {row_label: {year: float}}}
        self._raw: dict[str, dict[str, dict[int, float]]] = {}
        # {(schedule_name, field_key): {year: float}}
        self._fields: dict[tuple[str, str], dict[int, float]] = {}
        self._load()
        self._build_fields()

    # ── Raw sheet parsing ─────────────────────────────────────────────────────

    def _load(self):
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = xlrd.open_workbook(self.filepath)
        for sname in wb.sheet_names():
            ws = wb.sheet_by_name(sname)
            self._raw[sname] = self._parse_sheet(ws)
            # Parse company name and ticker from the first header line we find,
            if not self.company_name:
                self._parse_company_header(ws)

    def _parse_company_header(self, ws):
        """
        Scan the first ~10 rows for a CapIQ header like:
            "Duolingo, Inc. (NasdaqGS:DUOL) > Financials > ..."
        and extract company_name and ticker.
        """
        for r in range(min(10, ws.nrows)):
            cell = str(ws.cell_value(r, 0))
            m = re.match(r'^(.+?)\s+\([\w]+:([\w]+)\)\s*>', cell)
            if m:
                self.company_name = m.group(1).strip()
                self.ticker = m.group(2).strip()
                return

    def _parse_sheet(self, ws) -> dict[str, dict[int, float]]:
        """
        Parse one CapIQ sheet → {row_label: {year: value}}.

        Finds the year-header row (containing "For the Fiscal Period Ending"
        or "Balance Sheet as of:"), then reads subsequent data rows.
        """

        # A matrix of an entire excel sheet. rows is a list of rows: list[list]
        rows = [
            [ws.cell_value(r, c) for c in range(ws.ncols)]
            for r in range(ws.nrows)
        ]

        # Locate the year-header row
        year_row_idx = None
        for i, row in enumerate(rows):
            label = str(row[0])
            if "For the Fiscal Period Ending" in label or "Balance Sheet as of:" in label:
                year_row_idx = i
                break

        if year_row_idx is None:
            return {}

        # Parse year values from columns 1+
        year_row = rows[year_row_idx]
        years: list[int] = []
        for v in year_row[1:]:
            y = _extract_year(v)
            if y is not None:
                years.append(y)

        # Balance Sheet uses Excel date serial numbers — fall back to canonical list
        if not years:
            count = sum(1 for v in year_row[1:] if v not in ("", None, 0, "-"))
            years = HISTORICAL_YEARS[:count]

        # Parse data rows (skip Currency / Units header rows and blanks)
        skip_labels = {"Currency", "Units", ""}
        data: dict[str, dict[int, float]] = {}
        for row in rows[year_row_idx + 1:]:
            raw_label = row[0]
            stripped = str(raw_label).strip()
            if not stripped or stripped in skip_labels or stripped.startswith("\n"):
                continue
            values: dict[int, float] = {}
            for i, year in enumerate(years):
                raw = row[i + 1] if i + 1 < len(row) else None
                if raw is None or raw == "" or raw == "-":
                    continue
                try:
                    values[year] = float(raw)
                except (TypeError, ValueError):
                    pass
            # Store with the original (un-stripped) label so leading spaces match
            data[str(raw_label)] = values

        return data

    # ── Lookup helpers ────────────────────────────────────────────────────────

    def _raw_series(self, sheet: str, label: str, scale: float = 1.0) -> dict[int, float]:
        """Return a raw {year: value} series, applying an optional unit scale."""
        series = self._raw.get(sheet, {}).get(label, {})
        if scale == 1.0:
            return dict(series)
        return {y: v * scale for y, v in series.items()}

    def _derive_beginning(self, end_series: dict[int, float]) -> dict[int, float]:
        """Given ending-balance series, derive beginning-balance (prior year's ending)."""
        years = sorted(end_series)
        return {years[i]: end_series[years[i - 1]] for i in range(1, len(years))}

    # ── Field mapping ─────────────────────────────────────────────────────────

    def _build_fields(self):
        d = self._fields

        # ── Income Statement ──────────────────────────────────────────────────
        IS = "Income Statement"

        revenue = self._raw_series("Income Statement", "  Revenues")
        cost    = self._raw_series("Income Statement", "  Cost of Revenues")

        d[(IS, "revenue")]      = revenue
        d[(IS, "cost_of_sales")] = cost
        d[(IS, "gross_profit")] = {
            y: revenue.get(y, 0) + cost.get(y, 0) for y in revenue
        }

        sell  = self._raw_series("Income Statement", "  Sales and Marketing")
        admin = self._raw_series("Income Statement", "  General and Administrative")
        rd    = self._raw_series("Income Statement", "  Research and Development")

        d[(IS, "selling_expenses")]     = sell
        d[(IS, "admin_expenses")]       = admin
        d[(IS, "exploration_expenses")] = rd  # closest analog for DUOL
        d[(IS, "other")]                = self._raw_series("Income Statement", "  Other Income")
        d[(IS, "impairment")]           = self._raw_series(
            "Income Statement", "  Impairment of Capitalized Software"
        )
        d[(IS, "financial_income")]     = self._raw_series("Income Statement", "  Interest Income")
        d[(IS, "financial_costs")]      = self._raw_series("Income Statement", "  Other Expenses")
        d[(IS, "ebt")]                  = self._raw_series(
            "Income Statement", "  Earnings before Taxes"
        )
        d[(IS, "income_tax")]   = self._raw_series("Income Statement", "  Provision for Income Tax")
        d[(IS, "net_income")]   = self._raw_series("Income Statement", "  Net Income (Loss)")

        # D&A from the Cash Flow sheet
        d[(IS, "da")] = self._raw_series("Cash Flow", "  Depreciation and Amortization")

        # EBITDA / EBIT from Key Stats (in Millions → ×1000 to match Thousands)
        d[(IS, "ebitda")] = self._raw_series("Key Stats", "EBITDA", scale=1000.0)
        d[(IS, "ebit")]   = self._raw_series("Key Stats", "EBIT",   scale=1000.0)

        # Operating costs = selling + admin + R&D
        all_op_years = set(sell) | set(admin) | set(rd)
        d[(IS, "operating_costs")] = {
            y: sell.get(y, 0) + admin.get(y, 0) + rd.get(y, 0)
            for y in all_op_years
        }

        # Margins from the Ratios sheet (already fractional)
        d[(IS, "gross_margin")]  = self._raw_series("Ratios", "  Gross Margin %")
        d[(IS, "ebitda_margin")] = self._raw_series("Ratios", "  EBITDA Margin %")
        d[(IS, "ebit_margin")]   = self._raw_series("Ratios", "  EBIT Margin %")
        d[(IS, "roe")]           = self._raw_series("Ratios", "  Return on Equity %")

        # Revenue growth YoY
        rev_sorted = sorted(revenue.items())
        d[(IS, "revenue_growth")] = {
            y: (v - rev_sorted[i - 1][1]) / abs(rev_sorted[i - 1][1])
            for i, (y, v) in enumerate(rev_sorted)
            if i > 0 and rev_sorted[i - 1][1]
        }

        # COGS growth YoY
        cogs_sorted = sorted(cost.items())
        d[(IS, "cogs_growth")] = {
            y: (v - cogs_sorted[i - 1][1]) / abs(cogs_sorted[i - 1][1])
            for i, (y, v) in enumerate(cogs_sorted)
            if i > 0 and cogs_sorted[i - 1][1]
        }

        # ── Cash Flow Statement ───────────────────────────────────────────────
        CF = "Cash Flow Statement"

        d[(CF, "net_income")]       = self._raw_series("Cash Flow", "  Net Income")
        d[(CF, "depreciation_ppe")] = self._raw_series(
            "Cash Flow", "  Depreciation and Amortization"
        )
        d[(CF, "impairment")]      = self._raw_series(
            "Cash Flow", "  Impairment of Capitalized Software"
        )
        d[(CF, "fx_interest_other")] = self._raw_series(
            "Cash Flow", "  Stock based Compensation"
        )
        d[(CF, "cf_operating_total")] = self._raw_series(
            "Cash Flow", "  Cash Flow from Operating Activities"
        )

        # Working capital = sum of all WC-change lines in operating activities
        wc_labels = [
            "  Accounts Receivable",
            "  Accounts Payable",
            "  Deferred Revenue",
            "  Prepaid Expenses and Other Current Assets",
            "  Accrued Expenses and Other Current Liabilities",
            "  Deferred Cost of Revenue",
        ]
        wc_parts = [self._raw_series("Cash Flow", lbl) for lbl in wc_labels]
        all_wc_years = set().union(*[s.keys() for s in wc_parts])
        d[(CF, "working_capital")] = {
            y: sum(s.get(y, 0) for s in wc_parts) for y in all_wc_years
        }

        d[(CF, "capex")] = self._raw_series(
            "Cash Flow", "  Purchase of Property and Equipment"
        )

        # Acquisitions: two differently-named rows across different years — merge
        acq1 = self._raw_series(
            "Cash Flow", "  Acquisitions of Companies, Net of Cash Acquired"
        )
        acq2 = self._raw_series(
            "Cash Flow",
            "  Acquisitions of Companies, net of $0 and $5 Cash acquired, respectively",
        )
        d[(CF, "acquisitions_jv")] = {**acq1, **acq2}

        d[(CF, "assets_held_for_sale")] = self._raw_series(
            "Cash Flow", "  Proceeds from Sale of Capitalized Software"
        )
        d[(CF, "cf_investing_total")] = self._raw_series(
            "Cash Flow", "  Cash Flow from Investing Activities"
        )

        d[(CF, "loan_proceeds")] = self._raw_series(
            "Cash Flow", "  Proceeds from Exercise of Stock Options"
        )
        d[(CF, "buyback")] = self._raw_series(
            "Cash Flow", "  Repurchase of Common Stock"
        )
        d[(CF, "cf_financing_total")] = self._raw_series(
            "Cash Flow", "  Cash Flow from Financing Activities"
        )
        d[(CF, "change_in_cash")] = self._raw_series(
            "Cash Flow", "  Cash Flow Net Changes in Cash"
        )

        # Cash position
        cash_end = self._raw_series("Balance Sheet", "  Cash and Cash Equivalents")
        d[(CF, "ending_cash")]   = cash_end
        d[(CF, "beginning_cash")] = self._derive_beginning(cash_end)

        # ── Balance Sheet ─────────────────────────────────────────────────────
        BS = "Balance Sheet"

        d[(BS, "ca_cash")]             = self._raw_series(
            "Balance Sheet", "  Cash and Cash Equivalents"
        )
        d[(BS, "ca_investments")]      = self._raw_series(
            "Balance Sheet", "  Short-term Investments"
        )
        d[(BS, "ca_trade_receivables")] = self._raw_series(
            "Balance Sheet", "  Accounts Receivables"
        )
        d[(BS, "ca_other_receivables")] = self._raw_series(
            "Balance Sheet", "  Prepaid Expenses and Other Current Assets"
        )
        d[(BS, "ca_contract_asset")]   = self._raw_series(
            "Balance Sheet", "  Deferred Cost of Revenue"
        )
        d[(BS, "ca_total")]            = self._raw_series(
            "Balance Sheet", "  Total Current Assets"
        )

        d[(BS, "nca_rou_assets")]             = self._raw_series(
            "Balance Sheet", "  Operating Lease Right-of-use Assets"
        )
        d[(BS, "nca_ppe")]                    = self._raw_series(
            "Balance Sheet", "  Property and Equipment, net"
        )
        d[(BS, "nca_financial_investments")]  = self._raw_series(
            "Balance Sheet", "  Long-term Investments"
        )
        d[(BS, "nca_deferred_tax_asset")]     = self._raw_series(
            "Balance Sheet", "  Deferred Tax Asset-net"
        )
        d[(BS, "nca_intangible")]             = self._raw_series(
            "Balance Sheet", "  Intangible Assets, Net"
        )
        d[(BS, "nca_other_receivables")]      = self._raw_series(
            "Balance Sheet", "  Other Assets"
        )

        d[(BS, "total_assets")] = self._raw_series("Balance Sheet", "  Total Assets")

        d[(BS, "cl_accounts_payable")]     = self._raw_series(
            "Balance Sheet", "  Accounts Payable"
        )
        d[(BS, "cl_other_liabilities")]    = self._raw_series(
            "Balance Sheet", "  Accrued Expenses and Other Current Liabilities"
        )
        d[(BS, "cl_income_tax_payable")]   = self._raw_series(
            "Balance Sheet", "  Income Tax Payable"
        )
        d[(BS, "cl_contract_liabilities")] = self._raw_series(
            "Balance Sheet", "  Deferred Revenue"
        )
        d[(BS, "cl_total")]                = self._raw_series(
            "Balance Sheet", "  Total Current Liabilities"
        )

        d[(BS, "ncl_lease_liabilities")] = self._raw_series(
            "Balance Sheet", "  Long-term Obligations Under Operating Leases"
        )
        d[(BS, "ncl_deferred_tax_liabilities")] = self._raw_series(
            "Balance Sheet", "  Deferred Tax Liabilities Net"
        )

        # Retained earnings: "Accumulated Deficit" label for 2019–2023,
        # "Retained Earnings (Accumulated Deficit)" for 2024–2025 — merge both.
        re_old = self._raw_series("Balance Sheet", "  Accumulated Deficit")
        re_new = self._raw_series(
            "Balance Sheet", "  Retained Earnings (Accumulated Deficit)"
        )
        d[(BS, "eq_retained_earnings")] = {**re_old, **re_new}

        d[(BS, "eq_common_stock")] = self._raw_series(
            "Balance Sheet", "  Common Stock - Par Value"
        )
        d[(BS, "eq_total")] = self._raw_series(
            "Balance Sheet", "  Total Shareholders Equity"
        )
        d[(BS, "total_liabilities_and_equity")] = self._raw_series(
            "Balance Sheet", "  Total Liabilities & Shareholders Equity"
        )

        ta  = d[(BS, "total_assets")]
        tle = d[(BS, "total_liabilities_and_equity")]
        d[(BS, "check")] = {y: round(ta.get(y, 0) - tle.get(y, 0), 2) for y in ta}

        # ── Fixed Assets (PP&E) Schedule ──────────────────────────────────────
        FA = "Fixed Assets (PP&E) Schedule"

        d[(FA, "ppe_ending")]     = self._raw_series(
            "Balance Sheet", "  Property and Equipment, net"
        )
        d[(FA, "rou_ending")]     = self._raw_series(
            "Balance Sheet", "  Operating Lease Right-of-use Assets"
        )
        total_da = self._raw_series("Cash Flow", "  Depreciation and Amortization")
        d[(FA, "total_da")]       = total_da
        d[(FA, "ppe_depr_total")] = total_da

        # ── Working Capital Schedule ──────────────────────────────────────────
        WC = "Working Capital Schedule"

        ca_total = d[(BS, "ca_total")]
        cl_total = d[(BS, "cl_total")]
        nwc = {y: ca_total.get(y, 0) - cl_total.get(y, 0) for y in ca_total}
        d[(WC, "net_working_capital")] = nwc

        nwc_sorted = sorted(nwc.items())
        d[(WC, "change_in_working_capital")] = {
            nwc_sorted[i][0]: nwc_sorted[i][1] - nwc_sorted[i - 1][1]
            for i in range(1, len(nwc_sorted))
        }

        # ── Debt and Interest Schedule ────────────────────────────────────────
        DI = "Debt and Interest Schedule"

        # DUOL has no traditional debt — only operating leases
        d[(DI, "cash_ending")]                 = self._raw_series(
            "Balance Sheet", "  Cash and Cash Equivalents"
        )
        d[(DI, "totals_lt_loans_revolver")]    = self._raw_series(
            "Balance Sheet", "  Long-term Obligations Under Operating Leases"
        )
        d[(DI, "totals_total_loans_revolver")] = self._raw_series(
            "Balance Sheet", "  Long-term Obligations Under Operating Leases"
        )

        # ── Shareholders' Equity Schedule ─────────────────────────────────────
        SE = "Shareholders' Equity Schedule"

        d[(SE, "re_net_income")] = self._raw_series("Income Statement", "  Net Income (Loss)")
        d[(SE, "re_ending")]     = d[(BS, "eq_retained_earnings")]

        # YPF-specific schedules (Oil Revenue, Crude Products, etc.) have no
        # DUOL equivalent — field() will return {} for all their keys by default.

    # ── Public API ─────────────────────────────────────────────────────────────

    def field(self, schedule_name: str, key: str) -> dict[int, float]:
        """
        Return {year: value} for the given schedule + field key.
        Returns {} if the field is not available in the CapIQ data.
        """
        return self._fields.get((schedule_name, key), {})

    @property
    def sheet_names(self) -> list[str]:
        """Names of the raw CapIQ sheets that were parsed."""
        return list(self._raw.keys())

    def __repr__(self) -> str:
        return f"<CapIQLoader: {self.filepath!r}, {len(self._fields)} fields mapped>"
