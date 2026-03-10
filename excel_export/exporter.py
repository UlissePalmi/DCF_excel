"""
Excel exporter for the DCF Model.

Replicates the layout of the source YPF DCF.xlsx workbook:
- Multiple sheets: Cover, Summary, Model
- Per-schedule header block: company name / schedule title / separator
- Year headers formatted as "2020A" (historical) and "2025E" (projected)
- Historical data cells rendered in blue font
- Column structure: A-G are spacers/labels, H-V are data
"""

import sys
import os

# Add DCF_model to path for sheet builders
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'DCF_model'))

import xlsxwriter
from sheets import CoverSheetBuilder, SummarySheetBuilder, AssumptionsSheetBuilder
from sheets.layout_config import LayoutConfig

# Column indices (0-based)
COL_A        = 0   # A: narrow sentinel  (width 3.71)
COL_B        = 1   # B: spacer           (width 1.71)
COL_LABEL    = 2   # C: main label       (width 40)
COL_D        = 3   # D: sub-label        (width 11.71, blank in output)
COL_E        = 4   # E: unused           (width 14.29)
COL_UNIT     = 5   # F: unit label       (width 10.43, blank for now)
COL_G        = 6   # G: pre-data spacer  (width 1.71)
COL_DATA_0   = 7   # H: first data year  (width 9.71)


def _is_series(val) -> bool:
    """Return True if val is a {year: float} series dict."""
    return isinstance(val, dict) and bool(val) and all(isinstance(k, int) for k in val)


def _flatten(data: dict, depth: int = 0):
    """
    Recursively yield (depth, label, series_or_None) from a nested summary dict.

    Yields (depth, label, None)         – intermediate dicts  → section header row
    Yields (depth, label, {year: val})  – leaf year-series    → data row
    """
    for key, val in data.items():
        label = key.replace("_", " ").title()
        if _is_series(val):
            yield depth, label, val
        elif isinstance(val, dict):
            yield depth, label, None
            yield from _flatten(val, depth + 1)


def _center_across(ws, row: int, col_start: int, col_end: int, text, fmt):
    """
    Write text in col_start and blanks through col_end, all with the same
    center_across format.  This is the only way to make center_across span
    reliably all the way to col_end in xlsxwriter (mirrors old_code technique).
    """
    ws.write(row, col_start, text, fmt)
    for c in range(col_start + 1, col_end + 1):
        ws.write_blank(row, c, None, fmt)


def _num_fmt_key(label: str, series: dict) -> str:
    """
    Return one of "int" | "dec" | "pct" based on label keywords and value range.
    """
    lo = label.lower()
    if any(x in lo for x in ("margin", "growth", "yield", "roe", "return on", "rate")):
        return "pct"
    vals = [v for v in series.values() if isinstance(v, (int, float))]
    if vals and all(-2.5 <= v <= 2.5 for v in vals):
        return "pct"
    if any(x in lo for x in ("pricing", "price", "$/", "per boe", "per bbl")):
        return "dec"
    return "int"


class ExcelExporter:
    """Exports a ScheduleBuilder model to a single-sheet formatted Excel file."""

    _C_HIST_FONT = "#0000FF"   # blue – historical hard-coded input cells
    _C_FONT      = "Calibri"

    def __init__(self, model, output_path: str, num_projected_years: int = 0):
        self.model = model
        self.output_path = output_path
        self.company_name = getattr(model.loader, "company_name", None)
        self.ticker = getattr(model.loader, "ticker", None)
        self.num_projected_years = num_projected_years

        # Derive year lists from loader
        self.all_years = model.loader.ALL_YEARS
        self.hist_years = set(model.loader.HISTORICAL_YEARS)
        self.n_years = len(self.all_years)
        self.col_data_end = COL_DATA_0 + self.n_years - 1
        self.col_proj_0 = min(
            COL_DATA_0 + len(model.loader.HISTORICAL_YEARS),
            self.col_data_end + 1
        )

    def export(self) -> str:
        wb = xlsxwriter.Workbook(self.output_path)

        # Build supplementary sheets first (Cover, Summary, Assumptions)
        CoverSheetBuilder(wb, self.company_name, self.ticker).build()
        SummarySheetBuilder(wb, self.company_name, self.all_years, self.num_projected_years).build()
        AssumptionsSheetBuilder(wb).build()

        # Build Model sheet
        ws = wb.add_worksheet("Model")
        LayoutConfig.apply_to_sheet(ws, hide_gridlines=True, column_count=self.col_data_end + 1)

        fmts = self._make_formats(wb)

        row = 0
        for schedule in self.model.all_schedules:
            row = self._write_schedule(ws, fmts, schedule, row)

        wb.close()
        return self.output_path

    # ── Format factory ──────────────────────────────────────────────────────

    def _make_formats(self, wb) -> dict:
        base = {"font_name": self._C_FONT, "font_size": 11}
        f = {}

        # Company name – Calibri 18, bold, center-across
        f["company"] = wb.add_format({
            **base, "font_size": 18, "bold": True,
            "align": "center_across", "valign": "vcenter",
        })

        # Schedule title – Calibri 14, bold, center-across
        f["sched"] = wb.add_format({
            **base, "font_size": 14, "bold": True,
            "align": "center_across", "valign": "vcenter",
        })

        # Separator row – medium bottom border (written as blanks)
        f["sep"] = wb.add_format({"bottom": 2})

        # Closing border row at the bottom of each schedule (B:V)
        f["end"] = wb.add_format({"bottom": 1})

        # "Projected" label above year headers
        f["proj_hdr"] = wb.add_format({**base, "bold": True, "align": "center_across"})

        # Year headers: 2020A (historical) and 2025E (projected)
        f["yr_hist"] = wb.add_format({
            **base, "bold": True, "align": "center",
            "num_format": '0"A"',
        })
        f["yr_proj"] = wb.add_format({
            **base, "bold": True, "align": "center",
            "num_format": '0"E"', "bottom": 1,
        })

        # Section header labels (bold, no fill, two indent depths)
        f["sec0"] = wb.add_format({**base, "bold": True, "indent": 1})
        f["sec1"] = wb.add_format({**base, "bold": True, "indent": 2})

        # Data labels (regular, four indent depths)
        for i in range(4):
            f[f"lbl{i}"] = wb.add_format({**base, "indent": i + 1})

        # Numeric cells for three number formats × two font colours
        _NUM_FMTS = {
            "int": '#,##0_);(#,##0);-_)',
            "dec": '#,##0.0_);(#,##0.0);-_)',
            "pct": '0.0%;(0.0%)',
        }
        for key, num_fmt in _NUM_FMTS.items():
            f[f"hist_{key}"] = wb.add_format({
                **base, "num_format": num_fmt, "align": "right",
                "font_color": self._C_HIST_FONT,
            })
            f[f"proj_{key}"] = wb.add_format({
                **base, "num_format": num_fmt, "align": "right",
            })

        return f

    # ── Schedule writer ──────────────────────────────────────────────────────

    def _write_schedule(self, ws, fmts, schedule, start_row: int) -> int:
        row = start_row

        # ① Spacer row (blank, 12.75 pt)
        ws.set_row(row, 12.75)
        row += 1

        # ② Company title row (18 pt tall)
        _center_across(ws, row, COL_LABEL, self.col_data_end, self.company_name, fmts["company"])
        ws.set_row(row, 23.25)
        row += 1

        # ③ Schedule title row (18.75 pt tall)
        _center_across(ws, row, COL_LABEL, self.col_data_end, schedule.SCHEDULE_NAME, fmts["sched"])
        ws.set_row(row, 18.75)
        row += 1

        # ④ Separator row – 3 pt, medium bottom border across label + data cols
        ws.set_row(row, 3)
        for c in range(COL_LABEL, self.col_data_end + 1):
            ws.write_blank(row, c, None, fmts["sep"])
        row += 1

        # ⑤ "Projected" label above the projected year headers (skip if none)
        ws.set_row(row, 12.75)
        if self.col_proj_0 <= self.col_data_end:
            _center_across(ws, row, self.col_proj_0, self.col_data_end, "Projected", fmts["proj_hdr"])
        row += 1

        # ⑥ Year header row
        ws.set_row(row, 12.75)
        for i, year in enumerate(self.all_years):
            key = "yr_hist" if year in self.hist_years else "yr_proj"
            ws.write_number(row, COL_DATA_0 + i, year, fmts[key])
        row += 1

        # ⑦ Empty spacer row before data
        ws.set_row(row, 12.75)
        row += 1

        # ⑧ Data rows
        for depth, label, series in _flatten(schedule.summary()):
            ws.set_row(row, 12.75)
            if series is None:
                # Section header row
                d = min(depth, 1)
                ws.write(row, COL_LABEL, label, fmts[f"sec{d}"])
            else:
                # Data row
                d = min(depth, 3)
                ws.write(row, COL_LABEL, label, fmts[f"lbl{d}"])
                fmt_key = _num_fmt_key(label, series)
                for i, year in enumerate(self.all_years):
                    v = series.get(year)
                    if v is not None and isinstance(v, (int, float)):
                        prefix = "hist" if year in self.hist_years else "proj"
                        ws.write_number(
                            row, COL_DATA_0 + i, float(v),
                            fmts[f"{prefix}_{fmt_key}"],
                        )
            row += 1

        # Closing border row – medium bottom border from col B to col V
        ws.set_row(row, 12.75)
        for c in range(COL_B, self.col_data_end + 1):
            ws.write_blank(row, c, None, fmts["end"])
        row += 1

        # Extra empty row between schedules
        ws.set_row(row, 12.75)
        row += 1

        return row
