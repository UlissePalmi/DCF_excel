# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python DCF (Discounted Cash Flow) financial model for energy companies. The primary model (`DCF_model/`) is built around **YPF** (Argentine oil company) and reads data from an Excel/CSV "Model" sheet to reconstruct all financial schedules in Python. `excel_export/` writes the model output back to a formatted Excel file that replicates the source layout.  The `old_code/` directory contains an earlier prototype targeting **Amplify Energy Corp.** and is a useful reference for xlsxwriter patterns.

## Commands

Python is not on PATH — always use the venv executable:

```bash
# Export model to Excel  (output → finished_models/YPF_DCF.xlsx)
.venv/Scripts/python.exe excel_export/run.py data/model_data.csv

# Print all schedule outputs to terminal
.venv/Scripts/python.exe DCF_model/demo.py data/model_data.csv
```

No build step or test suite exists. Dependencies (`openpyxl`, `pandas`, `xlsxwriter`, Python ≥3.12) are declared in `pyproject.toml` and installed in `.venv/`.

## Architecture

### Data Flow
```
Excel/CSV file → DataLoader → BaseSchedule subclasses → YPFModel → ExcelExporter → finished_models/YPF_DCF.xlsx
```

### Key Design Pattern: Row-Mapped Schedules
All source data lives in an Excel sheet named **"Model"**. `DataLoader` reads it into a `dict[(row, col)] → value` using 1-indexed coordinates. Column H=8 maps to year 2020, incrementing through column V=22 for 2034.

Each schedule class inherits from `BaseSchedule` and defines `ROW_*` class-level constants pointing to specific Excel row numbers. Properties call `self._series(ROW_*)` to return `{year: value}` dicts. **Adding or changing a line item requires knowing its row number in the source Excel sheet.**

### `DCF_model/` Module Structure
- [DCF_model/data_loader.py](DCF_model/data_loader.py) — loads CSV or Excel into `(row, col)` dict; provides `get_by_year()`, `get_row_series()`, `get_historical()`, `get_projected()`
- [DCF_model/base_schedule.py](DCF_model/base_schedule.py) — base class with `_series()`, `_hist()`, `_proj()` helpers; all subclasses must implement `summary() -> dict`
- [DCF_model/schedules.py](DCF_model/schedules.py) — 14 schedule classes spanning Excel rows 1–764: Oil Revenue (1–30), Crude Products Revenue (31–105), Other Products Revenue (107–167), Downstream Revenue (169–214), Total Revenue (216–277), Production Costs (279–332), S&A Expenses (334–390), Income Statement (393–437), Cash Flow Statement (441–491), Balance Sheet (494–557), Fixed Assets (560–607), Working Capital (609–680), Debt & Interest (682–729), Shareholders' Equity (732–764)
- [DCF_model/ypf_model.py](DCF_model/ypf_model.py) — orchestrator; instantiates all 14 schedules, exposes them as named attributes and via `model.all_schedules`; `model.summary()` returns the full nested dict
- [DCF_model/demo.py](DCF_model/demo.py) — standalone script; uses `sys.path.insert` so it can be run directly

### `excel_export/` Module Structure
- [excel_export/exporter.py](excel_export/exporter.py) — `ExcelExporter` class; writes one "Model" sheet using xlsxwriter, replicating the source YPF DCF.xlsx layout
- [excel_export/run.py](excel_export/run.py) — CLI entry point; creates `finished_models/` if needed, saves as `{COMPANY_SHORT}_DCF.xlsx`

#### Excel Layout Constants (exporter.py)
- `COMPANY_NAME` / `COMPANY_SHORT` — full name written in the sheet; short name used for the filename
- `HIST_YEARS = {2020..2024}` — rendered in blue font (`#0000FF`) with `"A"` suffix on year headers
- Column indices `COL_A=0` … `COL_DATA_0=7` … `COL_DATA_END=21` match source Excel cols A–V

#### Critical xlsxwriter Pattern: `center_across`
To span text across a column range, write the text in the first cell **and then write blank cells with the same `center_across` format** across every column in the range. Writing only the first cell is not sufficient — xlsxwriter will not extend the span reliably. See `_center_across()` in `exporter.py` and `center_across_range()` in `old_code/xslx_setup.py`.

### Import Note
Files inside `DCF_model/` use **relative imports without the package prefix** (e.g., `from data_loader import DataLoader`). This works when running scripts directly from that directory. `__init__.py` uses the correct `from .ypf_model import YPFModel` style for package imports.

### `old_code/` (legacy reference)
`rev_build.py` + `xslx_setup.py` target **Amplify Energy Corp.** and write directly to Excel via xlsxwriter. Not integrated with `DCF_model/`, but the xlsxwriter helper functions in `xslx_setup.py` (`center_across_range`, `row_col_fmt`) are the canonical reference for formatting patterns used in `excel_export/`.

## Data Files
- `data/model_data.csv` — CSV export of the YPF Model sheet; primary input for `DataLoader`
- `data/AMPY_data.xlsx` — source data for the old Amplify Energy prototype
- `finished_models/YPF_DCF.xlsx` — Excel output generated by `excel_export/run.py`
