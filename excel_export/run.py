#!/usr/bin/env python3
"""
Export the DUOL DCF model to a formatted Excel file.

Usage (from repo root):
    python excel_export/run.py
    python excel_export/run.py "Duolingo Inc NasdaqGS DUOL Financials.xls"

Output is saved to finished_models/DUOL_DCF.xlsx.
"""

import sys
import os

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.join(_HERE, "..")

sys.path.insert(0, os.path.join(_ROOT, "DCF_model"))
sys.path.insert(0, _HERE)

from model.schedule_builder import ScheduleBuilder
from exporter import ExcelExporter


def find_data_file() -> str:
    candidates = [
        os.path.join(_ROOT, "data", "Duolingo Inc NasdaqGS DUOL Financials.xls"),
        os.path.join(_ROOT, "Duolingo Inc NasdaqGS DUOL Financials.xls"),
        "Duolingo Inc NasdaqGS DUOL Financials.xls",
    ]
    for path in candidates:
        if os.path.exists(path):
            return os.path.normpath(path)
    raise FileNotFoundError(
        "DUOL CapIQ file not found. Place 'Duolingo Inc NasdaqGS DUOL Financials.xls' "
        "in the data/ folder, or pass the path as an argument."
    )


def main():
    if len(sys.argv) > 1:
        data_file = sys.argv[1]
    else:
        data_file = find_data_file()

    out_dir = os.path.join(_ROOT, "finished_models")
    os.makedirs(out_dir, exist_ok=True)

    print(f"Data file:    {data_file}")

    model = ScheduleBuilder(data_file)
    print(model)

    ticker = model.loader.ticker or "DCF"
    output_file = os.path.join(out_dir, f"{ticker}_DCF.xlsx")
    print(f"Company:      {model.loader.company_name}")
    print(f"Ticker:       {ticker}")
    print(f"Exporting to: {output_file} ...")

    ExcelExporter(model, output_file).export()
    print("Done.")


if __name__ == "__main__":
    main()
