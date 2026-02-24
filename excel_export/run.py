#!/usr/bin/env python3
"""
Export a DCF model to a formatted Excel file.

Usage (from repo root):
    python excel_export/run.py <ticker>

Looks for data/<ticker>_historicals.csv or data/<ticker>_historicals.xlsx.
Output is saved to finished_models/<ticker>_DCF.xlsx.
"""

import sys
import os

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.join(_HERE, "..")

sys.path.insert(0, os.path.join(_ROOT, "DCF_model"))
sys.path.insert(0, _HERE)

from ypf_model import YPFModel
from exporter import ExcelExporter


def find_data_file(ticker: str) -> str:
    data_dir = os.path.join(_ROOT, "data")
    for ext in ("csv", "xlsx"):
        path = os.path.join(data_dir, f"{ticker}_historicals.{ext}")
        if os.path.exists(path):
            return path
    raise FileNotFoundError(
        f"No data file found for '{ticker}' in data/. "
        f"Expected {ticker}_historicals.csv or {ticker}_historicals.xlsx"
    )


def main():
    if len(sys.argv) != 2:
        print("Usage: python excel_export/run.py <ticker>")
        sys.exit(1)

    ticker = sys.argv[1]
    data_file = find_data_file(ticker)

    out_dir = os.path.join(_ROOT, "finished_models")
    os.makedirs(out_dir, exist_ok=True)
    output_file = os.path.join(out_dir, f"{ticker}_DCF.xlsx")

    print(f"Ticker:       {ticker}")
    print(f"Data file:    {data_file}")
    print(f"Exporting to: {output_file} ...")

    model = YPFModel(data_file)
    print(model)

    ExcelExporter(model, output_file).export()
    print("Done.")


if __name__ == "__main__":
    main()
