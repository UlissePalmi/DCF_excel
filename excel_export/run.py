#!/usr/bin/env python3
"""
Export the YPF DCF model to a formatted Excel file.

Usage (from repo root):
    python excel_export/run.py data/model_data.csv
    python excel_export/run.py data/model_data.csv [output.xlsx]

Output is saved to finished_models/<COMPANY_SHORT>_DCF.xlsx by default.
"""

import sys
import os

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.join(_HERE, "..")

sys.path.insert(0, os.path.join(_ROOT, "DCF_model"))  # YPFModel etc.
sys.path.insert(0, _HERE)                              # ExcelExporter

from ypf_model import YPFModel
from exporter import ExcelExporter, COMPANY_SHORT


def main():
    if len(sys.argv) < 2:
        print("Usage: python excel_export/run.py <data_file> [output.xlsx]")
        sys.exit(1)

    data_file = sys.argv[1]

    if len(sys.argv) > 2:
        output_file = sys.argv[2]
    else:
        out_dir = os.path.join(_ROOT, "finished_models")
        os.makedirs(out_dir, exist_ok=True)
        output_file = os.path.join(out_dir, f"{COMPANY_SHORT}_DCF.xlsx")

    print(f"Loading model from: {data_file}")
    model = YPFModel(data_file)
    print(model)

    print(f"Exporting to:       {output_file} ...")
    ExcelExporter(model, output_file).export()

    print("Done.")


if __name__ == "__main__":
    main()
