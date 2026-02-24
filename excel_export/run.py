#!/usr/bin/env python3
"""
Export the YPF DCF model to a formatted Excel file.

Usage (from repo root):
    python excel_export/run.py data/model_data.csv [output.xlsx]
"""

import sys
import os

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, "..", "DCF_model"))  # YPFModel etc.
sys.path.insert(0, _HERE)                                   # ExcelExporter

from ypf_model import YPFModel
from exporter import ExcelExporter


def main():
    if len(sys.argv) < 2:
        print("Usage: python excel_export/run.py <data_file> [output.xlsx]")
        sys.exit(1)

    data_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else "YPF_DCF_output.xlsx"

    print(f"Loading model from: {data_file}")
    model = YPFModel(data_file)
    print(model)

    print(f"Exporting to:       {output_file} ...")
    exporter = ExcelExporter(model, output_file)
    exporter.export()

    print("Done.")


if __name__ == "__main__":
    main()
