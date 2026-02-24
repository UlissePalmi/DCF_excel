#!/usr/bin/env python3
"""
Demo script – loads the YPF Model from CSV or Excel and prints key outputs.
"""

import sys
import os

# Allow running from this directory or parent
sys.path.insert(0, os.path.dirname(__file__))

from data_loader import DataLoader
from ypf_model import YPFModel


def fmt(val, width=12):
    """Format a number for aligned table display."""
    if val is None:
        return " " * width
    if isinstance(val, float):
        return f"{val:>{width},.1f}"
    return f"{str(val):>{width}}"


def print_table(title: str, data: dict[int, float], indent=2):
    """Print a {year: value} dict as a single-row table."""
    prefix = " " * indent
    years = sorted(data.keys())
    print(f"\n{prefix}{title}")
    print(f"{prefix}{'Year':<8}" + "".join(f"{y:>12}" for y in years))
    print(f"{prefix}{'Value':<8}" + "".join(fmt(data.get(y)) for y in years))


def main():
    # Default: try Excel first, fall back to CSV
    filepath = None
    candidates = [
        "/mnt/user-data/uploads/YPF_DCF__1_.xlsx",
        "model_data.csv",
        "../model_data.csv",
    ]
    for c in candidates:
        if os.path.exists(c):
            filepath = c
            break

    if filepath is None:
        print("ERROR: No data file found. Provide a path as argument.")
        print("Usage: python demo.py [path/to/YPF_DCF.xlsx]")
        sys.exit(1)

    if len(sys.argv) > 1:
        filepath = sys.argv[1]

    print(f"Loading model from: {filepath}")
    model = YPFModel(filepath)
    print(model)
    print(f"Schedules: {[s.SCHEDULE_NAME for s in model.all_schedules]}")

    # ── Income Statement highlights ──
    print("\n" + "=" * 80)
    print("INCOME STATEMENT HIGHLIGHTS")
    print("=" * 80)
    is_items = model.income_statement.line_items
    for key in ["revenue", "cost_of_sales", "gross_profit", "ebitda", "ebit", "net_income"]:
        print_table(key.replace("_", " ").title(), is_items[key])

    # ── Margins ──
    print("\n" + "=" * 80)
    print("MARGINS")
    print("=" * 80)
    margins = model.income_statement.margins
    for key in ["gross_margin", "ebitda_margin", "ebit_margin"]:
        data = margins[key]
        title = key.replace("_", " ").title()
        years = sorted(data.keys())
        print(f"\n  {title}")
        print(f"  {'Year':<8}" + "".join(f"{y:>12}" for y in years))
        print(f"  {'Value':<8}" + "".join(f"{data[y]:>12.1%}" for y in years))

    # ── Balance Sheet snapshot ──
    print("\n" + "=" * 80)
    print("BALANCE SHEET TOTALS")
    print("=" * 80)
    print_table("Total Assets", model.balance_sheet.total_assets)
    print_table("Total Equity", model.balance_sheet.shareholders_equity["total"])
    print_table("Check (should be 0)", model.balance_sheet.check)

    # ── Cash Flow highlights ──
    print("\n" + "=" * 80)
    print("CASH FLOW HIGHLIGHTS")
    print("=" * 80)
    print_table("Operating CF", model.cash_flow.operating["total"])
    print_table("Investing CF", model.cash_flow.investing["total"])
    print_table("Ending Cash", model.cash_flow.cash_position["ending"])

    # ── Oil Revenue ──
    print("\n" + "=" * 80)
    print("OIL REVENUE SCHEDULE")
    print("=" * 80)
    print_table("Oil Price ($/Boe)", model.oil_revenue.pricing["oil_and_consolidates"])
    print_table("Total Volume (MMBoe)", model.oil_revenue.volumes["total"])
    print_table("Total Revenue ($)", model.oil_revenue.revenue["total"])

    # ── Debt ──
    print("\n" + "=" * 80)
    print("DEBT & INTEREST")
    print("=" * 80)
    print_table("Total Loans & Revolver", model.debt_and_interest.totals["total_loans_revolver"])
    print_table("Total Interest Expense", model.debt_and_interest.totals["total_interest_expense"])
    print_table("Cash Interest Income", model.debt_and_interest.cash["annual_interest_income"])

    # ── Working Capital ──
    print("\n" + "=" * 80)
    print("WORKING CAPITAL")
    print("=" * 80)
    print_table("Net Working Capital", model.working_capital.net_working_capital)
    print_table("Change in WC", model.working_capital.change_in_working_capital)

    print("\n\nDone. All 14 schedules loaded successfully.")


if __name__ == "__main__":
    main()
