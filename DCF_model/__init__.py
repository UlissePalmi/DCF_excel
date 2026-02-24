"""
YPF DCF Model Package

Usage:
    from ypf_model import YPFModel
    model = YPFModel("path/to/YPF_DCF.xlsx")   # or .csv
"""

from .ypf_model import YPFModel
from .data_loader import DataLoader
from .base_schedule import BaseSchedule
from .schedules import (
    OilRevenueSchedule,
    CrudeProductsRevenueSchedule,
    OtherProductsRevenueSchedule,
    DownstreamRevenueSchedule,
    TotalRevenueSchedule,
    ProductionCostsSchedule,
    SellingAndAdminExpensesSchedule,
    IncomeStatement,
    CashFlowStatement,
    BalanceSheet,
    FixedAssetsSchedule,
    WorkingCapitalSchedule,
    DebtAndInterestSchedule,
    ShareholdersEquitySchedule,
)

__all__ = [
    "YPFModel",
    "DataLoader",
    "BaseSchedule",
    "OilRevenueSchedule",
    "CrudeProductsRevenueSchedule",
    "OtherProductsRevenueSchedule",
    "DownstreamRevenueSchedule",
    "TotalRevenueSchedule",
    "ProductionCostsSchedule",
    "SellingAndAdminExpensesSchedule",
    "IncomeStatement",
    "CashFlowStatement",
    "BalanceSheet",
    "FixedAssetsSchedule",
    "WorkingCapitalSchedule",
    "DebtAndInterestSchedule",
    "ShareholdersEquitySchedule",
]
