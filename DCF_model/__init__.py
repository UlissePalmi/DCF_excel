"""
YPF DCF Model Package

Usage:
    from ypf_model import YPFModel
    model = YPFModel("Duolingo Inc NasdaqGS DUOL Financials.xls")
"""

from .ypf_model import YPFModel
from .loaders.capiq_loader import CapIQLoader
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
    "CapIQLoader",
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
