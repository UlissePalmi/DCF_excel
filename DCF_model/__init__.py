"""
DCF Model Package

Usage:
    from model.schedule_builder import ScheduleBuilder
    model = ScheduleBuilder("Duolingo Inc NasdaqGS DUOL Financials.xls")
"""

from .model import ScheduleBuilder, IndustryType, get_industry_type
from .loaders.capiq_loader import CapIQLoader
from .base_schedule import BaseSchedule
from .schedules import (
    OilRevenueSchedule,
    CrudeProductsRevenueSchedule,
    OtherProductsRevenueSchedule,
    DownstreamRevenueSchedule,
    TotalRevenueSchedule,
    ProductionCostsSchedule,
    IncomeStatement,
    CashFlowStatement,
    BalanceSheet,
    FixedAssetsSchedule,
    WorkingCapitalSchedule,
    DebtAndInterestSchedule,
    ShareholdersEquitySchedule,
)

__all__ = [
    "ScheduleBuilder",
    "IndustryType",
    "get_industry_type",
    "CapIQLoader",
    "BaseSchedule",
    "OilRevenueSchedule",
    "CrudeProductsRevenueSchedule",
    "OtherProductsRevenueSchedule",
    "DownstreamRevenueSchedule",
    "TotalRevenueSchedule",
    "ProductionCostsSchedule",
    "IncomeStatement",
    "CashFlowStatement",
    "BalanceSheet",
    "FixedAssetsSchedule",
    "WorkingCapitalSchedule",
    "DebtAndInterestSchedule",
    "ShareholdersEquitySchedule",
]
