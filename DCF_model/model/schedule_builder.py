"""
Schedule builder orchestrator - builds all schedules based on company industry.

Loads CapIQ data and instantiates appropriate schedules based on the company's
industry type. Basic schedules are always created; industry-specific schedules
are optional.

Usage:
    from model.schedule_builder import ScheduleBuilder
    model = ScheduleBuilder("Duolingo Inc NasdaqGS DUOL Financials.xls")
    print(model.industry)
    print(model.all_schedules)
"""

from loaders.capiq_loader import CapIQLoader
from schedules import (
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
from .company_types import get_industry_type, IndustryType


class ScheduleBuilder:
    """
    Master model object that builds all schedules based on company industry.

    Loads data once via CapIQLoader and exposes each schedule as a named attribute.
    Industry-specific schedules are instantiated conditionally.
    """

    def __init__(self, filepath: str):
        self.loader = CapIQLoader(filepath)
        self.industry = get_industry_type(self.loader.ticker)

        # ── Always create: Basic financial schedules ──
        self.income_statement = IncomeStatement(self.loader)
        self.cash_flow = CashFlowStatement(self.loader)
        self.balance_sheet = BalanceSheet(self.loader)
        self.fixed_assets = FixedAssetsSchedule(self.loader)
        self.working_capital = WorkingCapitalSchedule(self.loader)
        self.debt_and_interest = DebtAndInterestSchedule(self.loader)
        self.shareholders_equity = ShareholdersEquitySchedule(self.loader)

        # ── Always create: General revenue schedules ──
        self.total_revenue = TotalRevenueSchedule(self.loader)

        # ── Conditionally create: Oil & Gas specific schedules ──
        self.oil_revenue = None
        self.crude_products_revenue = None
        self.other_products_revenue = None
        self.downstream_revenue = None
        self.production_costs = None

        if self.industry == IndustryType.OIL_AND_GAS:
            self.oil_revenue = OilRevenueSchedule(self.loader)
            self.crude_products_revenue = CrudeProductsRevenueSchedule(self.loader)
            self.other_products_revenue = OtherProductsRevenueSchedule(self.loader)
            self.downstream_revenue = DownstreamRevenueSchedule(self.loader)
            self.production_costs = ProductionCostsSchedule(self.loader)

    @property
    def all_schedules(self) -> list:
        """Return a list of all non-None schedules."""
        return [s for s in [
            self.oil_revenue,
            self.crude_products_revenue,
            self.other_products_revenue,
            self.downstream_revenue,
            self.total_revenue,
            self.production_costs,
            self.income_statement,
            self.cash_flow,
            self.balance_sheet,
            self.fixed_assets,
            self.working_capital,
            self.debt_and_interest,
            self.shareholders_equity,
        ] if s is not None]

    def summary(self) -> dict:
        """Return the full model as a nested dict (every schedule's summary)."""
        return {s.SCHEDULE_NAME: s.summary() for s in self.all_schedules}

    def __repr__(self):
        industry_name = {
            IndustryType.GENERIC: "Generic",
            IndustryType.OIL_AND_GAS: "Oil & Gas",
        }.get(self.industry, "Unknown")
        return f"<ScheduleBuilder ({industry_name}): {len(self.all_schedules)} schedules>"
