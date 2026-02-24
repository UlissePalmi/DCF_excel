"""
YPF DCF Model – Top-level orchestrator.

Instantiate with a path to the Model sheet data (CSV or Excel).
Provides access to every schedule as a named attribute.

Usage:
    from ypf_model import YPFModel
    model = YPFModel("YPF_DCF.xlsx")

    # Access any schedule
    print(model.income_statement.line_items["ebitda"])
    print(model.balance_sheet.current_assets["cash"])

    # Get a full summary dict of every schedule
    full = model.summary()
"""

from data_loader import DataLoader
from schedules import (
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


class YPFModel:
    """
    Master model object for the YPF DCF.
    
    Loads data once via DataLoader and exposes each schedule as a property.
    """

    def __init__(self, filepath: str):
        self.loader = DataLoader(filepath)

        # ── Revenue schedules ──
        self.oil_revenue = OilRevenueSchedule(self.loader)
        self.crude_products_revenue = CrudeProductsRevenueSchedule(self.loader)
        self.other_products_revenue = OtherProductsRevenueSchedule(self.loader)
        self.downstream_revenue = DownstreamRevenueSchedule(self.loader)
        self.total_revenue = TotalRevenueSchedule(self.loader)

        # ── Cost schedules ──
        self.production_costs = ProductionCostsSchedule(self.loader)
        self.selling_and_admin = SellingAndAdminExpensesSchedule(self.loader)

        # ── Financial statements ──
        self.income_statement = IncomeStatement(self.loader)
        self.cash_flow = CashFlowStatement(self.loader)
        self.balance_sheet = BalanceSheet(self.loader)

        # ── Supporting schedules ──
        self.fixed_assets = FixedAssetsSchedule(self.loader)
        self.working_capital = WorkingCapitalSchedule(self.loader)
        self.debt_and_interest = DebtAndInterestSchedule(self.loader)
        self.shareholders_equity = ShareholdersEquitySchedule(self.loader)

    # Convenience: list all schedules
    @property
    def all_schedules(self) -> list:
        return [
            self.oil_revenue,
            self.crude_products_revenue,
            self.other_products_revenue,
            self.downstream_revenue,
            self.total_revenue,
            self.production_costs,
            self.selling_and_admin,
            self.income_statement,
            self.cash_flow,
            self.balance_sheet,
            self.fixed_assets,
            self.working_capital,
            self.debt_and_interest,
            self.shareholders_equity,
        ]

    def summary(self) -> dict:
        """Return the full model as a nested dict (every schedule's summary)."""
        return {s.SCHEDULE_NAME: s.summary() for s in self.all_schedules}

    def __repr__(self):
        return f"<YPFModel: {len(self.all_schedules)} schedules loaded>"
