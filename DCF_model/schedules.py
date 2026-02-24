"""
Individual schedule classes for the YPF DCF Model.

Each class reads from its own sheet in the multi-sheet Excel workbook via
self._field(key), where key matches a row label in that sheet.
"""

from base_schedule import BaseSchedule


# ─────────────────────────────────────────────────────────
# 1. Oil Revenue Schedule
# ─────────────────────────────────────────────────────────
class OilRevenueSchedule(BaseSchedule):
    SCHEDULE_NAME = "Oil Revenue Schedule"
    SHEET_NAME    = "Oil Revenue Schedule"

    @property
    def pricing(self) -> dict:
        return {
            "oil_and_consolidates": self._field("price_oil_and_consolidates"),
            "ngl":                  self._field("price_ngl"),
            "natural_gas":          self._field("price_natural_gas"),
        }

    @property
    def volumes(self) -> dict:
        return {
            "oil_and_consolidates": self._field("vol_oil_and_consolidates"),
            "ngl":                  self._field("vol_ngl"),
            "natural_gas":          self._field("vol_natural_gas"),
            "total":                self._field("vol_total"),
        }

    @property
    def revenue(self) -> dict:
        return {
            "oil_and_consolidates": self._field("rev_oil_and_consolidates"),
            "ngl":                  self._field("rev_ngl"),
            "natural_gas":          self._field("rev_natural_gas"),
            "total":                self._field("rev_total"),
        }

    @property
    def purchases(self) -> dict[int, float]:
        return self._field("purchases")

    def summary(self) -> dict:
        return {
            "pricing":   self.pricing,
            "volumes":   self.volumes,
            "revenue":   self.revenue,
            "purchases": self.purchases,
        }


# ─────────────────────────────────────────────────────────
# 2. Crude Products Revenue Schedule
# ─────────────────────────────────────────────────────────
class CrudeProductsRevenueSchedule(BaseSchedule):
    SCHEDULE_NAME = "Crude Products Revenue Schedule"
    SHEET_NAME    = "Crude Products Revenue Schedule"

    def _product_block(self, prefix: str, has_actual: bool = True) -> dict:
        d = {
            "price":  {"domestic": self._field(f"price_{prefix}_domestic"),
                       "export":   self._field(f"price_{prefix}_export")},
            "volume": {"domestic": self._field(f"vol_{prefix}_domestic"),
                       "export":   self._field(f"vol_{prefix}_export"),
                       "total":    self._field(f"vol_{prefix}_total")},
            "revenue":{"domestic": self._field(f"rev_{prefix}_domestic"),
                       "export":   self._field(f"rev_{prefix}_export"),
                       "total":    self._field(f"rev_{prefix}_total")},
        }
        if has_actual:
            d["revenue"]["actual_total"] = self._field(f"rev_{prefix}_actual")
        return d

    @property
    def diesel(self):    return self._product_block("diesel")
    @property
    def gasolines(self): return self._product_block("gasoline")
    @property
    def jet_fuel(self):  return self._product_block("jet")
    @property
    def fuel_oil(self):  return self._product_block("fueloil")

    @property
    def total_revenue(self) -> dict[int, float]:
        return self._field("rev_total")

    def summary(self) -> dict:
        return {
            "diesel":        self.diesel,
            "gasolines":     self.gasolines,
            "jet_fuel":      self.jet_fuel,
            "fuel_oil":      self.fuel_oil,
            "total_revenue": self.total_revenue,
        }


# ─────────────────────────────────────────────────────────
# 3. Other Products Revenue Schedule
# ─────────────────────────────────────────────────────────
class OtherProductsRevenueSchedule(BaseSchedule):
    SCHEDULE_NAME = "Other Products Revenue Schedule"
    SHEET_NAME    = "Other Products Revenue Schedule"

    def _product_block(self, prefix: str) -> dict:
        return {
            "price":  {"domestic": self._field(f"price_{prefix}_domestic"),
                       "export":   self._field(f"price_{prefix}_export")},
            "volume": {"domestic": self._field(f"vol_{prefix}_domestic"),
                       "export":   self._field(f"vol_{prefix}_export"),
                       "total":    self._field(f"vol_{prefix}_total")},
            "revenue":{"domestic":    self._field(f"rev_{prefix}_domestic"),
                       "export":      self._field(f"rev_{prefix}_export"),
                       "total":       self._field(f"rev_{prefix}_total"),
                       "actual_total":self._field(f"rev_{prefix}_actual")},
        }

    @property
    def virgin_naphtha(self):   return self._product_block("naphtha")
    @property
    def petrochemicals(self):   return self._product_block("petrochem")

    @property
    def fertilizers(self):
        d = self._product_block("fert")
        d["revenue"]["crop_protection"] = self._field("rev_crop_protection")
        return d

    @property
    def total_revenue(self): return self._field("rev_total")

    def summary(self) -> dict:
        return {
            "virgin_naphtha":  self.virgin_naphtha,
            "petrochemicals":  self.petrochemicals,
            "fertilizers":     self.fertilizers,
            "total_revenue":   self.total_revenue,
        }


# ─────────────────────────────────────────────────────────
# 4. Downstream Revenue Schedule
# ─────────────────────────────────────────────────────────
class DownstreamRevenueSchedule(BaseSchedule):
    SCHEDULE_NAME = "Downstream Revenue Schedule"
    SHEET_NAME    = "Downstream Revenue Schedule"

    @property
    def prices(self):
        return {
            "base_oils":         self._field("price_base_oils"),
            "coke":              self._field("price_coke"),
            "lpg":               self._field("price_lpg"),
            "asphalt":           self._field("price_asphalt"),
            "fuel_oil":          self._field("price_fuel_oil"),
            "diesel":            self._field("price_diesel"),
            "gasolines":         self._field("price_gasolines"),
            "petrochem_naphtha": self._field("price_petrochem_naphtha"),
            "jet_fuel":          self._field("price_jet_fuel"),
        }

    @property
    def volumes(self):
        return {
            "base_oils":         self._field("vol_base_oils"),
            "coke":              self._field("vol_coke"),
            "lpg":               self._field("vol_lpg"),
            "asphalt":           self._field("vol_asphalt"),
            "fuel_oil":          self._field("vol_fuel_oil"),
            "diesel":            self._field("vol_diesel"),
            "gasolines":         self._field("vol_gasolines"),
            "petrochem_naphtha": self._field("vol_petrochem_naphtha"),
            "jet_fuel":          self._field("vol_jet_fuel"),
        }

    @property
    def revenue(self):
        return {
            "lubricants_byproducts": self._field("rev_lubricants_byproducts"),
            "petroleum_coke":        self._field("rev_petroleum_coke"),
            "lpg":                   self._field("rev_lpg"),
            "asphalts":              self._field("rev_asphalts"),
            "subtotal_1":            self._field("rev_subtotal_1"),
            "fuel_oil":              self._field("rev_fuel_oil"),
            "diesel":                self._field("rev_diesel"),
            "gasolines":             self._field("rev_gasolines"),
            "petrochem_naphtha":     self._field("rev_petrochem_naphtha"),
            "jet_fuel":              self._field("rev_jet_fuel"),
            "subtotal_2":            self._field("rev_subtotal_2"),
        }

    def summary(self) -> dict:
        return {"prices": self.prices, "volumes": self.volumes, "revenue": self.revenue}


# ─────────────────────────────────────────────────────────
# 5. Total Revenue Schedule
# ─────────────────────────────────────────────────────────
class TotalRevenueSchedule(BaseSchedule):
    SCHEDULE_NAME = "Total Revenue Schedule"
    SHEET_NAME    = "Total Revenue Schedule"

    @property
    def natural_gas(self):
        return {
            "price":  {"domestic": self._field("price_ng_domestic"),
                       "export":   self._field("price_ng_export")},
            "volume": {"domestic": self._field("vol_ng_domestic"),
                       "export":   self._field("vol_ng_export"),
                       "total":    self._field("vol_ng_total")},
            "revenue":{"domestic": self._field("rev_ng_domestic"),
                       "export":   self._field("rev_ng_export"),
                       "total":    self._field("rev_ng_total")},
        }

    @property
    def crude_oil(self):
        return {
            "price":  {"domestic": self._field("price_crude_domestic"),
                       "export":   self._field("price_crude_export")},
            "volume": {"domestic": self._field("vol_crude_domestic"),
                       "export":   self._field("vol_crude_export"),
                       "total":    self._field("vol_crude_total")},
            "revenue":{"domestic": self._field("rev_crude_domestic"),
                       "export":   self._field("rev_crude_export"),
                       "total":    self._field("rev_crude_total")},
        }

    @property
    def argentina_gdp(self): return self._field("argentina_gdp")

    @property
    def revenue_components(self):
        return {
            "main_crude_products": self._field("rev_component_main_crude_products"),
            "other_products":      self._field("rev_component_other_products"),
            "downstream":          self._field("rev_component_downstream"),
        }

    @property
    def other_revenue(self):
        return {
            "gas_stations":        self._field("other_rev_gas_stations"),
            "construction_contracts": self._field("other_rev_construction_contracts"),
            "lng_regasification":  self._field("other_rev_lng_regasification"),
            "other_goods_services":self._field("other_rev_other_goods_services"),
            "subtotal":            self._field("other_rev_subtotal"),
        }

    @property
    def total_revenue(self): return self._field("total_revenue")

    def summary(self) -> dict:
        return {
            "natural_gas":        self.natural_gas,
            "crude_oil":          self.crude_oil,
            "argentina_gdp":      self.argentina_gdp,
            "revenue_components": self.revenue_components,
            "other_revenue":      self.other_revenue,
            "total_revenue":      self.total_revenue,
        }


# ─────────────────────────────────────────────────────────
# 6. Production Costs Expenses Schedule
# ─────────────────────────────────────────────────────────
class ProductionCostsSchedule(BaseSchedule):
    SCHEDULE_NAME = "Production Costs Expenses Schedule"
    SHEET_NAME    = "Production Costs Expenses Schedule"

    @property
    def macro(self):
        return {
            "oil_prices":          self._field("macro_oil_prices"),
            "volumes_growth":      self._field("macro_volumes_growth"),
            "argentina_inflation": self._field("macro_argentina_inflation"),
            "usa_inflation":       self._field("macro_usa_inflation"),
            "fx_rate":             self._field("macro_fx_rate"),
            "depreciation_rate":   self._field("macro_depreciation_rate"),
        }

    @property
    def royalties_and_fees(self):
        return {
            "royalties_easements": self._field("royalties_easements"),
            "fees_compensation":   self._field("fees_compensation"),
            "total":               self._field("royalties_total"),
            "pct_revenue":         self._field("royalties_pct_revenue"),
        }

    @property
    def arg_inflation_linked(self):
        return {
            "salaries":          self._field("arg_salaries"),
            "other_personnel":   self._field("arg_other_personnel"),
            "rental":            self._field("arg_rental"),
            "transportation":    self._field("arg_transportation"),
            "preservation_repair":self._field("arg_preservation_repair"),
            "operation_services":self._field("arg_operation_services"),
            "taxes_charges":     self._field("arg_taxes_charges"),
        }

    @property
    def usa_inflation_linked(self):
        return {
            "industrial_inputs": self._field("usa_industrial_inputs"),
            "insurance":         self._field("usa_insurance"),
        }

    @property
    def oil_price_linked(self):
        return {"fuel_gas_energy": self._field("oil_fuel_gas_energy")}

    @property
    def total_production_costs(self): return self._field("total_production_costs")

    @property
    def cost_of_sale(self):
        return {
            "inventories_beginning": self._field("cos_inventories_beginning"),
            "purchases":             self._field("cos_purchases"),
            "production_costs":      self._field("cos_production_costs"),
            "currency_conversions":  self._field("cos_currency_conversions"),
            "inventories_ending":    self._field("cos_inventories_ending"),
            "total":                 self._field("cos_total"),
        }

    def summary(self) -> dict:
        return {
            "macro":                      self.macro,
            "royalties_and_fees":         self.royalties_and_fees,
            "arg_inflation_linked_costs": self.arg_inflation_linked,
            "usa_inflation_linked_costs": self.usa_inflation_linked,
            "oil_price_linked_costs":     self.oil_price_linked,
            "total_production_costs":     self.total_production_costs,
            "cost_of_sale":               self.cost_of_sale,
        }


# ─────────────────────────────────────────────────────────
# 7. S&A Expenses Schedule
# ─────────────────────────────────────────────────────────
class SellingAndAdminExpensesSchedule(BaseSchedule):
    SCHEDULE_NAME = "S&A Expenses Schedule"
    SHEET_NAME    = "S&A Expenses Schedule"

    @property
    def selling_expenses(self):
        return {
            "salaries":            self._field("sell_salaries"),
            "fees":                self._field("sell_fees"),
            "other_personnel":     self._field("sell_other_personnel"),
            "taxes":               self._field("sell_taxes"),
            "royalties":           self._field("sell_royalties"),
            "insurance":           self._field("sell_insurance"),
            "rental":              self._field("sell_rental"),
            "industrial_inputs":   self._field("sell_industrial_inputs"),
            "operation_services":  self._field("sell_operation_services"),
            "preservation":        self._field("sell_preservation"),
            "transportation":      self._field("sell_transportation"),
            "publicity":           self._field("sell_publicity"),
            "doubtful_receivables":self._field("sell_doubtful_receivables"),
            "fuel_gas_energy":     self._field("sell_fuel_gas_energy"),
            "total":               self._field("sell_total"),
        }

    @property
    def admin_expenses(self):
        return {
            "salaries":           self._field("admin_salaries"),
            "fees":               self._field("admin_fees"),
            "other_personnel":    self._field("admin_other_personnel"),
            "taxes":              self._field("admin_taxes"),
            "operation_services": self._field("admin_operation_services"),
            "preservation":       self._field("admin_preservation"),
            "publicity":          self._field("admin_publicity"),
            "other":              self._field("admin_other"),
            "fuel_gas_energy":    self._field("admin_fuel_gas_energy"),
            "total":              self._field("admin_total"),
        }

    @property
    def exploration_expenses(self): return self._field("exploration_expenses")

    def summary(self) -> dict:
        return {
            "selling_expenses":    self.selling_expenses,
            "admin_expenses":      self.admin_expenses,
            "exploration_expenses":self.exploration_expenses,
        }


# ─────────────────────────────────────────────────────────
# 8. Income Statement
# ─────────────────────────────────────────────────────────
class IncomeStatement(BaseSchedule):
    SCHEDULE_NAME = "Income Statement"
    SHEET_NAME    = "Income Statement"

    @property
    def line_items(self):
        return {
            "revenue":              self._field("revenue"),
            "cost_of_sales":        self._field("cost_of_sales"),
            "gross_profit":         self._field("gross_profit"),
            "selling_expenses":     self._field("selling_expenses"),
            "admin_expenses":       self._field("admin_expenses"),
            "exploration_expenses": self._field("exploration_expenses"),
            "other":                self._field("other"),
            "operating_costs":      self._field("operating_costs"),
            "impairment":           self._field("impairment"),
            "ebitda":               self._field("ebitda"),
            "da":                   self._field("da"),
            "ebit":                 self._field("ebit"),
            "equity_income":        self._field("equity_income"),
            "financial_income":     self._field("financial_income"),
            "financial_costs":      self._field("financial_costs"),
            "other_financial":      self._field("other_financial"),
            "ebt":                  self._field("ebt"),
            "income_tax":           self._field("income_tax"),
            "net_income":           self._field("net_income"),
            "nopat":                self._field("nopat"),
        }

    @property
    def margins(self):
        return {
            "revenue_growth": self._field("revenue_growth"),
            "cogs_growth":    self._field("cogs_growth"),
            "gross_margin":   self._field("gross_margin"),
            "ebitda_margin":  self._field("ebitda_margin"),
            "ebit_margin":    self._field("ebit_margin"),
            "roe":            self._field("roe"),
        }

    def summary(self) -> dict:
        return {"line_items": self.line_items, "margins": self.margins}


# ─────────────────────────────────────────────────────────
# 9. Cash Flow Statement
# ─────────────────────────────────────────────────────────
class CashFlowStatement(BaseSchedule):
    SCHEDULE_NAME = "Cash Flow Statement"
    SHEET_NAME    = "Cash Flow Statement"

    @property
    def operating(self):
        return {
            "net_income":         self._field("net_income"),
            "equity_interests":   self._field("equity_interests"),
            "depreciation_ppe":   self._field("depreciation_ppe"),
            "amortization_ia":    self._field("amortization_ia"),
            "depreciation_rou":   self._field("depreciation_rou"),
            "retirement_ppe":     self._field("retirement_ppe"),
            "impairment":         self._field("impairment"),
            "income_tax_charge":  self._field("income_tax_charge"),
            "provisions":         self._field("provisions"),
            "fx_interest_other":  self._field("fx_interest_other"),
            "working_capital":    self._field("working_capital"),
            "total":              self._field("cf_operating_total"),
        }

    @property
    def investing(self):
        return {
            "capex":               self._field("capex"),
            "assets_held_for_sale":self._field("assets_held_for_sale"),
            "acquisitions_jv":     self._field("acquisitions_jv"),
            "total":               self._field("cf_investing_total"),
        }

    @property
    def financing(self):
        return {
            "loan_payments":    self._field("loan_payments"),
            "loan_proceeds":    self._field("loan_proceeds"),
            "interest_payments":self._field("interest_payments"),
            "overdraft":        self._field("overdraft"),
            "buyback":          self._field("buyback"),
            "lease_payments":   self._field("lease_payments"),
            "total":            self._field("cf_financing_total"),
        }

    @property
    def cash_position(self):
        return {
            "change":    self._field("change_in_cash"),
            "beginning": self._field("beginning_cash"),
            "ending":    self._field("ending_cash"),
        }

    def summary(self) -> dict:
        return {
            "operating":     self.operating,
            "investing":     self.investing,
            "financing":     self.financing,
            "cash_position": self.cash_position,
        }


# ─────────────────────────────────────────────────────────
# 10. Balance Sheet
# ─────────────────────────────────────────────────────────
class BalanceSheet(BaseSchedule):
    SCHEDULE_NAME = "Balance Sheet"
    SHEET_NAME    = "Balance Sheet"

    @property
    def current_assets(self):
        return {
            "cash":                self._field("ca_cash"),
            "investments":         self._field("ca_investments"),
            "trade_receivables":   self._field("ca_trade_receivables"),
            "contract_asset":      self._field("ca_contract_asset"),
            "other_receivables":   self._field("ca_other_receivables"),
            "inventories":         self._field("ca_inventories"),
            "assets_held_for_sale":self._field("ca_assets_held_for_sale"),
            "total":               self._field("ca_total"),
        }

    @property
    def non_current_assets(self):
        return {
            "financial_investments":self._field("nca_financial_investments"),
            "trade_receivables":     self._field("nca_trade_receivables"),
            "other_receivables":     self._field("nca_other_receivables"),
            "deferred_tax_asset":    self._field("nca_deferred_tax_asset"),
            "associates_jv":         self._field("nca_associates_jv"),
            "rou_assets":            self._field("nca_rou_assets"),
            "ppe":                   self._field("nca_ppe"),
            "intangible":            self._field("nca_intangible"),
        }

    @property
    def current_liabilities(self):
        return {
            "accounts_payable":    self._field("cl_accounts_payable"),
            "other_liabilities":   self._field("cl_other_liabilities"),
            "loans":               self._field("cl_loans"),
            "lease_liabilities":   self._field("cl_lease_liabilities"),
            "salaries_ss":         self._field("cl_salaries_ss"),
            "taxes_payable":       self._field("cl_taxes_payable"),
            "income_tax_payable":  self._field("cl_income_tax_payable"),
            "contract_liabilities":self._field("cl_contract_liabilities"),
            "provisions":          self._field("cl_provisions"),
            "liab_held_for_sale":  self._field("cl_liab_held_for_sale"),
            "total":               self._field("cl_total"),
        }

    @property
    def non_current_liabilities(self):
        return {
            "accounts_payable":        self._field("ncl_accounts_payable"),
            "other_liabilities":       self._field("ncl_other_liabilities"),
            "loans":                   self._field("ncl_loans"),
            "lease_liabilities":       self._field("ncl_lease_liabilities"),
            "salaries_ss":             self._field("ncl_salaries_ss"),
            "taxes_payable":           self._field("ncl_taxes_payable"),
            "income_tax_payable":      self._field("ncl_income_tax_payable"),
            "deferred_tax_liabilities":self._field("ncl_deferred_tax_liabilities"),
            "contract_liabilities":    self._field("ncl_contract_liabilities"),
            "provisions":              self._field("ncl_provisions"),
            "total":                   self._field("ncl_total"),
        }

    @property
    def shareholders_equity(self):
        return {
            "common_stock":      self._field("eq_common_stock"),
            "retained_earnings": self._field("eq_retained_earnings"),
            "minority_interest": self._field("eq_minority_interest"),
            "total":             self._field("eq_total"),
        }

    @property
    def total_assets(self): return self._field("total_assets")

    @property
    def total_liabilities_and_equity(self): return self._field("total_liabilities_and_equity")

    @property
    def check(self): return self._field("check")

    def summary(self) -> dict:
        return {
            "current_assets":              self.current_assets,
            "non_current_assets":          self.non_current_assets,
            "total_assets":                self.total_assets,
            "current_liabilities":         self.current_liabilities,
            "non_current_liabilities":     self.non_current_liabilities,
            "shareholders_equity":         self.shareholders_equity,
            "total_liabilities_and_equity":self.total_liabilities_and_equity,
            "check":                       self.check,
        }


# ─────────────────────────────────────────────────────────
# 11. Fixed Assets (PP&E) Schedule
# ─────────────────────────────────────────────────────────
class FixedAssetsSchedule(BaseSchedule):
    SCHEDULE_NAME = "Fixed Assets (PP&E) Schedule"
    SHEET_NAME    = "Fixed Assets (PP&E) Schedule"

    @property
    def ppe(self):
        return {
            "beginning":       self._field("ppe_beginning"),
            "capex":           self._field("ppe_capex"),
            "new_intangibles": self._field("ppe_new_intangibles"),
            "depreciation": {
                "production_costs": self._field("ppe_depr_production_costs"),
                "selling":          self._field("ppe_depr_selling"),
                "admin":            self._field("ppe_depr_admin"),
                "total":            self._field("ppe_depr_total"),
            },
            "depreciation_pct": {
                "prod_to_capex":  self._field("ppe_pct_prod_to_capex"),
                "sell_to_capex":  self._field("ppe_pct_sell_to_capex"),
                "admin_to_capex": self._field("ppe_pct_admin_to_capex"),
            },
            "amortization": {
                "production_costs": self._field("ppe_amort_production_costs"),
                "selling":          self._field("ppe_amort_selling"),
                "admin":            self._field("ppe_amort_admin"),
                "total":            self._field("ppe_amort_total"),
            },
            "impairment": self._field("ppe_impairment"),
            "ending":     self._field("ppe_ending"),
        }

    @property
    def rou_assets(self):
        return {
            "beginning": self._field("rou_beginning"),
            "additions": self._field("rou_additions"),
            "depreciation": {
                "production_costs": self._field("rou_depr_production_costs"),
                "selling":          self._field("rou_depr_selling"),
                "total":            self._field("rou_depr_total"),
            },
            "ending": self._field("rou_ending"),
        }

    @property
    def total_da(self): return self._field("total_da")

    def summary(self) -> dict:
        return {"ppe": self.ppe, "rou_assets": self.rou_assets, "total_da": self.total_da}


# ─────────────────────────────────────────────────────────
# 12. Working Capital Schedule
# ─────────────────────────────────────────────────────────
class WorkingCapitalSchedule(BaseSchedule):
    SCHEDULE_NAME = "Working Capital Schedule"
    SHEET_NAME    = "Working Capital Schedule"

    @property
    def days_in(self):
        return {
            "current": {
                "trade_receivables":   self._field("days_c_trade_receivables"),
                "contract_asset":      self._field("days_c_contract_asset"),
                "other_receivables":   self._field("days_c_other_receivables"),
                "inventories":         self._field("days_c_inventories"),
                "accounts_payable":    self._field("days_c_accounts_payable"),
                "other_liabilities":   self._field("days_c_other_liabilities"),
                "lease_liabilities":   self._field("days_c_lease_liabilities"),
                "salaries":            self._field("days_c_salaries"),
                "taxes":               self._field("days_c_taxes"),
                "income_tax":          self._field("days_c_income_tax"),
                "contract_liabilities":self._field("days_c_contract_liabilities"),
                "provisions":          self._field("days_c_provisions"),
            },
            "non_current": {
                "trade_receivables":   self._field("days_nc_trade_receivables"),
                "other_receivables":   self._field("days_nc_other_receivables"),
                "accounts_payable":    self._field("days_nc_accounts_payable"),
                "other_liabilities":   self._field("days_nc_other_liabilities"),
                "lease_liabilities":   self._field("days_nc_lease_liabilities"),
                "salaries":            self._field("days_nc_salaries"),
                "taxes":               self._field("days_nc_taxes"),
                "income_tax":          self._field("days_nc_income_tax"),
                "contract_liabilities":self._field("days_nc_contract_liabilities"),
                "provisions":          self._field("days_nc_provisions"),
            },
        }

    @property
    def net_working_capital(self): return self._field("net_working_capital")

    @property
    def change_in_working_capital(self): return self._field("change_in_working_capital")

    def summary(self) -> dict:
        return {
            "days_in":                   self.days_in,
            "net_working_capital":       self.net_working_capital,
            "change_in_working_capital": self.change_in_working_capital,
        }


# ─────────────────────────────────────────────────────────
# 13. Debt and Interest Schedule
# ─────────────────────────────────────────────────────────
class DebtAndInterestSchedule(BaseSchedule):
    SCHEDULE_NAME = "Debt and Interest Schedule"
    SHEET_NAME    = "Debt and Interest Schedule"

    @property
    def cash(self):
        return {
            "beginning":             self._field("cash_beginning"),
            "change":                self._field("cash_change"),
            "ending":                self._field("cash_ending"),
            "interest_rate":         self._field("cash_interest_rate"),
            "interest_income":       self._field("cash_interest_income"),
            "annual_interest_income":self._field("cash_annual_interest_income"),
        }

    @property
    def loans(self):
        return {
            "beginning":            self._field("loans_beginning"),
            "additions_repayments": self._field("loans_additions_repayments"),
            "ending":               self._field("loans_ending"),
            "interest_rate":        self._field("loans_interest_rate"),
            "interest_expense":     self._field("loans_interest_expense"),
        }

    @property
    def revolver(self):
        return {
            "operating_cf":              self._field("revolver_operating_cf"),
            "investing_cf":              self._field("revolver_investing_cf"),
            "financing_cf_ex_revolver":  self._field("revolver_financing_cf_ex_revolver"),
            "fcf_after_debt":            self._field("revolver_fcf_after_debt"),
            "beginning":                 self._field("revolver_beginning"),
            "change":                    self._field("revolver_change"),
            "ending":                    self._field("revolver_ending"),
            "interest_rate":             self._field("revolver_interest_rate"),
            "interest_expense":          self._field("revolver_interest_expense"),
        }

    @property
    def totals(self):
        return {
            "st_loans_revolver":    self._field("totals_st_loans_revolver"),
            "lt_loans_revolver":    self._field("totals_lt_loans_revolver"),
            "total_loans_revolver": self._field("totals_total_loans_revolver"),
            "total_interest_expense":self._field("totals_total_interest_expense"),
        }

    def summary(self) -> dict:
        return {
            "cash": self.cash, "loans": self.loans,
            "revolver": self.revolver, "totals": self.totals,
        }


# ─────────────────────────────────────────────────────────
# 14. Shareholders' Equity Schedule
# ─────────────────────────────────────────────────────────
class ShareholdersEquitySchedule(BaseSchedule):
    SCHEDULE_NAME = "Shareholders' Equity Schedule"
    SHEET_NAME    = "Shareholders' Equity Schedule"

    @property
    def common_shares(self):
        return {
            "beginning":        self._field("shares_beginning"),
            "class_a":          self._field("shares_class_a"),
            "class_b":          self._field("shares_class_b"),
            "class_c":          self._field("shares_class_c"),
            "class_d":          self._field("shares_class_d"),
            "total_outstanding":self._field("shares_total_outstanding"),
            "new_shares":       self._field("shares_new_shares"),
            "buybacks":         self._field("shares_buybacks"),
            "ending":           self._field("shares_ending"),
            "growth_yoy":       self._field("shares_growth_yoy"),
            "share_price":      self._field("share_price"),
        }

    @property
    def dividends(self):
        return {
            "payout_rate":    self._field("dividends_payout_rate"),
            "net_income":     self._field("dividends_net_income"),
            "common_dividend":self._field("dividends_common_dividend"),
        }

    @property
    def retained_earnings(self):
        return {
            "beginning": self._field("re_beginning"),
            "net_income":self._field("re_net_income"),
            "dividend":  self._field("re_dividend"),
            "ending":    self._field("re_ending"),
        }

    def summary(self) -> dict:
        return {
            "common_shares":     self.common_shares,
            "dividends":         self.dividends,
            "retained_earnings": self.retained_earnings,
        }
