"""
Individual schedule classes replicating each section of the YPF Model sheet.

Each class maps descriptive field names to Excel row numbers so values
are retrieved from the DataLoader (CSV/Excel source).
"""

from base_schedule import BaseSchedule


# ─────────────────────────────────────────────────────────
# 1. Oil Revenue Schedule  (Rows 1–30)
# ─────────────────────────────────────────────────────────
class OilRevenueSchedule(BaseSchedule):
    SCHEDULE_NAME = "Oil Revenue Schedule"
    START_ROW, END_ROW = 1, 30

    # --- Pricing ($/Boe) ---
    ROW_PRICE_OIL = 10
    ROW_PRICE_NGL = 11
    ROW_PRICE_GAS = 12

    # --- Volumes (MMBbl / MMBoe) ---
    ROW_VOL_OIL = 16
    ROW_VOL_NGL = 17
    ROW_VOL_GAS = 18
    ROW_VOL_TOTAL = 19

    # --- Revenue ($) ---
    ROW_REV_OIL = 23
    ROW_REV_NGL = 24
    ROW_REV_GAS = 25
    ROW_REV_TOTAL = 26

    # --- Purchases ---
    ROW_PURCHASES = 29

    @property
    def pricing(self) -> dict:
        return {
            "oil_and_consolidates": self._series(self.ROW_PRICE_OIL),
            "ngl": self._series(self.ROW_PRICE_NGL),
            "natural_gas": self._series(self.ROW_PRICE_GAS),
        }

    @property
    def volumes(self) -> dict:
        return {
            "oil_and_consolidates": self._series(self.ROW_VOL_OIL),
            "ngl": self._series(self.ROW_VOL_NGL),
            "natural_gas": self._series(self.ROW_VOL_GAS),
            "total": self._series(self.ROW_VOL_TOTAL),
        }

    @property
    def revenue(self) -> dict:
        return {
            "oil_and_consolidates": self._series(self.ROW_REV_OIL),
            "ngl": self._series(self.ROW_REV_NGL),
            "natural_gas": self._series(self.ROW_REV_GAS),
            "total": self._series(self.ROW_REV_TOTAL),
        }

    @property
    def purchases(self) -> dict[int, float]:
        return self._series(self.ROW_PURCHASES)

    def summary(self) -> dict:
        return {
            "pricing": self.pricing,
            "volumes": self.volumes,
            "revenue": self.revenue,
            "purchases": self.purchases,
        }


# ─────────────────────────────────────────────────────────
# 2. Crude Products Revenue Schedule  (Rows 31–105)
# ─────────────────────────────────────────────────────────
class CrudeProductsRevenueSchedule(BaseSchedule):
    SCHEDULE_NAME = "Crude Products Revenue Schedule"
    START_ROW, END_ROW = 31, 105

    # ---- Prices ----
    # Diesel
    ROW_PRICE_DIESEL_DOM = 41
    ROW_PRICE_DIESEL_EXP = 42
    # Gasolines
    ROW_PRICE_GASOLINE_DOM = 45
    ROW_PRICE_GASOLINE_EXP = 46
    # Jet Fuel
    ROW_PRICE_JET_DOM = 49
    ROW_PRICE_JET_EXP = 50
    # Fuel Oil
    ROW_PRICE_FUELOIL_DOM = 53
    ROW_PRICE_FUELOIL_EXP = 54

    # ---- Volumes ----
    # Diesel
    ROW_VOL_DIESEL_DOM = 59
    ROW_VOL_DIESEL_EXP = 60
    ROW_VOL_DIESEL_TOT = 61
    # Gasolines
    ROW_VOL_GASOLINE_DOM = 64
    ROW_VOL_GASOLINE_EXP = 65
    ROW_VOL_GASOLINE_TOT = 66
    # Jet Fuel
    ROW_VOL_JET_DOM = 69
    ROW_VOL_JET_EXP = 70
    ROW_VOL_JET_TOT = 71
    # Fuel Oil
    ROW_VOL_FUELOIL_DOM = 74
    ROW_VOL_FUELOIL_EXP = 75
    ROW_VOL_FUELOIL_TOT = 76

    # ---- Revenue ----
    # Diesel
    ROW_REV_DIESEL_DOM = 81
    ROW_REV_DIESEL_EXP = 82
    ROW_REV_DIESEL_TOT = 83
    ROW_REV_DIESEL_ACTUAL = 84
    # Gasolines
    ROW_REV_GASOLINE_DOM = 87
    ROW_REV_GASOLINE_EXP = 88
    ROW_REV_GASOLINE_TOT = 89
    ROW_REV_GASOLINE_ACTUAL = 90
    # Jet Fuel
    ROW_REV_JET_DOM = 93
    ROW_REV_JET_EXP = 94
    ROW_REV_JET_TOT = 95
    ROW_REV_JET_ACTUAL = 96
    # Fuel Oil
    ROW_REV_FUELOIL_DOM = 99
    ROW_REV_FUELOIL_EXP = 100
    ROW_REV_FUELOIL_TOT = 101
    ROW_REV_FUELOIL_ACTUAL = 102

    ROW_REV_TOTAL = 104

    def _product_block(self, dom_price, exp_price, dom_vol, exp_vol, tot_vol,
                       dom_rev, exp_rev, tot_rev, actual_rev=None):
        d = {
            "price": {"domestic": self._series(dom_price), "export": self._series(exp_price)},
            "volume": {"domestic": self._series(dom_vol), "export": self._series(exp_vol),
                       "total": self._series(tot_vol)},
            "revenue": {"domestic": self._series(dom_rev), "export": self._series(exp_rev),
                        "total": self._series(tot_rev)},
        }
        if actual_rev:
            d["revenue"]["actual_total"] = self._series(actual_rev)
        return d

    @property
    def diesel(self): return self._product_block(
        self.ROW_PRICE_DIESEL_DOM, self.ROW_PRICE_DIESEL_EXP,
        self.ROW_VOL_DIESEL_DOM, self.ROW_VOL_DIESEL_EXP, self.ROW_VOL_DIESEL_TOT,
        self.ROW_REV_DIESEL_DOM, self.ROW_REV_DIESEL_EXP, self.ROW_REV_DIESEL_TOT,
        self.ROW_REV_DIESEL_ACTUAL)

    @property
    def gasolines(self): return self._product_block(
        self.ROW_PRICE_GASOLINE_DOM, self.ROW_PRICE_GASOLINE_EXP,
        self.ROW_VOL_GASOLINE_DOM, self.ROW_VOL_GASOLINE_EXP, self.ROW_VOL_GASOLINE_TOT,
        self.ROW_REV_GASOLINE_DOM, self.ROW_REV_GASOLINE_EXP, self.ROW_REV_GASOLINE_TOT,
        self.ROW_REV_GASOLINE_ACTUAL)

    @property
    def jet_fuel(self): return self._product_block(
        self.ROW_PRICE_JET_DOM, self.ROW_PRICE_JET_EXP,
        self.ROW_VOL_JET_DOM, self.ROW_VOL_JET_EXP, self.ROW_VOL_JET_TOT,
        self.ROW_REV_JET_DOM, self.ROW_REV_JET_EXP, self.ROW_REV_JET_TOT,
        self.ROW_REV_JET_ACTUAL)

    @property
    def fuel_oil(self): return self._product_block(
        self.ROW_PRICE_FUELOIL_DOM, self.ROW_PRICE_FUELOIL_EXP,
        self.ROW_VOL_FUELOIL_DOM, self.ROW_VOL_FUELOIL_EXP, self.ROW_VOL_FUELOIL_TOT,
        self.ROW_REV_FUELOIL_DOM, self.ROW_REV_FUELOIL_EXP, self.ROW_REV_FUELOIL_TOT,
        self.ROW_REV_FUELOIL_ACTUAL)

    @property
    def total_revenue(self) -> dict[int, float]:
        return self._series(self.ROW_REV_TOTAL)

    def summary(self) -> dict:
        return {
            "diesel": self.diesel, "gasolines": self.gasolines,
            "jet_fuel": self.jet_fuel, "fuel_oil": self.fuel_oil,
            "total_revenue": self.total_revenue,
        }


# ─────────────────────────────────────────────────────────
# 3. Other Products Revenue Schedule  (Rows 107–167)
# ─────────────────────────────────────────────────────────
class OtherProductsRevenueSchedule(BaseSchedule):
    SCHEDULE_NAME = "Other Products Revenue Schedule"
    START_ROW, END_ROW = 107, 167

    # Virgin naphtha
    ROW_PRICE_NAPHTHA_DOM = 117
    ROW_PRICE_NAPHTHA_EXP = 118
    ROW_VOL_NAPHTHA_DOM = 131
    ROW_VOL_NAPHTHA_EXP = 132
    ROW_VOL_NAPHTHA_TOT = 133
    ROW_REV_NAPHTHA_DOM = 148
    ROW_REV_NAPHTHA_EXP = 149
    ROW_REV_NAPHTHA_TOT = 150
    ROW_REV_NAPHTHA_ACTUAL = 151

    # Petrochemicals
    ROW_PRICE_PETROCHEM_DOM = 121
    ROW_PRICE_PETROCHEM_EXP = 122
    ROW_VOL_PETROCHEM_DOM = 136
    ROW_VOL_PETROCHEM_EXP = 137
    ROW_VOL_PETROCHEM_TOT = 138
    ROW_REV_PETROCHEM_DOM = 154
    ROW_REV_PETROCHEM_EXP = 155
    ROW_REV_PETROCHEM_TOT = 156
    ROW_REV_PETROCHEM_ACTUAL = 157

    # Fertilizers
    ROW_PRICE_FERT_DOM = 125
    ROW_PRICE_FERT_EXP = 126
    ROW_VOL_FERT_DOM = 141
    ROW_VOL_FERT_EXP = 142
    ROW_VOL_FERT_TOT = 143
    ROW_REV_FERT_DOM = 160
    ROW_REV_FERT_EXP = 161
    ROW_REV_FERT_TOT = 162
    ROW_REV_CROP_PROTECTION = 163
    ROW_REV_FERT_ACTUAL = 164

    ROW_REV_TOTAL = 166

    @property
    def virgin_naphtha(self):
        return {
            "price": {"domestic": self._series(self.ROW_PRICE_NAPHTHA_DOM),
                      "export": self._series(self.ROW_PRICE_NAPHTHA_EXP)},
            "volume": {"domestic": self._series(self.ROW_VOL_NAPHTHA_DOM),
                       "export": self._series(self.ROW_VOL_NAPHTHA_EXP),
                       "total": self._series(self.ROW_VOL_NAPHTHA_TOT)},
            "revenue": {"domestic": self._series(self.ROW_REV_NAPHTHA_DOM),
                        "export": self._series(self.ROW_REV_NAPHTHA_EXP),
                        "total": self._series(self.ROW_REV_NAPHTHA_TOT),
                        "actual_total": self._series(self.ROW_REV_NAPHTHA_ACTUAL)},
        }

    @property
    def petrochemicals(self):
        return {
            "price": {"domestic": self._series(self.ROW_PRICE_PETROCHEM_DOM),
                      "export": self._series(self.ROW_PRICE_PETROCHEM_EXP)},
            "volume": {"domestic": self._series(self.ROW_VOL_PETROCHEM_DOM),
                       "export": self._series(self.ROW_VOL_PETROCHEM_EXP),
                       "total": self._series(self.ROW_VOL_PETROCHEM_TOT)},
            "revenue": {"domestic": self._series(self.ROW_REV_PETROCHEM_DOM),
                        "export": self._series(self.ROW_REV_PETROCHEM_EXP),
                        "total": self._series(self.ROW_REV_PETROCHEM_TOT),
                        "actual_total": self._series(self.ROW_REV_PETROCHEM_ACTUAL)},
        }

    @property
    def fertilizers(self):
        return {
            "price": {"domestic": self._series(self.ROW_PRICE_FERT_DOM),
                      "export": self._series(self.ROW_PRICE_FERT_EXP)},
            "volume": {"domestic": self._series(self.ROW_VOL_FERT_DOM),
                       "export": self._series(self.ROW_VOL_FERT_EXP),
                       "total": self._series(self.ROW_VOL_FERT_TOT)},
            "revenue": {"domestic": self._series(self.ROW_REV_FERT_DOM),
                        "export": self._series(self.ROW_REV_FERT_EXP),
                        "total": self._series(self.ROW_REV_FERT_TOT),
                        "crop_protection": self._series(self.ROW_REV_CROP_PROTECTION),
                        "actual_total": self._series(self.ROW_REV_FERT_ACTUAL)},
        }

    @property
    def total_revenue(self): return self._series(self.ROW_REV_TOTAL)

    def summary(self) -> dict:
        return {
            "virgin_naphtha": self.virgin_naphtha,
            "petrochemicals": self.petrochemicals,
            "fertilizers": self.fertilizers,
            "total_revenue": self.total_revenue,
        }


# ─────────────────────────────────────────────────────────
# 4. Downstream Revenue Schedule  (Rows 169–214)
# ─────────────────────────────────────────────────────────
class DownstreamRevenueSchedule(BaseSchedule):
    SCHEDULE_NAME = "Downstream Revenue Schedule"
    START_ROW, END_ROW = 169, 214

    # Prices
    ROW_PRICE_BASE_OILS = 178
    ROW_PRICE_COKE = 179
    ROW_PRICE_LPG = 180
    ROW_PRICE_ASPHALT = 181
    ROW_PRICE_FUEL_OIL = 183
    ROW_PRICE_DIESEL = 184
    ROW_PRICE_GASOLINES = 185
    ROW_PRICE_PETROCHEM_NAPHTHA = 186
    ROW_PRICE_JET_FUEL = 187

    # Volumes
    ROW_VOL_BASE_OILS = 190
    ROW_VOL_COKE = 191
    ROW_VOL_LPG = 192
    ROW_VOL_ASPHALT = 193
    ROW_VOL_FUEL_OIL = 195
    ROW_VOL_DIESEL = 196
    ROW_VOL_GASOLINES = 197
    ROW_VOL_PETROCHEM_NAPHTHA = 198
    ROW_VOL_JET_FUEL = 199

    # Revenue
    ROW_REV_LUBRICANTS = 202
    ROW_REV_COKE = 203
    ROW_REV_LPG = 204
    ROW_REV_ASPHALT = 205
    ROW_REV_SUBTOTAL = 206
    ROW_REV_FUEL_OIL = 208
    ROW_REV_DIESEL = 209
    ROW_REV_GASOLINES = 210
    ROW_REV_PETROCHEM_NAPHTHA = 211
    ROW_REV_JET_FUEL = 212
    ROW_REV_SUBTOTAL_2 = 213

    @property
    def prices(self):
        return {
            "base_oils": self._series(self.ROW_PRICE_BASE_OILS),
            "coke": self._series(self.ROW_PRICE_COKE),
            "lpg": self._series(self.ROW_PRICE_LPG),
            "asphalt": self._series(self.ROW_PRICE_ASPHALT),
            "fuel_oil": self._series(self.ROW_PRICE_FUEL_OIL),
            "diesel": self._series(self.ROW_PRICE_DIESEL),
            "gasolines": self._series(self.ROW_PRICE_GASOLINES),
            "petrochem_naphtha": self._series(self.ROW_PRICE_PETROCHEM_NAPHTHA),
            "jet_fuel": self._series(self.ROW_PRICE_JET_FUEL),
        }

    @property
    def volumes(self):
        return {
            "base_oils": self._series(self.ROW_VOL_BASE_OILS),
            "coke": self._series(self.ROW_VOL_COKE),
            "lpg": self._series(self.ROW_VOL_LPG),
            "asphalt": self._series(self.ROW_VOL_ASPHALT),
            "fuel_oil": self._series(self.ROW_VOL_FUEL_OIL),
            "diesel": self._series(self.ROW_VOL_DIESEL),
            "gasolines": self._series(self.ROW_VOL_GASOLINES),
            "petrochem_naphtha": self._series(self.ROW_VOL_PETROCHEM_NAPHTHA),
            "jet_fuel": self._series(self.ROW_VOL_JET_FUEL),
        }

    @property
    def revenue(self):
        return {
            "lubricants_byproducts": self._series(self.ROW_REV_LUBRICANTS),
            "petroleum_coke": self._series(self.ROW_REV_COKE),
            "lpg": self._series(self.ROW_REV_LPG),
            "asphalts": self._series(self.ROW_REV_ASPHALT),
            "subtotal_1": self._series(self.ROW_REV_SUBTOTAL),
            "fuel_oil": self._series(self.ROW_REV_FUEL_OIL),
            "diesel": self._series(self.ROW_REV_DIESEL),
            "gasolines": self._series(self.ROW_REV_GASOLINES),
            "petrochem_naphtha": self._series(self.ROW_REV_PETROCHEM_NAPHTHA),
            "jet_fuel": self._series(self.ROW_REV_JET_FUEL),
            "subtotal_2": self._series(self.ROW_REV_SUBTOTAL_2),
        }

    def summary(self) -> dict:
        return {"prices": self.prices, "volumes": self.volumes, "revenue": self.revenue}


# ─────────────────────────────────────────────────────────
# 5. Total Revenue Schedule  (Rows 216–277)
# ─────────────────────────────────────────────────────────
class TotalRevenueSchedule(BaseSchedule):
    SCHEDULE_NAME = "Total Revenue Schedule"
    START_ROW, END_ROW = 216, 277

    # Prices
    ROW_PRICE_NG_DOM = 226
    ROW_PRICE_NG_EXP = 227
    ROW_PRICE_CRUDE_DOM = 230
    ROW_PRICE_CRUDE_EXP = 231

    # Volumes
    ROW_VOL_NG_DOM = 236
    ROW_VOL_NG_EXP = 237
    ROW_VOL_NG_TOT = 238
    ROW_VOL_CRUDE_DOM = 241
    ROW_VOL_CRUDE_EXP = 242
    ROW_VOL_CRUDE_TOT = 243

    # Revenue
    ROW_REV_NG_DOM = 248
    ROW_REV_NG_EXP = 249
    ROW_REV_NG_TOT = 250
    ROW_REV_NG_ACTUAL = 251
    ROW_REV_CRUDE_DOM = 254
    ROW_REV_CRUDE_EXP = 255
    ROW_REV_CRUDE_TOT = 256
    ROW_REV_CRUDE_ACTUAL = 257
    ROW_REV_CRUDE_AND_NG = 259

    # Macro
    ROW_ARG_GDP = 262

    # Revenue components
    ROW_MAIN_CRUDE_PRODUCTS = 265
    ROW_OTHER_PRODUCTS = 266
    ROW_DOWNSTREAM = 267

    # Other revenue
    ROW_OTHER_GAS_STATIONS = 270
    ROW_OTHER_CONSTRUCTION = 271
    ROW_OTHER_LNG = 272
    ROW_OTHER_GOODS_SERVICES = 273
    ROW_OTHER_SUBTOTAL = 274

    ROW_TOTAL_REVENUE = 276

    @property
    def natural_gas(self):
        return {
            "price": {"domestic": self._series(self.ROW_PRICE_NG_DOM),
                      "export": self._series(self.ROW_PRICE_NG_EXP)},
            "volume": {"domestic": self._series(self.ROW_VOL_NG_DOM),
                       "export": self._series(self.ROW_VOL_NG_EXP),
                       "total": self._series(self.ROW_VOL_NG_TOT)},
            "revenue": {"domestic": self._series(self.ROW_REV_NG_DOM),
                        "export": self._series(self.ROW_REV_NG_EXP),
                        "total": self._series(self.ROW_REV_NG_TOT)},
        }

    @property
    def crude_oil(self):
        return {
            "price": {"domestic": self._series(self.ROW_PRICE_CRUDE_DOM),
                      "export": self._series(self.ROW_PRICE_CRUDE_EXP)},
            "volume": {"domestic": self._series(self.ROW_VOL_CRUDE_DOM),
                       "export": self._series(self.ROW_VOL_CRUDE_EXP),
                       "total": self._series(self.ROW_VOL_CRUDE_TOT)},
            "revenue": {"domestic": self._series(self.ROW_REV_CRUDE_DOM),
                        "export": self._series(self.ROW_REV_CRUDE_EXP),
                        "total": self._series(self.ROW_REV_CRUDE_TOT)},
        }

    @property
    def argentina_gdp(self): return self._series(self.ROW_ARG_GDP)

    @property
    def revenue_components(self):
        return {
            "main_crude_products": self._series(self.ROW_MAIN_CRUDE_PRODUCTS),
            "other_products": self._series(self.ROW_OTHER_PRODUCTS),
            "downstream": self._series(self.ROW_DOWNSTREAM),
        }

    @property
    def other_revenue(self):
        return {
            "gas_stations": self._series(self.ROW_OTHER_GAS_STATIONS),
            "construction_contracts": self._series(self.ROW_OTHER_CONSTRUCTION),
            "lng_regasification": self._series(self.ROW_OTHER_LNG),
            "other_goods_services": self._series(self.ROW_OTHER_GOODS_SERVICES),
            "subtotal": self._series(self.ROW_OTHER_SUBTOTAL),
        }

    @property
    def total_revenue(self): return self._series(self.ROW_TOTAL_REVENUE)

    def summary(self) -> dict:
        return {
            "natural_gas": self.natural_gas, "crude_oil": self.crude_oil,
            "argentina_gdp": self.argentina_gdp,
            "revenue_components": self.revenue_components,
            "other_revenue": self.other_revenue,
            "total_revenue": self.total_revenue,
        }


# ─────────────────────────────────────────────────────────
# 6. Production Costs Expenses Schedule  (Rows 279–332)
# ─────────────────────────────────────────────────────────
class ProductionCostsSchedule(BaseSchedule):
    SCHEDULE_NAME = "Production Costs Expenses Schedule"
    START_ROW, END_ROW = 279, 332

    ROW_OIL_PRICES = 287
    ROW_VOLUMES_GROWTH = 288
    ROW_REVENUE = 290

    # Royalties & Fees
    ROW_ROYALTIES_FEES = 293
    ROW_FEES_COMPENSATION = 294
    ROW_ROYALTIES_TOTAL = 295
    ROW_ROYALTIES_PCT = 297

    # Macro
    ROW_ARG_INFLATION = 300
    ROW_USA_INFLATION = 301
    ROW_FX_RATE = 303
    ROW_DEPRECIATION_RATE = 304

    # Grows with Arg Inflation & Oil/NG produced
    ROW_SALARIES = 307
    ROW_OTHER_PERSONNEL = 308
    ROW_RENTAL = 309
    ROW_TRANSPORTATION = 310
    ROW_PRESERVATION = 311
    ROW_OPERATION_SERVICES = 312
    ROW_TAXES_CHARGES = 313

    # Grows with USA Inflation & Oil/NG produced
    ROW_INDUSTRIAL_INPUTS = 316
    ROW_INSURANCE = 317

    # Grows with Oil Prices
    ROW_FUEL_GAS_ENERGY = 320

    ROW_TOTAL_PRODUCTION_COSTS = 322

    # Total Costs of Sale
    ROW_INV_BEGINNING = 326
    ROW_PURCHASES = 327
    ROW_PROD_COSTS = 328
    ROW_CURRENCY_CONVERSIONS = 329
    ROW_INV_ENDING = 330
    ROW_TOTAL_COST_OF_SALE = 331

    @property
    def macro(self):
        return {
            "oil_prices": self._series(self.ROW_OIL_PRICES),
            "volumes_growth": self._series(self.ROW_VOLUMES_GROWTH),
            "argentina_inflation": self._series(self.ROW_ARG_INFLATION),
            "usa_inflation": self._series(self.ROW_USA_INFLATION),
            "fx_rate": self._series(self.ROW_FX_RATE),
            "depreciation_rate": self._series(self.ROW_DEPRECIATION_RATE),
        }

    @property
    def royalties_and_fees(self):
        return {
            "royalties_easements": self._series(self.ROW_ROYALTIES_FEES),
            "fees_compensation": self._series(self.ROW_FEES_COMPENSATION),
            "total": self._series(self.ROW_ROYALTIES_TOTAL),
            "pct_revenue": self._series(self.ROW_ROYALTIES_PCT),
        }

    @property
    def arg_inflation_linked(self):
        return {
            "salaries": self._series(self.ROW_SALARIES),
            "other_personnel": self._series(self.ROW_OTHER_PERSONNEL),
            "rental": self._series(self.ROW_RENTAL),
            "transportation": self._series(self.ROW_TRANSPORTATION),
            "preservation_repair": self._series(self.ROW_PRESERVATION),
            "operation_services": self._series(self.ROW_OPERATION_SERVICES),
            "taxes_charges": self._series(self.ROW_TAXES_CHARGES),
        }

    @property
    def usa_inflation_linked(self):
        return {
            "industrial_inputs": self._series(self.ROW_INDUSTRIAL_INPUTS),
            "insurance": self._series(self.ROW_INSURANCE),
        }

    @property
    def oil_price_linked(self):
        return {"fuel_gas_energy": self._series(self.ROW_FUEL_GAS_ENERGY)}

    @property
    def total_production_costs(self): return self._series(self.ROW_TOTAL_PRODUCTION_COSTS)

    @property
    def cost_of_sale(self):
        return {
            "inventories_beginning": self._series(self.ROW_INV_BEGINNING),
            "purchases": self._series(self.ROW_PURCHASES),
            "production_costs": self._series(self.ROW_PROD_COSTS),
            "currency_conversions": self._series(self.ROW_CURRENCY_CONVERSIONS),
            "inventories_ending": self._series(self.ROW_INV_ENDING),
            "total": self._series(self.ROW_TOTAL_COST_OF_SALE),
        }

    def summary(self) -> dict:
        return {
            "macro": self.macro,
            "royalties_and_fees": self.royalties_and_fees,
            "arg_inflation_linked_costs": self.arg_inflation_linked,
            "usa_inflation_linked_costs": self.usa_inflation_linked,
            "oil_price_linked_costs": self.oil_price_linked,
            "total_production_costs": self.total_production_costs,
            "cost_of_sale": self.cost_of_sale,
        }


# ─────────────────────────────────────────────────────────
# 7. S&A Expenses Schedule  (Rows 334–390)
# ─────────────────────────────────────────────────────────
class SellingAndAdminExpensesSchedule(BaseSchedule):
    SCHEDULE_NAME = "S&A Expenses Schedule"
    START_ROW, END_ROW = 334, 390

    ROW_OIL_PRICES = 342
    ROW_VOLUMES_GROWTH = 343
    ROW_USA_INFLATION = 344

    # Selling – USA Inflation
    ROW_SELL_SALARIES = 348
    ROW_SELL_FEES = 349
    ROW_SELL_OTHER_PERSONNEL = 350
    ROW_SELL_TAXES = 351
    ROW_SELL_ROYALTIES = 352
    ROW_SELL_INSURANCE = 353
    ROW_SELL_RENTAL = 354
    ROW_SELL_INDUSTRIAL_INPUTS = 355
    ROW_SELL_OPERATION_SERVICES = 356
    ROW_SELL_PRESERVATION = 357
    ROW_SELL_TRANSPORTATION = 358
    ROW_SELL_PUBLICITY = 359
    # Selling – Constant
    ROW_SELL_DOUBTFUL = 362
    # Selling – Oil Prices
    ROW_SELL_FUEL_GAS = 365
    ROW_SELL_TOTAL = 367

    # Admin – USA Inflation
    ROW_ADMIN_SALARIES = 372
    ROW_ADMIN_FEES = 373
    ROW_ADMIN_OTHER_PERSONNEL = 374
    ROW_ADMIN_TAXES = 375
    ROW_ADMIN_OPERATION_SERVICES = 376
    ROW_ADMIN_PRESERVATION = 377
    ROW_ADMIN_PUBLICITY = 378
    ROW_ADMIN_OTHER = 379
    # Admin – Oil Prices
    ROW_ADMIN_FUEL_GAS = 382
    ROW_ADMIN_TOTAL = 384

    # Exploration
    ROW_EXPLORATION_TOTAL = 389

    @property
    def selling_expenses(self):
        return {
            "salaries": self._series(self.ROW_SELL_SALARIES),
            "fees": self._series(self.ROW_SELL_FEES),
            "other_personnel": self._series(self.ROW_SELL_OTHER_PERSONNEL),
            "taxes": self._series(self.ROW_SELL_TAXES),
            "royalties": self._series(self.ROW_SELL_ROYALTIES),
            "insurance": self._series(self.ROW_SELL_INSURANCE),
            "rental": self._series(self.ROW_SELL_RENTAL),
            "industrial_inputs": self._series(self.ROW_SELL_INDUSTRIAL_INPUTS),
            "operation_services": self._series(self.ROW_SELL_OPERATION_SERVICES),
            "preservation": self._series(self.ROW_SELL_PRESERVATION),
            "transportation": self._series(self.ROW_SELL_TRANSPORTATION),
            "publicity": self._series(self.ROW_SELL_PUBLICITY),
            "doubtful_receivables": self._series(self.ROW_SELL_DOUBTFUL),
            "fuel_gas_energy": self._series(self.ROW_SELL_FUEL_GAS),
            "total": self._series(self.ROW_SELL_TOTAL),
        }

    @property
    def admin_expenses(self):
        return {
            "salaries": self._series(self.ROW_ADMIN_SALARIES),
            "fees": self._series(self.ROW_ADMIN_FEES),
            "other_personnel": self._series(self.ROW_ADMIN_OTHER_PERSONNEL),
            "taxes": self._series(self.ROW_ADMIN_TAXES),
            "operation_services": self._series(self.ROW_ADMIN_OPERATION_SERVICES),
            "preservation": self._series(self.ROW_ADMIN_PRESERVATION),
            "publicity": self._series(self.ROW_ADMIN_PUBLICITY),
            "other": self._series(self.ROW_ADMIN_OTHER),
            "fuel_gas_energy": self._series(self.ROW_ADMIN_FUEL_GAS),
            "total": self._series(self.ROW_ADMIN_TOTAL),
        }

    @property
    def exploration_expenses(self): return self._series(self.ROW_EXPLORATION_TOTAL)

    def summary(self) -> dict:
        return {
            "selling_expenses": self.selling_expenses,
            "admin_expenses": self.admin_expenses,
            "exploration_expenses": self.exploration_expenses,
        }


# ─────────────────────────────────────────────────────────
# 8. Income Statement  (Rows 393–437)
# ─────────────────────────────────────────────────────────
class IncomeStatement(BaseSchedule):
    SCHEDULE_NAME = "Income Statement"
    START_ROW, END_ROW = 393, 437

    ROW_REVENUE = 401
    ROW_COGS = 402
    ROW_GROSS_PROFIT = 403
    ROW_SELLING = 405
    ROW_ADMIN = 406
    ROW_EXPLORATION = 407
    ROW_OTHER = 408
    ROW_OPERATING_COSTS = 409
    ROW_IMPAIRMENT = 411
    ROW_EBITDA = 412
    ROW_DA = 414
    ROW_EBIT = 415
    ROW_EQUITY_INCOME = 417
    ROW_FINANCIAL_INCOME = 418
    ROW_FINANCIAL_COSTS = 419
    ROW_OTHER_FINANCIAL = 420
    ROW_EBT = 421
    ROW_INCOME_TAX = 424
    ROW_NET_INCOME = 425
    ROW_REVENUE_GROWTH = 427
    ROW_COGS_GROWTH = 428
    ROW_GROSS_MARGIN = 429
    ROW_NOPAT = 431
    ROW_EBITDA_MARGIN = 434
    ROW_EBIT_MARGIN = 435
    ROW_ROE = 436

    @property
    def line_items(self):
        return {
            "revenue": self._series(self.ROW_REVENUE),
            "cost_of_sales": self._series(self.ROW_COGS),
            "gross_profit": self._series(self.ROW_GROSS_PROFIT),
            "selling_expenses": self._series(self.ROW_SELLING),
            "admin_expenses": self._series(self.ROW_ADMIN),
            "exploration_expenses": self._series(self.ROW_EXPLORATION),
            "other": self._series(self.ROW_OTHER),
            "operating_costs": self._series(self.ROW_OPERATING_COSTS),
            "impairment": self._series(self.ROW_IMPAIRMENT),
            "ebitda": self._series(self.ROW_EBITDA),
            "da": self._series(self.ROW_DA),
            "ebit": self._series(self.ROW_EBIT),
            "equity_income": self._series(self.ROW_EQUITY_INCOME),
            "financial_income": self._series(self.ROW_FINANCIAL_INCOME),
            "financial_costs": self._series(self.ROW_FINANCIAL_COSTS),
            "other_financial": self._series(self.ROW_OTHER_FINANCIAL),
            "ebt": self._series(self.ROW_EBT),
            "income_tax": self._series(self.ROW_INCOME_TAX),
            "net_income": self._series(self.ROW_NET_INCOME),
            "nopat": self._series(self.ROW_NOPAT),
        }

    @property
    def margins(self):
        return {
            "revenue_growth": self._series(self.ROW_REVENUE_GROWTH),
            "cogs_growth": self._series(self.ROW_COGS_GROWTH),
            "gross_margin": self._series(self.ROW_GROSS_MARGIN),
            "ebitda_margin": self._series(self.ROW_EBITDA_MARGIN),
            "ebit_margin": self._series(self.ROW_EBIT_MARGIN),
            "roe": self._series(self.ROW_ROE),
        }

    def summary(self) -> dict:
        return {"line_items": self.line_items, "margins": self.margins}


# ─────────────────────────────────────────────────────────
# 9. Cash Flow Statement  (Rows 441–491)
# ─────────────────────────────────────────────────────────
class CashFlowStatement(BaseSchedule):
    SCHEDULE_NAME = "Cash Flow Statement"
    START_ROW, END_ROW = 441, 491

    # Operating
    ROW_NET_INCOME = 449
    ROW_EQUITY_INTERESTS = 450
    ROW_DEPRECIATION_PPE = 451
    ROW_AMORTIZATION_IA = 452
    ROW_DEPRECIATION_ROU = 453
    ROW_RETIREMENT_PPE = 454
    ROW_IMPAIRMENT = 455
    ROW_INCOME_TAX_CHARGE = 456
    ROW_PROVISIONS = 457
    ROW_FX_INTEREST_OTHER = 458
    ROW_SHARE_BASED = 459
    ROW_OTHER_INSURANCE = 460
    ROW_OTHER_OPERATIONS = 461
    ROW_WORKING_CAPITAL = 462
    ROW_CF_OPERATING = 463

    # Investing
    ROW_CAPEX = 466
    ROW_ASSETS_HELD_SALE = 467
    ROW_ACQUISITIONS_JV = 468
    ROW_LOANS_RELATED = 469
    ROW_PROCEEDS_FINANCIAL = 470
    ROW_PAYMENTS_FINANCIAL = 471
    ROW_INTEREST_FINANCIAL = 472
    ROW_COLLECTIONS_CONCESSIONS = 473
    ROW_CF_INVESTING = 474

    # Financing
    ROW_LOAN_PAYMENTS = 477
    ROW_LOAN_PROCEEDS = 478
    ROW_INTEREST_PAYMENTS = 479
    ROW_OVERDRAFT = 480
    ROW_BUYBACK = 481
    ROW_LEASE_PAYMENTS = 482
    ROW_TAX_INTEREST_PAYMENTS = 483
    ROW_CF_FINANCING = 484

    ROW_FX_EFFECT = 486
    ROW_CHANGE_CASH = 488
    ROW_BEGINNING_CASH = 489
    ROW_ENDING_CASH = 490

    @property
    def operating(self):
        return {
            "net_income": self._series(self.ROW_NET_INCOME),
            "equity_interests": self._series(self.ROW_EQUITY_INTERESTS),
            "depreciation_ppe": self._series(self.ROW_DEPRECIATION_PPE),
            "amortization_ia": self._series(self.ROW_AMORTIZATION_IA),
            "depreciation_rou": self._series(self.ROW_DEPRECIATION_ROU),
            "retirement_ppe": self._series(self.ROW_RETIREMENT_PPE),
            "impairment": self._series(self.ROW_IMPAIRMENT),
            "income_tax_charge": self._series(self.ROW_INCOME_TAX_CHARGE),
            "provisions": self._series(self.ROW_PROVISIONS),
            "fx_interest_other": self._series(self.ROW_FX_INTEREST_OTHER),
            "working_capital": self._series(self.ROW_WORKING_CAPITAL),
            "total": self._series(self.ROW_CF_OPERATING),
        }

    @property
    def investing(self):
        return {
            "capex": self._series(self.ROW_CAPEX),
            "assets_held_for_sale": self._series(self.ROW_ASSETS_HELD_SALE),
            "acquisitions_jv": self._series(self.ROW_ACQUISITIONS_JV),
            "total": self._series(self.ROW_CF_INVESTING),
        }

    @property
    def financing(self):
        return {
            "loan_payments": self._series(self.ROW_LOAN_PAYMENTS),
            "loan_proceeds": self._series(self.ROW_LOAN_PROCEEDS),
            "interest_payments": self._series(self.ROW_INTEREST_PAYMENTS),
            "overdraft": self._series(self.ROW_OVERDRAFT),
            "buyback": self._series(self.ROW_BUYBACK),
            "lease_payments": self._series(self.ROW_LEASE_PAYMENTS),
            "total": self._series(self.ROW_CF_FINANCING),
        }

    @property
    def cash_position(self):
        return {
            "change": self._series(self.ROW_CHANGE_CASH),
            "beginning": self._series(self.ROW_BEGINNING_CASH),
            "ending": self._series(self.ROW_ENDING_CASH),
        }

    def summary(self) -> dict:
        return {
            "operating": self.operating, "investing": self.investing,
            "financing": self.financing, "cash_position": self.cash_position,
        }


# ─────────────────────────────────────────────────────────
# 10. Balance Sheet  (Rows 494–557)
# ─────────────────────────────────────────────────────────
class BalanceSheet(BaseSchedule):
    SCHEDULE_NAME = "Balance Sheet"
    START_ROW, END_ROW = 494, 557

    # Current Assets
    ROW_CASH = 502
    ROW_INVESTMENTS = 503
    ROW_TRADE_RECV_CA = 504
    ROW_CONTRACT_ASSET = 505
    ROW_OTHER_RECV_CA = 506
    ROW_INVENTORIES = 507
    ROW_ASSETS_HELD_SALE = 508
    ROW_TOTAL_CA = 509

    # Non-Current Assets
    ROW_INV_FINANCIAL = 512
    ROW_TRADE_RECV_NCA = 513
    ROW_OTHER_RECV_NCA = 514
    ROW_DEFERRED_TAX = 515
    ROW_INV_ASSOCIATES = 516
    ROW_ROU_ASSETS = 517
    ROW_PPE = 518
    ROW_INTANGIBLE = 519
    ROW_TOTAL_ASSETS = 520

    # Current Liabilities
    ROW_AP_CL = 523
    ROW_OTHER_LIAB_CL = 524
    ROW_LOANS_CL = 525
    ROW_LEASE_CL = 526
    ROW_SALARIES_CL = 527
    ROW_TAX_CL = 528
    ROW_INCOME_TAX_CL = 529
    ROW_CONTRACT_LIAB_CL = 530
    ROW_PROVISIONS_CL = 531
    ROW_LIAB_HELD_SALE = 532
    ROW_TOTAL_CL = 533

    # Non-Current Liabilities
    ROW_AP_NCL = 536
    ROW_OTHER_LIAB_NCL = 537
    ROW_LOANS_NCL = 538
    ROW_LEASE_NCL = 539
    ROW_SALARIES_NCL = 540
    ROW_TAX_NCL = 541
    ROW_INCOME_TAX_NCL = 542
    ROW_DEFERRED_TAX_LIAB = 543
    ROW_CONTRACT_LIAB_NCL = 544
    ROW_PROVISIONS_NCL = 545
    ROW_TOTAL_NCL = 546

    # Shareholders' Equity
    ROW_COMMON_STOCK = 549
    ROW_RETAINED_EARNINGS = 550
    ROW_MINORITY_INTEREST = 551
    ROW_TOTAL_EQUITY = 552
    ROW_TOTAL_LIAB_EQUITY = 554

    ROW_CHECK = 557

    @property
    def current_assets(self):
        return {
            "cash": self._series(self.ROW_CASH),
            "investments": self._series(self.ROW_INVESTMENTS),
            "trade_receivables": self._series(self.ROW_TRADE_RECV_CA),
            "contract_asset": self._series(self.ROW_CONTRACT_ASSET),
            "other_receivables": self._series(self.ROW_OTHER_RECV_CA),
            "inventories": self._series(self.ROW_INVENTORIES),
            "assets_held_for_sale": self._series(self.ROW_ASSETS_HELD_SALE),
            "total": self._series(self.ROW_TOTAL_CA),
        }

    @property
    def non_current_assets(self):
        return {
            "financial_investments": self._series(self.ROW_INV_FINANCIAL),
            "trade_receivables": self._series(self.ROW_TRADE_RECV_NCA),
            "other_receivables": self._series(self.ROW_OTHER_RECV_NCA),
            "deferred_tax_asset": self._series(self.ROW_DEFERRED_TAX),
            "associates_jv": self._series(self.ROW_INV_ASSOCIATES),
            "rou_assets": self._series(self.ROW_ROU_ASSETS),
            "ppe": self._series(self.ROW_PPE),
            "intangible": self._series(self.ROW_INTANGIBLE),
        }

    @property
    def current_liabilities(self):
        return {
            "accounts_payable": self._series(self.ROW_AP_CL),
            "other_liabilities": self._series(self.ROW_OTHER_LIAB_CL),
            "loans": self._series(self.ROW_LOANS_CL),
            "lease_liabilities": self._series(self.ROW_LEASE_CL),
            "salaries_ss": self._series(self.ROW_SALARIES_CL),
            "taxes_payable": self._series(self.ROW_TAX_CL),
            "income_tax_payable": self._series(self.ROW_INCOME_TAX_CL),
            "contract_liabilities": self._series(self.ROW_CONTRACT_LIAB_CL),
            "provisions": self._series(self.ROW_PROVISIONS_CL),
            "liab_held_for_sale": self._series(self.ROW_LIAB_HELD_SALE),
            "total": self._series(self.ROW_TOTAL_CL),
        }

    @property
    def non_current_liabilities(self):
        return {
            "accounts_payable": self._series(self.ROW_AP_NCL),
            "other_liabilities": self._series(self.ROW_OTHER_LIAB_NCL),
            "loans": self._series(self.ROW_LOANS_NCL),
            "lease_liabilities": self._series(self.ROW_LEASE_NCL),
            "salaries_ss": self._series(self.ROW_SALARIES_NCL),
            "taxes_payable": self._series(self.ROW_TAX_NCL),
            "income_tax_payable": self._series(self.ROW_INCOME_TAX_NCL),
            "deferred_tax_liabilities": self._series(self.ROW_DEFERRED_TAX_LIAB),
            "contract_liabilities": self._series(self.ROW_CONTRACT_LIAB_NCL),
            "provisions": self._series(self.ROW_PROVISIONS_NCL),
            "total": self._series(self.ROW_TOTAL_NCL),
        }

    @property
    def shareholders_equity(self):
        return {
            "common_stock": self._series(self.ROW_COMMON_STOCK),
            "retained_earnings": self._series(self.ROW_RETAINED_EARNINGS),
            "minority_interest": self._series(self.ROW_MINORITY_INTEREST),
            "total": self._series(self.ROW_TOTAL_EQUITY),
        }

    @property
    def total_assets(self): return self._series(self.ROW_TOTAL_ASSETS)

    @property
    def total_liabilities_and_equity(self): return self._series(self.ROW_TOTAL_LIAB_EQUITY)

    @property
    def check(self): return self._series(self.ROW_CHECK)

    def summary(self) -> dict:
        return {
            "current_assets": self.current_assets,
            "non_current_assets": self.non_current_assets,
            "total_assets": self.total_assets,
            "current_liabilities": self.current_liabilities,
            "non_current_liabilities": self.non_current_liabilities,
            "shareholders_equity": self.shareholders_equity,
            "total_liabilities_and_equity": self.total_liabilities_and_equity,
            "check": self.check,
        }


# ─────────────────────────────────────────────────────────
# 11. Fixed Assets (PP&E) Schedule  (Rows 560–607)
# ─────────────────────────────────────────────────────────
class FixedAssetsSchedule(BaseSchedule):
    SCHEDULE_NAME = "Fixed Assets (PP&E) Schedule"
    START_ROW, END_ROW = 560, 607

    ROW_PPE_BEGINNING = 569
    ROW_CAPEX = 570
    ROW_NEW_INTANGIBLES = 571

    # Depreciation of PP&E
    ROW_DEPR_PROD_COSTS = 574
    ROW_DEPR_SELLING = 575
    ROW_DEPR_ADMIN = 576
    ROW_DEPR_TOTAL = 577

    # Depreciation % ratios
    ROW_PCT_PROD_DEPR_CAPEX = 579
    ROW_PCT_SELL_DEPR_CAPEX = 580
    ROW_PCT_ADMIN_DEPR_CAPEX = 581

    # Amortization of IA
    ROW_AMORT_PROD = 584
    ROW_AMORT_SELLING = 585
    ROW_AMORT_ADMIN = 586
    ROW_AMORT_TOTAL = 587

    ROW_IMPAIRMENT = 589
    ROW_PPE_ENDING = 591

    # ROU Assets
    ROW_ROU_BEGINNING = 595
    ROW_ROU_ADDITIONS = 596
    ROW_ROU_DEPR_PROD = 599
    ROW_ROU_DEPR_SELLING = 600
    ROW_ROU_DEPR_TOTAL = 601
    ROW_ROU_ENDING = 603

    ROW_TOTAL_DA = 606

    @property
    def ppe(self):
        return {
            "beginning": self._series(self.ROW_PPE_BEGINNING),
            "capex": self._series(self.ROW_CAPEX),
            "new_intangibles": self._series(self.ROW_NEW_INTANGIBLES),
            "depreciation": {
                "production_costs": self._series(self.ROW_DEPR_PROD_COSTS),
                "selling": self._series(self.ROW_DEPR_SELLING),
                "admin": self._series(self.ROW_DEPR_ADMIN),
                "total": self._series(self.ROW_DEPR_TOTAL),
            },
            "depreciation_pct": {
                "prod_to_capex": self._series(self.ROW_PCT_PROD_DEPR_CAPEX),
                "sell_to_capex": self._series(self.ROW_PCT_SELL_DEPR_CAPEX),
                "admin_to_capex": self._series(self.ROW_PCT_ADMIN_DEPR_CAPEX),
            },
            "amortization": {
                "production_costs": self._series(self.ROW_AMORT_PROD),
                "selling": self._series(self.ROW_AMORT_SELLING),
                "admin": self._series(self.ROW_AMORT_ADMIN),
                "total": self._series(self.ROW_AMORT_TOTAL),
            },
            "impairment": self._series(self.ROW_IMPAIRMENT),
            "ending": self._series(self.ROW_PPE_ENDING),
        }

    @property
    def rou_assets(self):
        return {
            "beginning": self._series(self.ROW_ROU_BEGINNING),
            "additions": self._series(self.ROW_ROU_ADDITIONS),
            "depreciation": {
                "production_costs": self._series(self.ROW_ROU_DEPR_PROD),
                "selling": self._series(self.ROW_ROU_DEPR_SELLING),
                "total": self._series(self.ROW_ROU_DEPR_TOTAL),
            },
            "ending": self._series(self.ROW_ROU_ENDING),
        }

    @property
    def total_da(self): return self._series(self.ROW_TOTAL_DA)

    def summary(self) -> dict:
        return {"ppe": self.ppe, "rou_assets": self.rou_assets, "total_da": self.total_da}


# ─────────────────────────────────────────────────────────
# 12. Working Capital Schedule  (Rows 609–680)
# ─────────────────────────────────────────────────────────
class WorkingCapitalSchedule(BaseSchedule):
    SCHEDULE_NAME = "Working Capital Schedule"
    START_ROW, END_ROW = 609, 680

    ROW_DAYS_PER_YEAR = 617
    ROW_NET_REVENUE = 620
    ROW_COST_OF_SALES = 621
    ROW_ROU_ASSETS = 622

    # Days-in for Current items
    ROW_DAYS_TRADE_RECV = 626
    ROW_DAYS_CONTRACT_ASSET = 627
    ROW_DAYS_OTHER_RECV = 628
    ROW_DAYS_INVENTORIES = 629
    ROW_DAYS_AP = 630
    ROW_DAYS_OTHER_LIAB = 631
    ROW_DAYS_LEASE_LIAB = 632
    ROW_DAYS_SALARIES = 633
    ROW_DAYS_TAXES = 634
    ROW_DAYS_INCOME_TAX = 635
    ROW_DAYS_CONTRACT_LIAB = 636
    ROW_DAYS_PROVISIONS = 637

    # Days-in for Non-Current items
    ROW_DAYS_NC_TRADE_RECV = 640
    ROW_DAYS_NC_OTHER_RECV = 641
    ROW_DAYS_NC_AP = 642
    ROW_DAYS_NC_OTHER_LIAB = 643
    ROW_DAYS_NC_LEASE_LIAB = 644
    ROW_DAYS_NC_SALARIES = 645
    ROW_DAYS_NC_TAXES = 646
    ROW_DAYS_NC_INCOME_TAX = 647
    ROW_DAYS_NC_CONTRACT_LIAB = 648
    ROW_DAYS_NC_PROVISIONS = 649

    ROW_NET_WORKING_CAPITAL = 677
    ROW_CHANGE_WC = 679

    @property
    def days_in(self):
        return {
            "current": {
                "trade_receivables": self._series(self.ROW_DAYS_TRADE_RECV),
                "contract_asset": self._series(self.ROW_DAYS_CONTRACT_ASSET),
                "other_receivables": self._series(self.ROW_DAYS_OTHER_RECV),
                "inventories": self._series(self.ROW_DAYS_INVENTORIES),
                "accounts_payable": self._series(self.ROW_DAYS_AP),
                "other_liabilities": self._series(self.ROW_DAYS_OTHER_LIAB),
                "lease_liabilities": self._series(self.ROW_DAYS_LEASE_LIAB),
                "salaries": self._series(self.ROW_DAYS_SALARIES),
                "taxes": self._series(self.ROW_DAYS_TAXES),
                "income_tax": self._series(self.ROW_DAYS_INCOME_TAX),
                "contract_liabilities": self._series(self.ROW_DAYS_CONTRACT_LIAB),
                "provisions": self._series(self.ROW_DAYS_PROVISIONS),
            },
            "non_current": {
                "trade_receivables": self._series(self.ROW_DAYS_NC_TRADE_RECV),
                "other_receivables": self._series(self.ROW_DAYS_NC_OTHER_RECV),
                "accounts_payable": self._series(self.ROW_DAYS_NC_AP),
                "other_liabilities": self._series(self.ROW_DAYS_NC_OTHER_LIAB),
                "lease_liabilities": self._series(self.ROW_DAYS_NC_LEASE_LIAB),
                "salaries": self._series(self.ROW_DAYS_NC_SALARIES),
                "taxes": self._series(self.ROW_DAYS_NC_TAXES),
                "income_tax": self._series(self.ROW_DAYS_NC_INCOME_TAX),
                "contract_liabilities": self._series(self.ROW_DAYS_NC_CONTRACT_LIAB),
                "provisions": self._series(self.ROW_DAYS_NC_PROVISIONS),
            },
        }

    @property
    def net_working_capital(self): return self._series(self.ROW_NET_WORKING_CAPITAL)

    @property
    def change_in_working_capital(self): return self._series(self.ROW_CHANGE_WC)

    def summary(self) -> dict:
        return {
            "days_in": self.days_in,
            "net_working_capital": self.net_working_capital,
            "change_in_working_capital": self.change_in_working_capital,
        }


# ─────────────────────────────────────────────────────────
# 13. Debt and Interest Schedule  (Rows 682–729)
# ─────────────────────────────────────────────────────────
class DebtAndInterestSchedule(BaseSchedule):
    SCHEDULE_NAME = "Debt and Interest Schedule"
    START_ROW, END_ROW = 682, 729

    # Cash
    ROW_CASH_BEGINNING = 693
    ROW_CASH_CHANGE = 694
    ROW_CASH_ENDING = 695
    ROW_CASH_INT_RATE = 697
    ROW_CASH_INT_INCOME = 698
    ROW_ANNUAL_INT_INCOME = 700

    # Loans
    ROW_LOANS_BEGINNING = 703
    ROW_LOANS_ADDITIONS = 704
    ROW_LOANS_ENDING = 705
    ROW_LOANS_INT_RATE = 707
    ROW_LOANS_INT_EXPENSE = 708

    # Revolver
    ROW_OPER_CF = 711
    ROW_INVEST_CF = 712
    ROW_FIN_CF_EX_REVOLVER = 713
    ROW_STOCK_ISSUANCE = 714
    ROW_DIVIDENDS = 715
    ROW_FCF_AFTER_DEBT = 716
    ROW_REVOLVER_BEGINNING = 718
    ROW_REVOLVER_CHANGE = 719
    ROW_REVOLVER_ENDING = 720
    ROW_REVOLVER_INT_RATE = 722
    ROW_REVOLVER_INT_EXPENSE = 723

    # Totals
    ROW_ST_LOANS_REVOLVER = 725
    ROW_LT_LOANS_REVOLVER = 726
    ROW_TOTAL_LOANS_REVOLVER = 727
    ROW_TOTAL_INT_EXPENSE = 728

    @property
    def cash(self):
        return {
            "beginning": self._series(self.ROW_CASH_BEGINNING),
            "change": self._series(self.ROW_CASH_CHANGE),
            "ending": self._series(self.ROW_CASH_ENDING),
            "interest_rate": self._series(self.ROW_CASH_INT_RATE),
            "interest_income": self._series(self.ROW_CASH_INT_INCOME),
            "annual_interest_income": self._series(self.ROW_ANNUAL_INT_INCOME),
        }

    @property
    def loans(self):
        return {
            "beginning": self._series(self.ROW_LOANS_BEGINNING),
            "additions_repayments": self._series(self.ROW_LOANS_ADDITIONS),
            "ending": self._series(self.ROW_LOANS_ENDING),
            "interest_rate": self._series(self.ROW_LOANS_INT_RATE),
            "interest_expense": self._series(self.ROW_LOANS_INT_EXPENSE),
        }

    @property
    def revolver(self):
        return {
            "operating_cf": self._series(self.ROW_OPER_CF),
            "investing_cf": self._series(self.ROW_INVEST_CF),
            "financing_cf_ex_revolver": self._series(self.ROW_FIN_CF_EX_REVOLVER),
            "fcf_after_debt": self._series(self.ROW_FCF_AFTER_DEBT),
            "beginning": self._series(self.ROW_REVOLVER_BEGINNING),
            "change": self._series(self.ROW_REVOLVER_CHANGE),
            "ending": self._series(self.ROW_REVOLVER_ENDING),
            "interest_rate": self._series(self.ROW_REVOLVER_INT_RATE),
            "interest_expense": self._series(self.ROW_REVOLVER_INT_EXPENSE),
        }

    @property
    def totals(self):
        return {
            "st_loans_revolver": self._series(self.ROW_ST_LOANS_REVOLVER),
            "lt_loans_revolver": self._series(self.ROW_LT_LOANS_REVOLVER),
            "total_loans_revolver": self._series(self.ROW_TOTAL_LOANS_REVOLVER),
            "total_interest_expense": self._series(self.ROW_TOTAL_INT_EXPENSE),
        }

    def summary(self) -> dict:
        return {
            "cash": self.cash, "loans": self.loans,
            "revolver": self.revolver, "totals": self.totals,
        }


# ─────────────────────────────────────────────────────────
# 14. Shareholders' Equity Schedule  (Rows 732–764)
# ─────────────────────────────────────────────────────────
class ShareholdersEquitySchedule(BaseSchedule):
    SCHEDULE_NAME = "Shareholders' Equity Schedule"
    START_ROW, END_ROW = 732, 764

    ROW_SHARES_BEGINNING = 741
    ROW_CLASS_A = 742
    ROW_CLASS_B = 743
    ROW_CLASS_C = 744
    ROW_CLASS_D = 745
    ROW_TOTAL_OUTSTANDING = 746
    ROW_NEW_SHARES = 747
    ROW_BUYBACKS = 748
    ROW_SHARES_ENDING = 749
    ROW_GROWTH_YOY = 750
    ROW_SHARE_PRICE = 752
    ROW_DIVIDEND_PAYOUT = 754
    ROW_NET_INCOME = 755
    ROW_COMMON_DIVIDEND = 756

    # Retained Earnings
    ROW_RE_BEGINNING = 760
    ROW_RE_NET_INCOME = 761
    ROW_RE_DIVIDEND = 762
    ROW_RE_ENDING = 763

    @property
    def common_shares(self):
        return {
            "beginning": self._series(self.ROW_SHARES_BEGINNING),
            "class_a": self._series(self.ROW_CLASS_A),
            "class_b": self._series(self.ROW_CLASS_B),
            "class_c": self._series(self.ROW_CLASS_C),
            "class_d": self._series(self.ROW_CLASS_D),
            "total_outstanding": self._series(self.ROW_TOTAL_OUTSTANDING),
            "new_shares": self._series(self.ROW_NEW_SHARES),
            "buybacks": self._series(self.ROW_BUYBACKS),
            "ending": self._series(self.ROW_SHARES_ENDING),
            "growth_yoy": self._series(self.ROW_GROWTH_YOY),
            "share_price": self._series(self.ROW_SHARE_PRICE),
        }

    @property
    def dividends(self):
        return {
            "payout_rate": self._series(self.ROW_DIVIDEND_PAYOUT),
            "net_income": self._series(self.ROW_NET_INCOME),
            "common_dividend": self._series(self.ROW_COMMON_DIVIDEND),
        }

    @property
    def retained_earnings(self):
        return {
            "beginning": self._series(self.ROW_RE_BEGINNING),
            "net_income": self._series(self.ROW_RE_NET_INCOME),
            "dividend": self._series(self.ROW_RE_DIVIDEND),
            "ending": self._series(self.ROW_RE_ENDING),
        }

    def summary(self) -> dict:
        return {
            "common_shares": self.common_shares,
            "dividends": self.dividends,
            "retained_earnings": self.retained_earnings,
        }
