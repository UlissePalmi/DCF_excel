"""
Company industry type registry and detection.

Maps ticker symbols to industry types, determining which schedules to instantiate.
"""

from enum import IntEnum


class IndustryType(IntEnum):
    """Company industry classification."""
    GENERIC = 0
    OIL_AND_GAS = 1


# Mapping of known ticker symbols to their industry type
TICKER_INDUSTRY_MAP: dict[str, IndustryType] = {
    # Oil & Gas companies
    "YPF": IndustryType.OIL_AND_GAS,
    "XOM": IndustryType.OIL_AND_GAS,
    "CVX": IndustryType.OIL_AND_GAS,
    "COP": IndustryType.OIL_AND_GAS,
    "MPC": IndustryType.OIL_AND_GAS,
    "PSX": IndustryType.OIL_AND_GAS,
    "VLO": IndustryType.OIL_AND_GAS,
    "EOG": IndustryType.OIL_AND_GAS,
    "OKE": IndustryType.OIL_AND_GAS,
    "MRO": IndustryType.OIL_AND_GAS,
    "DVN": IndustryType.OIL_AND_GAS,
    "OXY": IndustryType.OIL_AND_GAS,
    "HES": IndustryType.OIL_AND_GAS,
    "PXD": IndustryType.OIL_AND_GAS,
    "SLB": IndustryType.OIL_AND_GAS,
    "HAL": IndustryType.OIL_AND_GAS,
    "BE": IndustryType.OIL_AND_GAS,
    "RIG": IndustryType.OIL_AND_GAS,
}


def get_industry_type(ticker: str) -> IndustryType:
    """
    Determine industry type from ticker symbol.

    Returns IndustryType.GENERIC if ticker is not found in the registry.
    """
    return TICKER_INDUSTRY_MAP.get(ticker.upper(), IndustryType.GENERIC)
