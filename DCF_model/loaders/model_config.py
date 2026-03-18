"""
Model configuration for the DCF model.

Contains global settings for year counts and other model parameters.
"""

# Year counts
HISTORICAL_YEARS = 7
PROJECTED_YEARS = 10

# Start years (used to generate actual year lists)
HIST_START_YEAR = 2019
PROJ_START_YEAR = 2026


def get_historical_years() -> list[int]:
    """Generate list of historical years."""
    return list(range(HIST_START_YEAR, HIST_START_YEAR + HISTORICAL_YEARS))


def get_projected_years() -> list[int]:
    """Generate list of projected years."""
    return list(range(PROJ_START_YEAR, PROJ_START_YEAR + PROJECTED_YEARS))


def get_all_years() -> list[int]:
    """Generate list of all years (historical + projected)."""
    return get_historical_years() + get_projected_years()
