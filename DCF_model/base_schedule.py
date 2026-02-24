"""
Base schedule class that all individual schedule classes inherit from.
"""

from data_loader import DataLoader


class BaseSchedule:
    """
    Base class for all schedules in the YPF DCF Model.
    
    Each schedule corresponds to a labeled section of the Model sheet.
    Subclasses define ROW_MAP dicts that map descriptive names to Excel row numbers,
    and the DataLoader retrieves the actual numbers.
    """

    SCHEDULE_NAME: str = "Base Schedule"
    START_ROW: int = 1
    END_ROW: int = 1

    def __init__(self, loader: DataLoader):
        self.loader = loader
        self.years = loader.ALL_YEARS
        self.historical_years = loader.HISTORICAL_YEARS
        self.projected_years = loader.PROJECTED_YEARS

    def _series(self, row: int, years=None) -> dict[int, float]:
        return self.loader.get_row_series(row, years)

    def _hist(self, row: int) -> dict[int, float]:
        return self.loader.get_historical(row)

    def _proj(self, row: int) -> dict[int, float]:
        return self.loader.get_projected(row)

    def summary(self) -> dict:
        """Override in subclasses to return a structured dict of all schedule data."""
        raise NotImplementedError

    def __repr__(self):
        return f"<{self.__class__.__name__}: {self.SCHEDULE_NAME}>"
