"""
Base schedule class that all individual schedule classes inherit from.
"""

from multi_sheet_loader import MultiSheetLoader


class BaseSchedule:
    """
    Base class for all schedules in the YPF DCF Model.

    Each subclass defines SHEET_NAME and calls self._field(key) to retrieve
    {year: value} series from the corresponding sheet in the MultiSheetLoader.
    """

    SCHEDULE_NAME: str = "Base Schedule"
    SHEET_NAME: str = ""   # set in each subclass (usually == SCHEDULE_NAME)

    def __init__(self, loader: MultiSheetLoader):
        self.loader = loader
        self.years = loader.ALL_YEARS
        self.historical_years = loader.HISTORICAL_YEARS
        self.projected_years = loader.PROJECTED_YEARS

    def _field(self, key: str) -> dict[int, float]:
        """Return {year: value} for the given field key from this schedule's sheet."""
        return self.loader.field(self.SHEET_NAME, key)

    def summary(self) -> dict:
        """Override in subclasses to return a structured dict of all schedule data."""
        raise NotImplementedError

    def __repr__(self):
        return f"<{self.__class__.__name__}: {self.SCHEDULE_NAME}>"
