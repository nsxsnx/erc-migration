"General data structures"

from dataclasses import dataclass
from typing import Self
import calendar


@dataclass(frozen=True)
class MonthYear:
    "Immutable helper that stores month/year information"
    month: int
    year: int

    def __lt__(self, other) -> bool:
        return (self.year, self.month) < (other.year, other.month)

    @property
    def previous(self) -> Self:
        "Returns an instance of the previous month"
        if self.month > 1:
            return MonthYear(self.month - 1, self.year)
        else:
            return MonthYear(12, self.year - 1)

    @property
    def next(self) -> Self:
        "Returns an instance of the next month"
        if self.month < 12:
            return MonthYear(self.month + 1, self.year)
        else:
            return MonthYear(1, self.year + 1)

    @property
    def month_abbr(self) -> str:
        return calendar.month_abbr[self.month].lower()
