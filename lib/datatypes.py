"General data structures"

from __future__ import annotations

import calendar
from dataclasses import dataclass
from enum import StrEnum


@dataclass(frozen=True)
class MonthYear:
    "Immutable helper that stores month/year information"
    month: int
    year: int

    def __lt__(self, other) -> bool:
        return (self.year, self.month) < (other.year, other.month)

    @property
    def previous(self) -> MonthYear:
        "Returns an instance of the previous month"
        if self.month > 1:
            return MonthYear(self.month - 1, self.year)
        else:
            return MonthYear(12, self.year - 1)

    @property
    def next(self) -> MonthYear:
        "Returns an instance of the next month"
        if self.month < 12:
            return MonthYear(self.month + 1, self.year)
        else:
            return MonthYear(1, self.year + 1)

    @property
    def month_abbr(self) -> str:
        "Returns month's abbreviations, eg. 'jan', 'feb', 'mar', etc."
        return calendar.month_abbr[self.month].lower()

    def __str__(self) -> str:
        return f"{self.month:02d}.{self.year}"

    @property
    def first_day(self) -> str:
        "Returns the first day of the month as str"
        return f"01.{str(self)}"


class Service(StrEnum):
    "Services enumerator"
    HEATING = "Отопление"
    GVS = "Тепловая энергия для подогрева воды"
    GVS_ELEVATED = "Тепловая энергия для подогрева воды (повышенный %)"
