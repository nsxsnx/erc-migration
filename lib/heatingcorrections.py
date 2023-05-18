"Yearly corrections of heating service"


import calendar
from dataclasses import dataclass
from decimal import Decimal
from enum import Flag, auto
from lib.datatypes import MonthYear

from lib.helpers import BaseMultisheetWorkBookData


@dataclass
class HeatingCorrectionRecord:
    "Record of the last-year heating corrections Excel table"
    line_num: str
    municipality: str
    street: str
    house: str
    building: str
    appartment: str
    account: str
    account_status: str
    square: str
    jan: float
    feb: float
    mar: float
    apr: float
    may: float
    jun: float
    jul: float
    aug: float
    sep: float
    oct: float
    nov: float
    dec: float
    total: float
    year_correction: Decimal
    vkv_jan: float
    vkv_feb: float
    vkv_mar: float
    vkv_apr: float
    vkv_may: float
    vkv_jun: float
    vkv_jul: float
    vkv_aug: float
    vkv_sep: float
    vkv_oct: float
    vkv_nov: float
    vkv_dec: float

    def __post_init__(self):
        if self.year_correction is not None:
            self.year_correction = Decimal(self.year_correction).quantize(
                Decimal("0.01")
            )

    def __iter__(self):
        self._month_index = 0
        return self

    def __next__(self):
        if self._month_index < 12:
            self._month_index += 1
            month_abbrs = [m.lower() for m in calendar.month_abbr if m]
            return getattr(self, month_abbrs[self._month_index - 1])
        raise StopIteration

    def as_month_list(self):
        return [i for i in self]

    def get_by_month_num(self, num: int):
        monthes = [month for month in calendar.month_abbr if month]
        abbrs = {index: month.lower() for index, month in enumerate(monthes, start=1)}
        return getattr(self, abbrs[num])


class HeatingCorrectionsFile(BaseMultisheetWorkBookData):
    "Last-year heating corrections Excel table"


@dataclass
class HeatingVolumesOdpuRecord:
    "Record of the last-year ODPU volumes Excel table"
    line_num: str
    municipality: str
    street: str
    house: str
    building: str
    jan: float
    feb: float
    mar: float
    apr: float
    may: float
    jun: float
    jul: float
    aug: float
    sep: float
    oct: float
    nov: float
    dec: float
    total: float


class HeatingVolumesOdpuFile(BaseMultisheetWorkBookData):
    "Last-year heating ODPU volumes Excel table"


class HeatingPositiveCorrectionType(Flag):
    OPEN_ACCOUNT = 0
    CLOSED_LAST_YEAR = auto()
    CLOSED_CURRENT_YEAR = auto()
    CLOSED_BOTH_YEARS = CLOSED_LAST_YEAR | CLOSED_CURRENT_YEAR


class HeatingPositiveCorrection:
    """
    Heating positive correction data.
    Get's correction sums for current and last years and determines correction type"""

    def __init__(
        self,
        account: str,
        heating_corrections: HeatingCorrectionsFile,
        curent_date: MonthYear,
    ) -> None:
        self.account = account
        self.current_year = curent_date.year
        self.last_year = self.current_year - 1
        # no need to try here, because we would not be here unless last year correction exists
        self.last_year_correction: HeatingCorrectionRecord = (
            heating_corrections.get_account_row(
                self.account,
                f"{self.last_year}",
            )
        )
        self.is_active_current_year = True
        try:
            self.current_year_correction: HeatingCorrectionRecord = (
                heating_corrections.get_account_row(
                    self.account,
                    f"{self.current_year}",
                )
            )
        except ValueError:
            self.is_active_current_year = False
        self.type = HeatingPositiveCorrectionType.OPEN_ACCOUNT
        if not all(self.last_year_correction):
            self.type |= HeatingPositiveCorrectionType.CLOSED_LAST_YEAR
        if not self.is_active_current_year or not all(self.current_year_correction):
            self.type |= HeatingPositiveCorrectionType.CLOSED_CURRENT_YEAR
