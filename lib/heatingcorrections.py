"Yearly corrections of heating service"


from dataclasses import dataclass
from decimal import Decimal

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
