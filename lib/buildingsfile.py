"""Buildings details table"""

from decimal import Decimal
import logging
import re
from dataclasses import dataclass
from functools import cache
from typing import Any, Callable

from lib.datatypes import MonthYear
from lib.exceptions import NoAddressRow
from lib.helpers import BaseMultisheetWorkBookData

ADDRESS_REGEXP = (
    r"(ул|мкр) (?P<street>.*), д\.(?P<house>\d+)"
    r"( К\.(?P<building>\d))?( /(?P<drob_building>\d))?, .*"
)


@dataclass
class BuildingRecord:
    "Record of Buildings table"
    num: int
    municipality: str
    street: str
    house: str
    building: str
    has_odpu: str
    has_heating_average: str
    correction_month: int
    tariff_first: Decimal
    tariff_second: Decimal
    coefficient: Decimal

    def __post_init__(self):
        self.tariff_first = Decimal(self.tariff_first).quantize(Decimal("0.01"))
        self.tariff_second = Decimal(self.tariff_second).quantize(Decimal("0.01"))
        self.coefficient = Decimal(self.coefficient).quantize(Decimal("0.01"))


class BuildingsFile(BaseMultisheetWorkBookData):
    "Table with buildings addresses"

    def __init__(
        self,
        filename: str,
        header_row: int,
        record_class: type,
        tariffs_special: str | None = None,
        filter_func: Callable[[Any], bool] | None = None,
        max_col: int | None = None,
    ) -> None:
        super().__init__(filename, header_row, record_class, filter_func, max_col)
        self.tariff_special: dict[MonthYear, Decimal] = dict()
        if tariffs_special:
            for tariff in tariffs_special.split("|"):
                tariff = tariff.strip()
                if not tariff:
                    continue
                date_str, value_str = tariff.split(":", 2)
                date = MonthYear.from_str(date_str.strip())
                value = Decimal(value_str.strip())
                self.tariff_special[date] = value
                logging.info("Special tariff %s applied for %s", value, date)

    def _reg_match_address(self, address: str) -> dict[str, str]:
        match = re.match(ADDRESS_REGEXP, address)
        if not match:
            raise ValueError(f"Can't understand address: {address}")
        return match.groupdict()

    @cache  # pylint: disable=W1518
    def get_address_row(self, address: str, sheet_name: str) -> BuildingRecord:
        "Finds and returns row data for a given address in a given sheet"
        address_dict = self._reg_match_address(address)
        rows: list[BuildingRecord] = self.as_filtered_list(
            ("street", "house"),
            (address_dict["street"], address_dict["house"]),
            sheet_name,
        )
        if not rows:
            raise NoAddressRow
        if len(rows) > 1:
            logging.warning(
                "Too many rows for '%s' in '%s' of Buildings table. Using the first one.",
                address,
                sheet_name,
            )
        return rows[0]

    @cache  # pylint: disable=W1518
    def get_tariff(
        self,
        address: str,
        date: MonthYear,
        use_reduction_factor: bool = False,
    ) -> Decimal:
        "Returns tariff for a given address on a given date"
        row: BuildingRecord = self.get_address_row(address, str(date.year))
        tariff = row.tariff_first if date.month < 7 else row.tariff_second
        if date in self.tariff_special:
            return self.tariff_special[date]
        return tariff * row.coefficient if use_reduction_factor else tariff
