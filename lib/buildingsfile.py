"""Buildings details table"""

import logging
import re
from dataclasses import dataclass
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
    correction_month: int
    tariff_first: float
    tariff_second: float
    coefficient: float


class BuildingsFile(BaseMultisheetWorkBookData):
    "Table with buildings addresses"

    def _reg_match_address(self, address: str) -> dict[str, str]:
        match = re.match(ADDRESS_REGEXP, address)
        if not match:
            raise ValueError(f"Can't understand address: {address}")
        return match.groupdict()

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

    def get_tariff(self, address: str, date: MonthYear):
        "Returns tariff for a given address on a given date"
        row: BuildingRecord = self.get_address_row(address, str(date.year))
        tariff = row.tariff_first if date.month < 7 else row.tariff_second
        return tariff * row.coefficient
