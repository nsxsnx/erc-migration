"Accounts sheet of the resulting workbook"


from dataclasses import dataclass
import re

from lib.datatypes import MonthYear
from lib.osvfile import OsvAddressRecord
from results.resultrow import ResultRow

STREET_ADDRESS_REGEXP = (
    r"ул (?P<street>.*), д\.(?P<house>\d+), кв\.(?P<apartment>\d+).*"
)


@dataclass
class StreetAddress:
    "Disassembled address string"
    street: str
    house: str
    apartment: str


class AccountsResultRow(ResultRow):
    "Row of Accounts sheet"

    def _get_address(self, address: str) -> StreetAddress:
        match = re.match(STREET_ADDRESS_REGEXP, address)
        if not match:
            raise ValueError(f"Can't understand address: {address}")
        address_dict = match.groupdict()
        return StreetAddress(**address_dict)

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
    ) -> None:
        super().__init__(max_fields=32)
        self.set_field(0, date.first_day)
        self.set_field(1, data.account)
        self.set_field(4, data.name)
        self.set_field(7, data.type)
        self.set_field(8, data.address)
        address = self._get_address(data.address)
        self.set_field(12, address.street)
        self.set_field(13, address.house)
        self.set_field(17, address.apartment)
        self.set_field(19, data.area)
        self.set_field(24, data.population)
        self.set_field(25, data.population)
        self.set_field(26, "Общий")
