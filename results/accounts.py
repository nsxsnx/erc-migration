"Accounts list of the resulting workbook"


from lib.datatypes import MonthYear
from lib.osvfile import OsvAddressRecord
from results.resultrow import ResultRow


class AccountsResultRow(ResultRow):
    "Base class for a row of result table"

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
        self.set_field(19, data.area)
        self.set_field(24, data.population)
        self.set_field(25, data.population)
        self.set_field(26, "Общий")
