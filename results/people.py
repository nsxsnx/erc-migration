"People sheet of the resulting workbook"


from lib.datatypes import MonthYear
from lib.osvfile import OsvAddressRecord
from results.resultrow import ResultRow


class PeopleResultRow(ResultRow):
    "Row of People sheet"

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
    ) -> None:
        super().__init__(max_fields=4)
        self.set_field(0, date.first_day)
        self.set_field(1, data.account)
        self.set_field(2, data.address)
        self.set_field(3, data.name)
