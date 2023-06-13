"Files with details of accounts and gvs accurence"

from dataclasses import dataclass
from functools import lru_cache
from typing import Any, Callable, Self

from lib.datatypes import MonthYear
from lib.exceptions import NoServiceRow
from lib.helpers import BaseWorkBookData, SingletonWithArg

MONTHES_RUS = {
    "Январь": 1,
    "Февраль": 2,
    "Март": 3,
    "Апрель": 4,
    "Май": 5,
    "Июнь": 6,
    "Июль": 7,
    "Август": 8,
    "Сентябрь": 9,
    "Октябрь": 10,
    "Ноябрь": 11,
    "Декабрь": 12,
}


@dataclass
class AccountDetailsRecord:
    "A valuable row of account details table"
    __slots__ = (
        "year",
        "month_str",
        "service",
        "population",
        "opening_balance",
        "accural",
        "reaccural",
        "totalaccural",
        "payment",
        "closing_balance",
        "date",
    )

    year: int
    month_str: str
    service: str
    population: int
    opening_balance: float
    accural: float
    reaccural: float
    totalaccural: float
    payment: float
    closing_balance: float
    # date: MonthYear | None = None

    def __post_init__(self):
        self.date = MonthYear(MONTHES_RUS[self.month_str], self.year)


@dataclass
class GvsDetailsRecord:
    "Records of GVS Excel table"
    __slots__ = (
        "address",
        "account",
        "people_registered",
        "people_living",
        "account_status",
        "counter_id",
        "counter_number",
        "metric_date_previos",
        "metric_date_current",
        "metric_previos",
        "metric_current",
        "metric_difference",
        "consumption_ipu",
        "consumption_normative",
        "consumption_average",
        "rise",
        "consumptions_odn",
        "recalculation_volume",
        "total",
    )

    address: str
    account: str
    people_registered: int
    people_living: int
    account_status: str
    counter_id: str
    counter_number: str
    metric_date_previos: str
    metric_date_current: str
    metric_previos: float
    metric_current: float
    metric_difference: float
    consumption_ipu: float
    consumption_normative: float
    consumption_average: float
    rise: float
    consumptions_odn: float
    recalculation_volume: float
    total: float

    @classmethod
    def get_dummy_instance(cls) -> Self:
        "Returns class instance wih all attributes set to None"
        return cls(*[None] * 19)


class AccountDetailsFileSingleton(BaseWorkBookData, metaclass=SingletonWithArg):
    """
    Singleton Excel table with the details of account accurance
    Singleton is created for each file (account) to avoid expensive reading of *.xlsx file
    """

    def __init__(
        self,
        account,
        filename: str,
        header_row: int,
        filter_func: Callable[[Any], bool] | None = None,
        max_col: int | None = None,
    ) -> None:
        super().__init__(
            filename, header_row, AccountDetailsRecord, filter_func, max_col
        )
        self.account = account

    @lru_cache(maxsize=1)
    def _get_month_service_row(
        self, date: MonthYear, service: str
    ) -> AccountDetailsRecord:
        result = list(
            filter(lambda r: (r.date == date and r.service == service), self.records)
        )
        match len(result):
            case 0:
                raise NoServiceRow
            case 1:
                return result[0]
            case _:
                raise ValueError(
                    "More than one details row found for {service} on {date} in {self.filename}"
                )

    def get_service_month_payment(self, date: MonthYear, service: str) -> float:
        "Returns value of payment for service for a particular month"
        return self._get_month_service_row(date, service).payment

    def get_service_month_closing_balance(self, date: MonthYear, service: str) -> float:
        "Returns closing balance of service for a particular month"
        return self._get_month_service_row(date, service).closing_balance

    def get_service_month_reaccural(self, date: MonthYear, service: str) -> float:
        "Returns accural for service for a particular month"
        return self._get_month_service_row(date, service).reaccural

    def get_service_month_accural(self, date: MonthYear, service: str) -> float:
        "Returns accural for service for a particular month"
        return self._get_month_service_row(date, service).accural

    def get_service_year_accurals(self, year: int, service: str) -> list[float]:
        "Returns all acurances for a given service in a particular year"
        res = []
        for month in range(1, 13):
            try:
                accural = self.get_service_month_accural(
                    MonthYear(month, year), service
                )
                res.append(accural)
            except NoServiceRow:
                res.append(0.00)
        return res

    def get_service_closing_month(self, year: int, servce: str) -> int:
        """
        Returns the number of month in a given year when account was closed.
        or zero if closed in previous year
        or -1 if was not closed
        """
        accurals = self.get_service_year_accurals(year, servce)
        if accurals[11]:
            return -1
        for i in range(10, -1, -1):
            if accurals[i]:
                return i + 1
        return 0


class GvsDetailsFileSingleton(BaseWorkBookData, metaclass=SingletonWithArg):
    """
    Singleton Excel table with the details of GVS accurance
    Singleton is created for each file (month) to avoid expensive reading of *.xlsx file
    """

    def __init__(
        self,
        filename: str,
        header_row: int,
        filter_func: Callable[[Any], bool] | None = None,
        max_col: int | None = None,
    ) -> None:
        super().__init__(filename, header_row, GvsDetailsRecord, filter_func, max_col)

    def get_account_row(self, account: str) -> GvsDetailsRecord:
        "Returns table row with given account"
        return self.get_row_by_field_value("account", account)
