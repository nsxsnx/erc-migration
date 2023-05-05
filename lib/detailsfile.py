"Files with details of accounts and gvs accurence"

from dataclasses import dataclass
from typing import Any, Callable, Type

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
    date: MonthYear | None = None

    def __post_init__(self):
        self.date = MonthYear(MONTHES_RUS[self.month_str], self.year)


@dataclass
class GvsDetailsRecord:
    "Records of GVS Excel table"
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
    def get_dummy_instance(cls):
        "Returns class instance wih all attributes set to None"
        return cls(
            *[
                None,
            ]
            * 19
        )


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
        record_class: Type,
        filter_func: Callable[[Any], bool] | None = None,
        max_col: int | None = None,
    ) -> None:
        super().__init__(filename, header_row, record_class, filter_func, max_col)
        self.account = account

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
                    "More than one payment row found for {service} on {date} in {self.filename}"
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


class GvsDetailsFileSingleton(BaseWorkBookData, metaclass=SingletonWithArg):
    """
    Singleton Excel table with the details of GVS accurance
    Singleton is created for each file (month) to avoid expensive reading of *.xlsx file
    """

    def get_account_row(self, account: str) -> GvsDetailsRecord:
        "Returns table row with given account"
        return self.get_row_by_field_value("account", account)
