" Classes to search for GVS IPUs changes"

import logging
from dataclasses import dataclass, fields
from typing import Any, Callable, Type

from lib.datatypes import MonthYear
from lib.helpers import BaseWorkBookData
from lib.resultfile import GvsIpuInstallDates


@dataclass
class AccountRow:
    "Base class for all typed result rows"
    row_num: int
    date: MonthYear
    account: str


@dataclass
class GvsIpuMetric(AccountRow):
    "Dates, numbers and metrics of GVS IPUs to find IPU replacement"
    counter_number: str
    metric: float
    metric_date: str


@dataclass
class AccountClosingBalance(AccountRow):
    "Closing balance of a record"
    type_name: str
    closing_balance: float


@dataclass
class FilledResultRawRecord:
    "Records of result table after it was saved"
    __slots__ = (
        "month",
        "year",
        "account",
        "address",
        "type_name",
        "service",
        "f06",
        "f07",
        "f08",
        "f09",
        "f10",
        "f11",
        "f12",
        "f13",
        "counter_number",
        "f15",
        "f16",
        "f17",
        "f18",
        "metric_date",
        "f20",
        "metric",
        "f22",
        "f23",
        "f24",
        "f25",
        "f26",
        "f27",
        "f28",
        "f29",
        "f30",
        "f31",
        "f32",
        "f33",
        "f34",
        "f35",
        "f36",
        "f37",
        "f38",
        "f39",
        "f40",
        "f41",
        "f42",
        "f43",
        "f44",
        "closing_balance",
        "f46",
    )

    month: int
    year: int
    account: str
    address: str
    type_name: str
    service: str
    f06: str
    f07: str
    f08: str
    f09: str
    f10: str
    f11: str
    f12: str
    f13: str
    counter_number: str
    f15: str
    f16: str
    f17: str
    f18: str
    metric_date: str
    f20: str
    metric: float
    f22: str
    f23: str
    f24: str
    f25: str
    f26: str
    f27: str
    f28: str
    f29: str
    f30: str
    f31: str
    f32: str
    f33: str
    f34: str
    f35: str
    f36: str
    f37: str
    f38: str
    f39: str
    f40: str
    f41: str
    f42: str
    f43: str
    f44: str
    closing_balance: str
    f46: str


class FilledResultTable(BaseWorkBookData):
    "Represents result table to be opened for the second time"

    def __init__(
        self,
        filename: str,
        header_row: int,
        record_class: Type[Any],
        filter_func: Callable[[Any], bool] | None = None,
        max_col: int | None = None,
    ) -> None:
        self._filter_func = filter_func
        self._header_row = header_row
        if not max_col:
            max_col = len(fields(record_class))
        super().__init__(filename, header_row, record_class, None, max_col)

    def records_to_class(self, record_class: Type):
        "Converts raw result record to typed one"
        record_class_fields = [
            field.name for field in fields(record_class)[len(fields(AccountRow)) :]
        ]
        res: list[record_class] = []
        for i, rec in enumerate(self.records):
            if self._filter_func is not None and not self._filter_func(rec):
                continue
            res.append(
                record_class(
                    i + self._header_row + 1,
                    MonthYear(rec.month, rec.year),
                    rec.account,
                    *[getattr(rec, field_name) for field_name in record_class_fields],
                )
            )
        self.records = res


class FilledTableUpdater:
    """
    Opens result table and performs some additional data modifications
    """

    def __init__(self, file_name_full: str, filter_func: Callable) -> None:
        logging.info("Reading data from result table")
        self.table = FilledResultTable(
            file_name_full,
            3,
            FilledResultRawRecord,
            filter_func=filter_func,
        )
        logging.info("Reading data from result table done")

    def find_gvs_ipu_replacements(self):
        """Finds replacements of IPUs relying on changes of IPU number"""
        logging.info("Looking for IPU replacement...")
        active_sheet = self.table.workbook.active
        gvs_accounts = self.table.get_field_values("account")
        total_ipu_replacements = 0
        for gvs_account in gvs_accounts:
            counters: list[GvsIpuMetric] = self.table.as_filtered_list(
                ("account",), (gvs_account,)
            )
            ignore_next = False
            for i, current_el in enumerate(counters[:-1]):
                next_el = counters[i + 1]
                if ignore_next:
                    ignore_next = False
                    continue
                if current_el.date == next_el.date:
                    ignore_next = True
                    continue
                if current_el.counter_number != next_el.counter_number:
                    row_num = current_el.row_num
                    type_cell = active_sheet[f"U{row_num}"]
                    type_cell.value = "При снятии прибора"
                    if next_el.counter_number:
                        row_num = next_el.row_num
                        type_cell = active_sheet[f"U{row_num}"]
                        type_cell.value = "При установке"
                    GvsIpuInstallDates[gvs_account] = next_el.metric_date
                    logging.debug(current_el)
                    logging.debug(next_el)
                    total_ipu_replacements += 1
                if gvs_account in GvsIpuInstallDates:
                    row_num = next_el.row_num
                    date_cell = active_sheet[f"K{row_num}"]
                    date_cell.value = GvsIpuInstallDates[gvs_account]
        logging.info(
            "Total additional IPU replacements found: %s", total_ipu_replacements
        )

    def decrease_closing_balance(self):
        """
        Closing balance of main heating record must be decreased by the value of
        positive correction (installment) record
        """
        logging.info("Decreasing closing balance...")
        heating_corrections: list[AccountClosingBalance] = self.table.as_filtered_list(
            ("type_name",), ("HEATING_POSITIVE_CORRECTION",)
        )
        heating_accurals: list[AccountClosingBalance] = self.table.as_filtered_list(
            ("type_name",), ("HEATING_ACCURAL",)
        )
        account_accurals: list[AccountClosingBalance] = []
        for correction in heating_corrections:
            if (
                not account_accurals
                or correction.account != account_accurals[0].account
            ):
                account_accurals = [
                    rec for rec in heating_accurals if rec.account == correction.account
                ]
            account_date_accurals: list[AccountClosingBalance] = [
                rec for rec in account_accurals if rec.date == correction.date
            ]
            match len(account_date_accurals):
                case 0:
                    continue
                case 1:
                    pass
                case _:
                    raise ValueError(
                        f"Too many corresponding heating accural found: \
                            {correction.account} {correction.date}"
                    )
            accural_row = account_date_accurals[0]
            cell = self.table.workbook.active[f"AT{accural_row.row_num}"]
            cell.value = accural_row.closing_balance - correction.closing_balance
        logging.info("Decreasing closing balance done")

    def save(self):
        "Saves opened table"
        logging.info("Saving results...")
        self.table.save()
        self.table.close()
        logging.info("Saving results done")
