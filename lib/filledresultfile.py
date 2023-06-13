" Classes to search for GVS IPUs changes"

import logging
from collections import Counter
from dataclasses import dataclass, fields
from typing import Any, Callable, Type

from lib.datatypes import MonthYear
from lib.resultfile import GvsIpuInstallDates, ResultFile


@dataclass
class AccountRow:
    "Base class for all typed result rows"
    __slots__ = "row_num", "date", "account"
    row_num: int
    date: MonthYear
    account: str


@dataclass
class GvsIpuMetric(AccountRow):
    "Dates, numbers and metrics of GVS IPUs to find IPU replacement"
    __slots__ = "counter_number", "metric", "metric_date"
    counter_number: str
    metric: float
    metric_date: str


@dataclass
class AccountClosingBalance(AccountRow):
    "Closing balance of a record"
    __slots__ = "type_name", "closing_balance"
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


class FilledTableUpdater:
    """
    Performs some additional data modifications or result records
    """

    def __init__(self, results: ResultFile, max_col: int | None = None) -> None:
        logging.info("Re-reading result rows...")
        record_class = FilledResultRawRecord
        self.changes_counter = Counter()
        if not max_col:
            max_col = len(fields(record_class))
        self.table = results
        sheet = self.table.sheet
        self.records: list[record_class] = list()
        self._records: list = list()
        for row in sheet.iter_rows(  # type: ignore
            min_row=self.table.header_row + 1, max_col=max_col, values_only=True
        ):
            record = record_class(*row)
            self.records.append(record)

    def prepare_records_cache(
        self,
        record_class: Type,
        filter_func: Callable[[Any], bool] | None = None,
    ):
        "Filters raw result records and converts them to typed one for additional processing"
        record_class_fields = [
            field.name for field in fields(record_class)[len(fields(AccountRow)) :]
        ]
        res: list[record_class] = []
        for i, rec in enumerate(self.records):
            if filter_func is not None and not filter_func(rec):
                continue
            res.append(
                record_class(
                    i + self.table.header_row + 1,
                    MonthYear(rec.month, rec.year),
                    rec.account,
                    *[getattr(rec, field_name) for field_name in record_class_fields],
                )
            )
        self._records = res

    def find_gvs_ipu_replacements(self):
        """Finds replacements of IPUs relying on changes of IPU number"""
        logging.info("Looking for IPU replacement...")
        counter_name = "IPU_replacement"
        active_sheet = self.table.workbook.active
        gvs_accounts = sorted({r.account for r in self._records})
        for gvs_account in gvs_accounts:
            counters: list[GvsIpuMetric] = [
                r for r in self._records if r.account == gvs_account
            ]
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
                    self.changes_counter.update([counter_name])
                if gvs_account in GvsIpuInstallDates:
                    row_num = next_el.row_num
                    date_cell = active_sheet[f"K{row_num}"]
                    date_cell.value = GvsIpuInstallDates[gvs_account]

    def decrease_closing_balance(self):
        """
        Closing balance of heating accural record must be decreased by the value
        of corresponding positive correction (installment) record
        """
        logging.info("Decreasing closing balance...")
        counter_name = "Closing_balance_decrease"
        heating_corrections: list[AccountClosingBalance] = [
            r for r in self._records if r.type_name == "HEATING_POSITIVE_CORRECTION"
        ]
        heating_accurals: list[AccountClosingBalance] = [
            r for r in self._records if r.type_name == "HEATING_ACCURAL"
        ]
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
            cell.value = accural_row.closing_balance - float(correction.closing_balance)
            self.changes_counter.update([counter_name])
