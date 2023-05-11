" Classes to search for GVS IPUs changes"

import logging
from dataclasses import dataclass
from typing import Any, Callable, Type

from lib.datatypes import MonthYear
from lib.helpers import BaseWorkBookData
from lib.resultfile import GvsIpuInstallDates, ResultRecordType


@dataclass
class GvsIpuMetric:
    "Dates, numbers and metrics of GVS IPUs to find IPU replacement"
    date: MonthYear
    account: str
    counter_number: str
    metric: float
    metric_date: str
    row: int


@dataclass
class FilledResultRecord:
    "Records of result table after it was saved"
    month: int
    year: int
    account: str
    address: str
    type: str
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


class FilledResultTable(BaseWorkBookData):
    "Represents result table to be opened for the second time"

    def __init__(
        self,
        filename: str,
        header_row: int,
        record_class: Type,
        filter_func: Callable[[Any], bool] | None = None,
        max_col: int | None = None,
    ) -> None:
        super().__init__(filename, header_row, record_class, None, max_col)
        res: list[GvsIpuMetric] = []
        for i, rec in enumerate(self.records):
            if not filter_func(rec):
                continue
            res.append(
                GvsIpuMetric(
                    MonthYear(rec.month, rec.year),
                    rec.account,
                    rec.counter_number,
                    rec.metric,
                    rec.metric_date,
                    i + header_row + 1,
                )
            )
        self.records = res


class IpuReplacementFinder:
    """
    Opens result table and searches for replacemnt
    of IPUs relying on changes of IPU number
    """

    def __init__(self, file_name_full) -> None:
        logging.info("Reading GVS IPU data from results table...")
        self.table = FilledResultTable(
            file_name_full,
            3,
            FilledResultRecord,
            filter_func=lambda s: s.type == ResultRecordType.GVS_ACCURAL.name,
            max_col=22,
        )
        logging.info("Reading GVS IPU data from results table done")

    def find_replacements(self):
        "Actually does the job"
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
                    row_num = current_el.row
                    type_cell = active_sheet[f"U{row_num}"]
                    type_cell.value = "При снятии прибора"
                    if next_el.counter_number:
                        row_num = next_el.row
                        type_cell = active_sheet[f"U{row_num}"]
                        type_cell.value = "При установке"
                    GvsIpuInstallDates[gvs_account] = next_el.metric_date
                    logging.debug(current_el)
                    logging.debug(next_el)
                    total_ipu_replacements += 1
                if gvs_account in GvsIpuInstallDates:
                    row_num = next_el.row
                    date_cell = active_sheet[f"K{row_num}"]
                    date_cell.value = GvsIpuInstallDates[gvs_account]
        logging.info(
            "Total additional IPU replacements found: %s", total_ipu_replacements
        )

    def save(self):
        "Saves opened table"
        logging.info("Saving results...")
        self.table.save()
        self.table.close()
        logging.info("Saving results done")
