" Classes to search for GVS IPUs changes"

from dataclasses import dataclass
from typing import Any, Callable, Type

from lib.datatypes import MonthYear
from lib.helpers import BaseWorkBookData


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
