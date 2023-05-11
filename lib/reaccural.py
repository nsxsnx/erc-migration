"Reaccural calculations"

from dataclasses import dataclass
from decimal import Decimal
from enum import Enum, auto

from lib.datatypes import MonthYear
from lib.detailsfile import AccountDetailsFileSingleton
from lib.exceptions import NoServiceRow

MAX_DEPTH = 36


@dataclass
class ReaccuralMonthRec:
    "Reaccural consists of records of this class"
    date: MonthYear
    sum: float


class ReaccuralType(Enum):
    "Types of reaccural calculation"
    IPU = auto()
    AVERAGE = auto()
    NORMATIVE = auto()


class Reaccural:
    "Reaccural data for GVS service"
    records: list[ReaccuralMonthRec]
    date: MonthYear
    totalsum: Decimal
    valid: bool

    def try_decompose_to_zero(self) -> None:
        """
        First method of calculating reaccurance value for each month.
        Sum of reaccurance and accurance of N previous months must be zero.
        """
        date = self.date
        floating_sum = self.totalsum
        rec = []
        for _ in range(MAX_DEPTH):
            try:
                prev_accural = Decimal(
                    self.account_data.get_service_month_accural(
                        date := date.previous, "Тепловая энергия для подогрева воды"
                    )
                ).quantize(Decimal("0.01"))
            except NoServiceRow:
                continue
            rec.append(ReaccuralMonthRec(date, float(prev_accural)))
            floating_sum += prev_accural
            if not floating_sum:
                self.records = rec[::-1]
                self.valid = True
                return

    def try_decompose_to_previous_accurance(self) -> None:
        """
        Second method of calculating reaccurance value for each month.
        Reminder of the addition of N previous months must be less then
        accurance of the next previous month.
        """
        date = self.date
        floating_sum = self.totalsum
        rec = []
        for _ in range(MAX_DEPTH):
            try:
                prev_accural = Decimal(
                    self.account_data.get_service_month_accural(
                        date := date.previous, "Тепловая энергия для подогрева воды"
                    )
                ).quantize(Decimal("0.01"))
                second_prev_accural = Decimal(
                    self.account_data.get_service_month_accural(
                        date.previous, "Тепловая энергия для подогрева воды"
                    )
                ).quantize(Decimal("0.01"))
            except NoServiceRow:
                continue
            rec.append(ReaccuralMonthRec(date, float(prev_accural)))
            floating_sum += prev_accural
            if abs(floating_sum) < second_prev_accural:
                rec.append(ReaccuralMonthRec(date.previous, abs(float(floating_sum))))
                self.records = rec[::-1]
                self.valid = True
                return

    def set_type(self, p_type: ReaccuralType) -> None:
        "Reaccural type setter"
        self.type = p_type

    def __init__(
        self,
        account_details: AccountDetailsFileSingleton,
        reaccural_date: MonthYear,
        reaccural_sum: float,
    ) -> None:
        self.date = reaccural_date
        self.totalsum = Decimal(reaccural_sum).quantize(Decimal("0.01"))
        self.valid = False
        self.records = []
        self.account_data = account_details
        self.type = None
        self.try_decompose_to_zero()
        if self.valid:
            return
        try:
            next_month_accural = Decimal(
                self.account_data.get_service_month_accural(
                    self.date.next, "Тепловая энергия для подогрева воды"
                )
            )
            if not next_month_accural:
                # account is closed, try another algorithm
                self.try_decompose_to_previous_accurance()
        except NoServiceRow:
            # another sign of a closed account
            self.try_decompose_to_previous_accurance()
        if not self.valid:
            self.records.append(ReaccuralMonthRec(reaccural_date, reaccural_sum))
