"Table to store results"

from decimal import Decimal
import logging
import os
import shutil
from enum import Enum
from typing import Any

from openpyxl import load_workbook

from lib.buildingsfile import BuildingRecord, BuildingsFile
from lib.datatypes import MonthYear
from lib.detailsfile import AccountDetailsFileSingleton, GvsDetailsRecord
from lib.exceptions import NoServiceRow, ZeroDataResultRow
from lib.helpers import BaseWorkBook
from lib.osvfile import OsvAccuralRecord, OsvAddressRecord
from lib.reaccural import ReaccuralType

GvsIpuInstallDates: dict[str, str] = {}


class ResultRecordType(Enum):
    "Types of rows in result file"
    HEATING_ACCURAL = 1
    HEATING_REACCURAL = 2
    GVS_ACCURAL = 3
    GVS_ELEVATED = 4
    GVS_REACCURAL = 5
    GVS_REACCURAL_ELEVATED = 6
    HEATING_CORRECTION = 7
    HEATING_CORRECTION_ZERO = 8
    HEATING_POSITIVE_CORRECTION = 9
    HEATING_POSITIVE_CORRECTION_EXCESSIVE_REACCURAL = 10


class BaseResultRow:
    "Base class for a row of result table"
    MAX_FIELDS = 47

    def set_field(self, ind: int, value: str | int | float | None = None):
        "Field setter by field number"
        setattr(self, f"f{ind:02d}", value)

    def get_field(self, ind: int) -> str | None:
        "Field getter by field number"
        return getattr(self, f"f{ind:02d}")

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
        buildings: BuildingsFile,
        use_reduction_factor: bool = False,
    ) -> None:
        for ind in range(self.MAX_FIELDS):
            setattr(self, f"f{ind:02d}", None)
        self.set_field(0, date.month)
        self.set_field(1, date.year)
        self.set_field(2, data.account)
        self.set_field(3, data.address)
        self.set_field(6, date.month)
        self.set_field(7, date.year)
        self.price = buildings.get_tariff(data.address, date, use_reduction_factor)
        self.set_field(8, self.price)

    def as_list(self) -> list[Any]:
        "Returns list of all fields"
        result = []
        for ind in range(self.MAX_FIELDS):
            result.append(getattr(self, f"f{ind:02d}"))
        return result

    def _set_odpu_fields(self) -> None:
        self.set_field(9, "Общедомовый")
        self.set_field(10, "01.01.2018")
        self.set_field(11, "Подвал")
        self.set_field(12, 1)
        self.set_field(13, "ВКТ-5")
        self.set_field(14, 1)
        self.set_field(15, 6)
        self.set_field(16, 3)


class HeatingResultRow(BaseResultRow):
    "Result row for Heating service"

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
        accural: OsvAccuralRecord,
        has_odpu: str,
        account_details: AccountDetailsFileSingleton,
        buildings: BuildingsFile,
    ) -> None:
        super().__init__(date, data, buildings)
        self.set_field(4, ResultRecordType.HEATING_ACCURAL.name)
        self.set_field(5, "Отопление")
        # chapter 2:
        if has_odpu:
            self._set_odpu_fields()
        # chapter 3:
        building: BuildingRecord = buildings.get_address_row(
            data.address, str(date.year)
        )
        has_heating_average = building.has_heating_average
        if has_odpu and has_heating_average:
            quantity = f"{float(accural.heating) / self.price:.4f}".replace(".", ",")
            quantity_average = quantity
            sum_average = accural.heating
            self.set_field(26, quantity_average)
            self.set_field(27, sum_average)
            self.set_field(28, sum_average)
        else:
            # chapter 4:
            self.set_field(30, data.population)
            quantity = f"{accural.heating / self.price:.4f}".replace(".", ",")
            quantity_normative = quantity
            sum_normative = accural.heating
            self.set_field(31, quantity_normative)
            self.set_field(32, sum_normative)
            self.set_field(33, sum_normative)
        # chapter 5:
        self.set_field(35, quantity)
        self.set_field(36, accural.heating)
        self.set_field(37, accural.heating)
        # chapter 6:
        payment_sum = account_details.get_service_month_payment(date, "Отопление")
        if payment_sum != 0:
            self.set_field(40, f"20.{date.month:02d}.{date.year}")
            self.set_field(41, f"20.{date.month:02d}.{date.year}")
            self.set_field(42, payment_sum)
            self.set_field(43, "Оплата" if payment_sum >= 0 else "Возврат оплаты")
        # chapter 7:
        self.set_field(
            45, account_details.get_service_month_closing_balance(date, "Отопление")
        )


class HeatingReaccuralResultRow(BaseResultRow):
    "Result row for heating reaccural"

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
        buildings: BuildingsFile,
        has_odpu: str,
        accural_sum: float,
        service,
    ) -> None:
        super().__init__(date, data, buildings)
        self.set_field(4, ResultRecordType.HEATING_REACCURAL.name)
        self.set_field(5, service)
        self.set_field(6, date.previous.month)
        self.set_field(7, date.previous.year)
        if has_odpu:
            self._set_odpu_fields()
        # chapter 3:
        building: BuildingRecord = buildings.get_address_row(
            data.address, str(date.year)
        )
        has_heating_average = building.has_heating_average
        if has_odpu and has_heating_average:
            quantity = f"{float(accural_sum) / self.price:.4f}".replace(".", ",")
            quantity_average = quantity
            sum_average = accural_sum
            self.set_field(26, quantity_average)
            self.set_field(27, sum_average)
            self.set_field(28, sum_average)
        else:
            # chapter 4:
            self.set_field(30, data.population)
            quantity = f"{accural_sum / self.price:.4f}".replace(".", ",")
            quantity_normative = quantity
            sum_normative = accural_sum
            self.set_field(31, quantity_normative)
            self.set_field(32, sum_normative)
            self.set_field(33, sum_normative)
        # chapter 5:
        self.set_field(35, quantity)
        self.set_field(36, accural_sum)
        self.set_field(37, accural_sum)


class GvsSingleResultRow(BaseResultRow):
    "Result row for GVS service for cases where there is only one GVS details record"

    @staticmethod
    def _get_new_counter_number(seed: str):
        return f"{seed}_2"

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
        accural: OsvAccuralRecord,
        account_details: AccountDetailsFileSingleton,
        gvs_details_row: GvsDetailsRecord,
        buildings: BuildingsFile,
        service: str,
    ) -> None:
        super().__init__(date, data, buildings, use_reduction_factor=True)
        self.set_field(4, ResultRecordType.GVS_ACCURAL.name)
        self.set_field(5, service)
        # chapter 3:
        gvs = gvs_details_row
        if gvs.counter_id or gvs.counter_number:
            self.set_field(9, "Индивидуальный")
            self.set_field(10, GvsIpuInstallDates.get(data.account, "01.01.2019"))
            self.set_field(13, "СГВ-15")
            if not gvs.counter_number:
                gvs.counter_number = self._get_new_counter_number(gvs.counter_id)
            self.set_field(14, gvs.counter_number)
            self.set_field(15, 6)
            self.set_field(16, 3)
            # chapter 4:
            if gvs.metric_current is not None:
                self.set_field(19, gvs.metric_date_current)
                self.set_field(20, "От абонента (прочие)")
                self.set_field(21, gvs.metric_current)
            self.set_field(22, gvs.consumption_ipu)
        # chapter 5:
        quantity = f"{accural.gvs/self.price:.4f}".replace(".", ",")
        if gvs.consumption_ipu:
            self.set_field(23, quantity)
            self.set_field(24, accural.gvs)
            self.set_field(25, accural.gvs)
        # chapter 6:
        if gvs.consumption_average:
            self.set_field(26, quantity)
            self.set_field(27, accural.gvs)
            self.set_field(28, accural.gvs)
        # chapter 7:
        if gvs.consumption_normative:
            self.set_field(30, gvs.people_registered)
            self.set_field(31, quantity)
            self.set_field(32, accural.gvs)
            self.set_field(33, accural.gvs)
        # chapter 8:
        self.set_field(35, quantity)
        self.set_field(36, accural.gvs)
        self.set_field(37, accural.gvs)
        # chapter 9:
        try:
            payment_sum = account_details.get_service_month_payment(date, service)
        except NoServiceRow:
            payment_sum = 0
        if payment_sum != 0:
            self.set_field(40, f"20.{date.month:02d}.{date.year}")
            self.set_field(41, f"20.{date.month:02d}.{date.year}")
            self.set_field(42, payment_sum)
            self.set_field(43, "Оплата" if payment_sum >= 0 else "Возврат оплаты")
        # chapter 10:
        try:
            closing_balance = account_details.get_service_month_closing_balance(
                date, service
            )
        except NoServiceRow:
            closing_balance = 0
        self.set_field(45, closing_balance)


class GvsMultipleResultFirstRow(GvsSingleResultRow):
    """
    Result row for GVS service for cases where there are two GVS details records.
    The first of two such rows
    """

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
        accural: OsvAccuralRecord,
        account_details: AccountDetailsFileSingleton,
        gvs_details_row: GvsDetailsRecord,
        buildings: BuildingsFile,
        service: str,
    ) -> None:
        super().__init__(
            date, data, accural, account_details, gvs_details_row, buildings, service
        )
        gvs = gvs_details_row
        if gvs.metric_current is not None:
            self.set_field(20, "При снятии прибора")


class GvsMultipleResultSecondRow(GvsSingleResultRow):
    """
    Result row for GVS service for cases where there are two GVS details records.
    The second of two such rows
    """

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
        accural: OsvAccuralRecord,
        account_details: AccountDetailsFileSingleton,
        gvs_details_row: GvsDetailsRecord,
        buildings: BuildingsFile,
        service: str,
    ) -> None:
        super().__init__(
            date, data, accural, account_details, gvs_details_row, buildings, service
        )
        gvs = gvs_details_row
        self.set_field(10, gvs.metric_date_current)
        GvsIpuInstallDates[gvs.account] = gvs.metric_date_current
        if gvs.metric_current is not None:
            self.set_field(20, "При установке")
        for i in range(23, 46):
            self.set_field(i, None)


class GvsReaccuralResultRow(BaseResultRow):
    "Result row for GVS reaccural"

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
        gvs_details_row: GvsDetailsRecord,
        reaccural_date: MonthYear,
        reaccural_sum: float,
        reaccural_type: ReaccuralType,
        buildings: BuildingsFile,
        service,
        record_type,
    ) -> None:
        super().__init__(date, data, buildings, use_reduction_factor=True)
        self.set_field(4, record_type.name)
        self.set_field(5, service)
        # chapter 2:
        self.set_field(6, reaccural_date.month)
        self.set_field(7, reaccural_date.year)
        self.set_field(8, self.price)
        # chapter 3: same as GvsSingleResultRow
        gvs = gvs_details_row
        if gvs.counter_id or gvs.counter_number:
            self.set_field(9, "Индивидуальный")
            self.set_field(10, GvsIpuInstallDates.get(data.account, "01.01.2019"))
            self.set_field(13, "СГВ-15")
            if not gvs.counter_number:
                gvs.counter_number = GvsSingleResultRow._get_new_counter_number(
                    gvs.counter_id
                )
            self.set_field(14, gvs.counter_number)
            self.set_field(15, 6)
            self.set_field(16, 3)
        quantity = f"{reaccural_sum/self.price:.4f}".replace(".", ",")
        # chapter 5: same as chapter 7 of GvsSingleResultRow
        match reaccural_type:
            case ReaccuralType.IPU:
                self.set_field(23, quantity)
                self.set_field(24, reaccural_sum)
                self.set_field(25, reaccural_sum)
            case ReaccuralType.AVERAGE:
                self.set_field(26, quantity)
                self.set_field(27, reaccural_sum)
                self.set_field(28, reaccural_sum)
            case ReaccuralType.NORMATIVE:
                self.set_field(31, quantity)
                self.set_field(32, reaccural_sum)
                self.set_field(33, reaccural_sum)
            case _:
                raise ValueError
        # chapter 6: same as chapter 8 of GvsSingleResultRow
        self.set_field(35, quantity)
        self.set_field(36, reaccural_sum)
        self.set_field(37, reaccural_sum)


class GvsElevatedResultRow(GvsSingleResultRow):
    "Result row for GVS elevated percent accural"

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
        accural: OsvAccuralRecord,
        account_details: AccountDetailsFileSingleton,
        gvs_details_row: GvsDetailsRecord,
        buildings: BuildingsFile,
        service: str,
    ) -> None:
        super().__init__(
            date, data, accural, account_details, gvs_details_row, buildings, service
        )
        self.set_field(4, ResultRecordType.GVS_ELEVATED.name)
        self.set_field(5, service)
        # chapter 3:
        gvs = gvs_details_row
        # chapter 4:
        self.set_field(19, None)
        self.set_field(20, None)
        self.set_field(21, None)
        self.set_field(22, None)
        # chapter 5:
        try:
            accural_sum = account_details.get_service_month_accural(date, service)
        except NoServiceRow:
            accural_sum = 0
        quantity = f"{accural_sum/self.price:.4f}".replace(".", ",")
        if gvs.consumption_ipu:
            self.set_field(23, quantity)
            self.set_field(24, accural_sum)
            self.set_field(25, accural_sum)
        # chapter 6:
        if gvs.consumption_average:
            self.set_field(26, quantity)
            self.set_field(27, accural_sum)
            self.set_field(28, accural_sum)
        # chapter 7:
        if gvs.consumption_normative:
            self.set_field(30, gvs.people_registered)
            self.set_field(31, quantity)
            self.set_field(32, accural_sum)
            self.set_field(33, accural_sum)
        # chapter 8:
        self.set_field(35, quantity)
        self.set_field(36, accural_sum)
        self.set_field(37, accural_sum)
        if not any(
            (
                accural_sum,
                self.get_field(42),  # payment_sum
                self.get_field(45),  # closing_balance
            )
        ):
            raise ZeroDataResultRow


class HeatingCorrectionResultRow(BaseResultRow):
    "Result row for heating last-year correction"
    rounding_error: list = [0.0]

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
        correction_date: MonthYear,
        correction_sum: float,
        correction_volume: float,
        odpu_volume: float,
        service: str,
        buildings: BuildingsFile,
    ) -> None:
        super().__init__(date, data, buildings)
        self.set_field(4, ResultRecordType.HEATING_CORRECTION.name)
        self.set_field(5, service)
        self.set_field(6, correction_date.month)
        self.set_field(7, correction_date.year)
        self.price = buildings.get_tariff(data.address, correction_date)
        self.set_field(8, self.price)
        self._set_odpu_fields()
        self.set_field(19, f"23.{correction_date.month:02d}.{correction_date.year}")
        self.set_field(20, "Контрольное")
        self.set_field(22, odpu_volume)
        accural_sum = correction_volume * self.price - correction_sum
        accural_sum_rounded = round(accural_sum, 2)
        # loosing some penny because of rounding
        # compensate for it:
        self.rounding_error[0] += accural_sum - accural_sum_rounded
        HALF_PENNY = 0.005  # pylint: disable=C0103
        if abs(self.rounding_error[0]) >= HALF_PENNY:
            SIGNED_PENNY = round(self.rounding_error[0], 2)  # pylint: disable=C0103
            accural_sum_rounded += SIGNED_PENNY
            self.rounding_error[0] -= SIGNED_PENNY
        accural_sum = accural_sum_rounded
        self.set_field(24, accural_sum)
        self.set_field(25, accural_sum)
        self.set_field(36, accural_sum)
        self.set_field(37, accural_sum)


class HeatingNegativeCorrectionZeroResultRow(BaseResultRow):
    "Result row for heating last-year correction closing balance only record"

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
        correction_date: MonthYear,
        account_details: AccountDetailsFileSingleton,
        service: str,
        buildings: BuildingsFile,
    ) -> None:
        super().__init__(date, data, buildings)
        self.set_field(4, ResultRecordType.HEATING_CORRECTION_ZERO.name)
        self.set_field(5, service)
        self.set_field(6, correction_date.month)
        self.set_field(7, correction_date.year)
        self.price = buildings.get_tariff(data.address, correction_date)
        self.set_field(8, self.price)
        self._set_odpu_fields()
        self.set_field(
            45,
            account_details.get_service_month_closing_balance(date, service),
        )


class HeatingPositiveCorrectionResultRow(BaseResultRow):
    "Result row for heating last-year correction closing balance only record"

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
        correction_date: MonthYear,
        service: str,
        future_installment: Decimal | None,
        total_closing_balance: Decimal,
        total_future_installment: Decimal,
        buildings: BuildingsFile,
    ) -> None:
        super().__init__(date, data, buildings)
        self.set_field(0, correction_date.month)
        self.set_field(4, ResultRecordType.HEATING_POSITIVE_CORRECTION.name)
        self.set_field(5, service)
        self.set_field(6, correction_date.month)
        self.set_field(7, correction_date.year)
        self.price = Decimal(buildings.get_tariff(data.address, correction_date))
        self.set_field(8, self.price)
        self._set_odpu_fields()
        self.set_field(19, f"31.12.{correction_date.year}")
        self.set_field(20, "Контрольное")
        self.set_field(39, future_installment)
        self.set_field(45, total_closing_balance)
        self.set_field(46, total_future_installment)


class HeatingPositiveCorrectionExcessiveReaccuralResultRow(BaseResultRow):
    "Result row for reaccural that can not be distributed to correction rows"

    def __init__(
        self,
        date: MonthYear,
        data: OsvAddressRecord,
        correction_date: MonthYear,
        accural_sum: Decimal,
        service: str,
        buildings: BuildingsFile,
    ) -> None:
        super().__init__(date, data, buildings)
        self.set_field(
            4, ResultRecordType.HEATING_POSITIVE_CORRECTION_EXCESSIVE_REACCURAL.name
        )
        self.set_field(5, service)
        self.set_field(6, correction_date.month)
        self.set_field(7, correction_date.year)
        self.price = Decimal(buildings.get_tariff(data.address, correction_date))
        self.set_field(8, self.price)
        self._set_odpu_fields()
        quantity = f"{accural_sum / self.price:.4f}".replace(".", ",")
        self.set_field(23, quantity)
        self.set_field(24, accural_sum)
        self.set_field(25, accural_sum)
        self.set_field(35, quantity)
        self.set_field(36, accural_sum)
        self.set_field(37, accural_sum)


class ResultFile(BaseWorkBook):
    """Table of results"""

    def __init__(self, base_dir: str, conf: dict) -> None:
        self.base_dir = base_dir
        self.conf = conf
        file_name, self.sheet_name = conf["result_file"].split("@", 2)
        self.file_name_full = os.path.join(self.base_dir, file_name)
        template_name_full = os.path.join(
            os.path.dirname(self.base_dir), conf["result_template"]
        )
        logging.info("Initialazing result table %s ...", self.file_name_full)
        shutil.copyfile(template_name_full, self.file_name_full)
        self.workbook = load_workbook(filename=self.file_name_full)
        self.sheet = self.workbook[self.sheet_name]
        logging.info("Initialazing result table done")

    def save(self) -> None:
        """Saves result table data to disk"""
        logging.info("Saving results table...")
        self.workbook.save(filename=self.file_name_full)
        logging.info("Saving results table done")

    def add_row(self, row: BaseResultRow):
        "Adds row to table"
        self.sheet.append(row.as_list())
