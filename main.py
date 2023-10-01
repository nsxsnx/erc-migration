"""Energobilling data calculation/formatting"""


import calendar
import configparser
import logging
import os
import re
from decimal import Decimal
from typing import Mapping

from lib.buildingsfile import BuildingRecord, BuildingsFile
from lib.datatypes import MonthYear, Service
from lib.detailsfile import (
    AccountDetailsFileSingleton,
    GvsDetailsFileSingleton,
    GvsDetailsRecord,
)
from lib.errormessage import ErrorMessageConsoleHandler
from lib.exceptions import NoServiceRow, ZeroDataResultRow
from lib.heatingcorrections import (
    HeatingCorrectionAccountStatus,
    HeatingCorrectionRecord,
    HeatingCorrectionsFile,
    HeatingPositiveCorrection,
    HeatingVolumesOdpuFile,
    HeatingVolumesOdpuRecord,
)
from lib.helpers import BaseWorkBook
from lib.osvfile import (
    OSVDATA_REGEXP,
    OsvFile,
    OsvPath,
    osvdata_regexp_compiled,
)
from lib.reaccural import Reaccural
from results.accounts import AccountsResultRow
from results.people import PeopleResultRow
from results.workbook import ResultWorkBook
from results import ResultSheet
from results.calculations import (
    GvsElevatedResultRow,
    GvsMultipleResultFirstRow,
    GvsMultipleResultSecondRow,
    GvsOpeningBalanceResultRow,
    GvsReaccuralResultRow,
    GvsSingleResultRow,
    HeatingCorrectionResultRow,
    HeatingNegativeCorrectionZeroResultRow,
    HeatingOpeningBalanceResultRow,
    HeatingPositiveCorrectionExcessiveReaccuralResultRow,
    HeatingPositiveCorrectionResultRow,
    HeatingReaccuralResultRow,
    HeatingResultRow,
    CalculationRecordType,
)
from results.filledworkbook import (
    AccountClosingBalance,
    WorkBookDataUpdater,
    GvsIpuMetric,
)

CONFIG_PATH = "./config.ini"


AccountChangebleInfo = tuple[str, str]


class RegionDir:
    "A directory with the data of a particular region"

    account_details: AccountDetailsFileSingleton
    osv: OsvFile
    account: str
    building_record: BuildingRecord
    seen_account_info: dict[str, AccountChangebleInfo]

    def is_config_option_true(self, option_name: str) -> bool:
        "Returns option bool value if it is set in config file"
        if option_name in self.conf:
            return bool(self.conf[option_name])
        return False

    def _is_debugging_account(self, account: str):
        "Returns True if debugging is enabled in config file for a given `account`"
        try:
            if config["DEFAULT"]["account"] == account:
                return True
        except KeyError:
            return True
        return False

    def __init__(self, base_dir: str, conf: Mapping[str, str]) -> None:
        logging.info("Initialazing %s region data...", base_dir)
        self.conf = {k: v.strip() for k, v in conf.items()}
        self.base_dir = base_dir
        self.osv_path = os.path.join(self.base_dir, self.conf["osv.dir"])
        self.error_handler = ErrorMessageConsoleHandler()
        self.seen_account_info = dict()
        self.seen_people_names = dict()
        self.buildings: BuildingsFile = BuildingsFile(
            os.path.join(self.base_dir, conf["file.buildings"]),
            1,
            BuildingRecord,
            conf["tariffs_special"] if "tariffs_special" in conf else None,
        )
        logging.info(
            "Read buildings: %s sheets, %s elements total",
            self.buildings.get_sheets_count(),
            self.buildings.get_strings_count(),
        )
        self.heating_corrections: HeatingCorrectionsFile = HeatingCorrectionsFile(
            os.path.join(self.base_dir, self.conf["file.heating_corrections"]),
            1,
            HeatingCorrectionRecord,
            filter_func=lambda x: x.account,
        )
        logging.info(
            "Read heating corrections: %s sheets, %s elements total",
            self.heating_corrections.get_sheets_count(),
            self.heating_corrections.get_strings_count(),
        )
        self.heating_volumes_odpu: HeatingVolumesOdpuFile = HeatingVolumesOdpuFile(
            os.path.join(self.base_dir, self.conf["file.heating_volumes_odpu"]),
            1,
            HeatingVolumesOdpuRecord,
        )
        logging.info(
            "Read heating ODPU volumes: %s sheets, %s elements total",
            self.heating_volumes_odpu.get_sheets_count(),
            self.heating_volumes_odpu.get_strings_count(),
        )
        osv_files = [
            OsvPath(os.path.join(self.osv_path, f))
            for f in os.listdir(self.osv_path)
            if os.path.isfile(os.path.join(self.osv_path, f)) and not f.startswith(".")
        ]
        self.osv_files: list[OsvPath] = sorted(osv_files)
        self.osv_files = self.osv_files[: int(self.conf["max_osv_files"])]
        _ = [file.validate() for file in self.osv_files]
        self.results = ResultWorkBook(self.base_dir, self.conf)

    def _add_initial_balance_row(self, service):
        if service in self.account_details.seen_opening_balance:
            return
        self.account_details.seen_opening_balance.append(service)
        match service:
            case Service.HEATING:
                row = HeatingOpeningBalanceResultRow(
                    self.osv.date,
                    self.osv.record.address,
                    self.building_record.has_odpu,
                    self.account_details,
                    self.buildings,
                    service,
                )
            case Service.GVS | Service.GVS_ELEVATED:
                gvs_details = GvsDetailsFileSingleton(
                    os.path.join(
                        self.base_dir,
                        self.conf["gvs.dir"],
                        f"{self.osv.date.month:02d}.{self.osv.date.year}.xlsx",
                    ),
                    int(self.conf["gvs_details.header_row"]),
                    lambda x: x.account,
                )
                gvs_details_rows: list[GvsDetailsRecord] = gvs_details.as_filtered_list(
                    ("account",), (self.osv.record.address.account,)
                )
                try:
                    gvs_details_row = gvs_details_rows[0]
                except IndexError:
                    gvs_details_row = None
                row = GvsOpeningBalanceResultRow(
                    self.osv.date,
                    self.osv.record.address,
                    self.account_details,
                    gvs_details_row,
                    self.buildings,
                    service,
                )
            case _:
                raise NotImplementedError
        self.results.calculations.add_row(row)

    def _add_heating(self):
        if not any(
            (
                self.osv.record.accural.heating,
                self.osv.record.accural.reaccural,
                self.osv.record.accural.payment,
            )
        ):
            return
        service = Service.HEATING
        try:
            heating_row = HeatingResultRow(
                self.osv.date,
                self.osv.record.address,
                self.osv.record.accural,
                self.building_record.has_odpu,
                self.account_details,
                self.buildings,
                service,
            )
            self.results.calculations.add_row(heating_row)
        except NoServiceRow:
            pass
        self._add_initial_balance_row(service)

    def _add_gvs(self):
        if not any(
            (
                self.osv.record.accural.heating,
                self.osv.record.accural.gvs,
                self.osv.record.accural.reaccural,
                self.osv.record.accural.payment,
            )
        ):
            return
        service = Service.GVS
        gvs_details = GvsDetailsFileSingleton(
            os.path.join(
                self.base_dir,
                self.conf["gvs.dir"],
                f"{self.osv.date.month:02d}.{self.osv.date.year}.xlsx",
            ),
            int(self.conf["gvs_details.header_row"]),
            lambda x: x.account,
        )
        gvs_details_rows: list[GvsDetailsRecord] = gvs_details.as_filtered_list(
            ("account",), (self.osv.record.address.account,)
        )
        if len(gvs_details_rows) > 2:
            gvs_details_rows = [gvs_details_rows[0], gvs_details_rows[-1]]
            logging.warning(
                "Too many GVS details records for account %s in %s, skipped",
                gvs_details_rows[0].account,
                self.osv.date,
            )
        match len(gvs_details_rows):
            case 0:
                try:
                    closing_balance = (
                        self.account_details.get_service_month_closing_balance(
                            self.osv.date, service
                        )
                    )
                except NoServiceRow:
                    closing_balance = 0.0
                if closing_balance or self.osv.record.accural.payment:
                    gvs_row = GvsSingleResultRow(
                        self.osv.date,
                        self.osv.record.address,
                        self.osv.record.accural,
                        self.account_details,
                        GvsDetailsRecord.get_dummy_instance(),
                        self.buildings,
                        service,
                    )
                    self.results.calculations.add_row(gvs_row)
            case 1:
                gvs_row = GvsSingleResultRow(
                    self.osv.date,
                    self.osv.record.address,
                    self.osv.record.accural,
                    self.account_details,
                    gvs_details_rows[0],
                    self.buildings,
                    service,
                )
                self.results.calculations.add_row(gvs_row)
            case 2:
                for num, gvs_details_row in enumerate(gvs_details_rows):
                    if not num:
                        gvs_row = GvsMultipleResultFirstRow(
                            self.osv.date,
                            self.osv.record.address,
                            self.osv.record.accural,
                            self.account_details,
                            gvs_details_row,
                            self.buildings,
                            service,
                        )
                    else:
                        gvs_row = GvsMultipleResultSecondRow(
                            self.osv.date,
                            self.osv.record.address,
                            self.osv.record.accural,
                            self.account_details,
                            gvs_details_row,
                            self.buildings,
                            service,
                        )
                    self.results.calculations.add_row(gvs_row)
        self._add_initial_balance_row(service)

    def _add_gvs_reaccural(self, record_type: CalculationRecordType):
        match record_type:
            case CalculationRecordType.GVS_REACCURAL:
                service = Service.GVS
            case CalculationRecordType.GVS_REACCURAL_ELEVATED:
                service = Service.GVS_ELEVATED
            case _:
                raise ValueError("Unknown result record type")
        try:
            reaccural_sum = self.account_details.get_service_month_reaccural(
                self.osv.date,
                service,
            )
        except NoServiceRow:
            return
        if not reaccural_sum:
            return
        gvs_details = GvsDetailsFileSingleton(
            os.path.join(
                self.base_dir,
                self.conf["gvs.dir"],
                f"{self.osv.date.month:02d}.{self.osv.date.year}.xlsx",
            ),
            int(self.conf["gvs_details.header_row"]),
            filter_func=lambda x: x.account,
        )
        gvs_details_rows: list[GvsDetailsRecord] = gvs_details.as_filtered_list(
            ("account",), (self.osv.record.address.account,)
        )
        try:
            gvs_details_row: GvsDetailsRecord = gvs_details_rows[0]
        except IndexError:
            gvs_details_row = GvsDetailsRecord.get_dummy_instance()
        reaccural_details = Reaccural(
            self.account_details, self.osv.date, reaccural_sum, service
        )
        reaccural_details.init_type(
            os.path.join(self.base_dir, self.conf["gvs.dir"]),
            int(self.conf["gvs_details.header_row"]),
        )
        for rec in reaccural_details.records:
            gvs_reaccural_row = GvsReaccuralResultRow(
                self.osv.date,
                self.osv.record.address,
                gvs_details_row,
                rec.date,
                rec.sum,
                reaccural_details.type,
                self.buildings,
                service,
                record_type,
            )
            if not reaccural_details.valid:
                gvs_reaccural_row.set_field(
                    38, "Не удалось разложить начисление на месяцы"
                )
            self.results.calculations.add_row(gvs_reaccural_row)

    def _add_gvs_elevated(self):
        service = Service.GVS_ELEVATED
        gvs_details = GvsDetailsFileSingleton(
            os.path.join(
                self.base_dir,
                self.conf["gvs.dir"],
                f"{self.osv.date.month:02d}.{self.osv.date.year}.xlsx",
            ),
            int(self.conf["gvs_details.header_row"]),
            lambda x: x.account,
        )
        gvs_details_rows: list[GvsDetailsRecord] = gvs_details.as_filtered_list(
            ("account",), (self.osv.record.address.account,)
        )
        if not gvs_details_rows:
            return
        try:
            row = GvsElevatedResultRow(
                self.osv.date,
                self.osv.record.address,
                self.osv.record.accural,
                self.account_details,
                gvs_details_rows[0],
                self.buildings,
                service,
            )
            self.results.calculations.add_row(row)
        except (NoServiceRow, ZeroDataResultRow):
            pass
        self._add_initial_balance_row(service)

    def _create_heating_reaccural_record(self, correction_date, correction_sum):
        service = Service.HEATING
        row = HeatingReaccuralResultRow(
            correction_date,
            self.osv.record.address,
            self.buildings,
            self.building_record.has_odpu,
            correction_sum,
            service,
        )
        self.results.calculations.add_row(row)

    def _add_heating_correction(self):
        service = Service.HEATING
        if self.osv.date.month != self.building_record.correction_month:
            return
        try:
            reaccural = self.account_details.get_service_month_reaccural(
                self.osv.date, service
            )
        except NoServiceRow:
            return
        if not reaccural:
            return
        reaccural = Decimal(reaccural).quantize(Decimal("0.01"))
        try:
            correction_record: HeatingCorrectionRecord = (
                self.heating_corrections.get_account_row(
                    self.account,
                    f"{self.osv.date.year-1}",
                )
            )
        except ValueError:
            return
        is_positive_correction: bool = False
        if correction_record.year_correction >= 0:
            is_positive_correction = True
        if (
            not is_positive_correction
            and reaccural != correction_record.year_correction
        ):
            logging.warning(
                "No heating correction for %s in %s. Reaccural/Correction: %s/%s",
                self.account,
                self.osv.date,
                reaccural,
                correction_record.year_correction,
            )
            return
        for month_num, month_abbr in enumerate(
            [m.lower() for m in calendar.month_abbr if m], start=1
        ):
            correction_sum = getattr(correction_record, month_abbr)
            correction_volume = getattr(correction_record, f"vkv_{month_abbr}")
            correction_date = MonthYear(month_num, self.osv.date.year - 1)
            if not correction_sum and not correction_volume:
                continue
            if correction_sum < 0:
                self._create_heating_reaccural_record(correction_date, correction_sum)
            odpu_records: list[
                HeatingVolumesOdpuRecord
            ] = self.heating_volumes_odpu.as_filtered_list(
                ("street", "house"),
                (correction_record.street, correction_record.house),
                f"{correction_date.year}",
            )
            if len(odpu_records) != 1:
                raise ValueError(
                    "Can't correctly determine an address for the last year correction: "
                    f"{correction_record.street} {correction_record.house}"
                )
            odpu_volume = getattr(odpu_records[0], month_abbr)
            row = HeatingCorrectionResultRow(
                self.osv.date,
                self.osv.record.address,
                correction_date,
                correction_sum,
                correction_volume,
                odpu_volume,
                service,
                self.buildings,
            )
            self.results.calculations.add_row(row)
        if is_positive_correction:
            self._add_future_installment_records(service)
            # self._add_closing_balance_records(service)

    def _add_future_installment_records(self, service):
        correction = HeatingPositiveCorrection(
            self.account_details,
            self.heating_corrections,
            self.osv.date,
            service,
        )
        if HeatingCorrectionAccountStatus.CLOSED_LAST_YEAR not in correction.type:
            account_closing_month = self.account_details.get_service_closing_month(
                correction.current_year, service
            )
            total_closing_balance: Decimal
            total_future_installment: Decimal
            for month_num in range(self.osv.date.month, 13):
                correction_date = MonthYear(month_num, correction.current_year)
                reaccural_sum = Decimal(
                    self.account_details.get_service_month_reaccural(
                        MonthYear(month_num, self.osv.date.year),
                        service,
                    )
                ).quantize(Decimal("0.01"))
                if month_num == self.osv.date.month:
                    _total_correction = Decimal(
                        correction.last_year_correction.year_correction
                    )
                    future_installment = _total_correction - reaccural_sum
                    total_closing_balance = reaccural_sum
                    total_future_installment = future_installment
                else:
                    future_installment = None
                    total_closing_balance += reaccural_sum  # type: ignore
                    total_future_installment -= reaccural_sum  # type: ignore
                row = HeatingPositiveCorrectionResultRow(
                    self.osv.date,
                    self.osv.record.address,
                    correction_date,
                    service,
                    future_installment,
                    total_closing_balance,
                    total_future_installment,
                    self.buildings,
                )
                self.results.calculations.add_row(row)
                if month_num == account_closing_month:
                    last_total_future_installment = total_future_installment
                    total_future_installment = Decimal("0.00")
                    future_installment = None
                    total_closing_balance = Decimal(
                        correction.last_year_correction.year_correction
                    ).quantize(Decimal("0.01"))
                    correction_date = MonthYear(month_num + 1, correction.current_year)
                    row = HeatingPositiveCorrectionResultRow(
                        self.osv.date,
                        self.osv.record.address,
                        correction_date,
                        service,
                        future_installment,
                        total_closing_balance,
                        total_future_installment,
                        self.buildings,
                    )
                    self.results.calculations.add_row(row)
                    try:
                        next_month_reaccural = Decimal(
                            self.account_details.get_service_month_reaccural(
                                MonthYear(month_num + 1, self.osv.date.year),
                                service,
                            )
                        ).quantize(Decimal("0.01"))
                    except NoServiceRow:
                        # not exactly account_closing_month,
                        # but the last month we have data for
                        break
                    excessive_reaccural = (
                        next_month_reaccural - last_total_future_installment
                    )
                    correction_date = MonthYear(month_num, correction.current_year)
                    if excessive_reaccural:
                        row = HeatingPositiveCorrectionExcessiveReaccuralResultRow(
                            correction_date.next,
                            self.osv.record.address,
                            correction_date,
                            excessive_reaccural,
                            service,
                            self.buildings,
                        )
                        self.results.calculations.add_row(row)
                    break
        else:
            correction_date = self.osv.date
            future_installment = None
            total_closing_balance = Decimal(
                correction.last_year_correction.year_correction
            ).quantize(Decimal("0.01"))
            total_future_installment = Decimal("0.00")
            row = HeatingPositiveCorrectionResultRow(
                self.osv.date,
                self.osv.record.address,
                correction_date,
                service,
                future_installment,
                total_closing_balance,
                total_future_installment,
                self.buildings,
            )
            self.results.calculations.add_row(row)

    def _add_closing_balance_records(self, service):
        for cur_year, start_month in [
            (self.osv.date.year - 1, 12),
            (self.osv.date.year, self.osv.date.month),
        ]:
            try:
                correction_record: HeatingCorrectionRecord = (
                    self.heating_corrections.get_account_row(
                        self.account,
                        f"{cur_year}",
                    )
                )
            except ValueError:
                continue
            except KeyError:
                # sheet not found, which means the year is current and yet no data
                # nothing needs to be done here
                return
            month_abbrs = reversed(
                [m.lower() for m in calendar.month_abbr if m][:start_month]
            )
            for cnt, month_abbr in enumerate(month_abbrs):
                month_num = start_month - cnt
                correction_sum = getattr(correction_record, month_abbr)
                correction_date = MonthYear(month_num, cur_year)
                if correction_sum:
                    break
                row = HeatingNegativeCorrectionZeroResultRow(
                    self.osv.date,
                    self.osv.record.address,
                    correction_date,
                    self.account_details,
                    service,
                    self.buildings,
                )
                self.results.calculations.add_row(row)

    def _add_accounts_record(self):
        rec = self.osv.record.address
        account_data: AccountChangebleInfo = (rec.name, rec.population)
        if (
            rec.account in self.seen_account_info
            and self.seen_account_info[rec.account] == account_data
        ):
            return
        self.seen_account_info[rec.account] = account_data
        row = AccountsResultRow(self.osv.date, rec)
        self.results.accounts.add_row(row)

    def _add_people_records(self):
        fio_delimeter = ";"
        rec = self.osv.record.address
        if not rec.name:
            return
        if fio_delimeter not in rec.name:
            return
        if (
            rec.account in self.seen_people_names
            and self.seen_people_names[rec.account] == rec.name
        ):
            return
        self.seen_people_names[rec.account] = rec.name
        names = rec.name.split(fio_delimeter)
        for name in names:
            name = name.strip()
            if not name:
                continue
            row = PeopleResultRow(self.osv.date, rec, name)
            self.results.people.add_row(row)

    def process_osv(self, osv_file_name) -> None:
        "Process OSV file currently set as self.osv_file"
        self.osv = OsvFile(osv_file_name, self.conf)
        for _ in self.osv.init_next_record(self.buildings):
            if not self._is_debugging_account(self.osv.record.address.account):
                continue
            self.account = self.osv.record.address.account
            try:
                account_details_path = os.path.join(
                    self.base_dir,
                    self.conf["account_details.dir"],
                    f"{self.osv.record.address.account}.xlsx",
                )
                self.account_details = AccountDetailsFileSingleton(
                    self.account,
                    account_details_path,
                    int(self.conf["account_details.header_row"]),
                )
            except FileNotFoundError:
                self.error_handler.show(
                    "no_account_details",
                    self.account,
                    "Account details file not found: %s",
                    self.account,
                )
                continue
            self.building_record = self.buildings.get_address_row(
                self.osv.record.address.address,
                str(self.osv.date.year),
            )
            if self.is_config_option_true("fill_accounts"):
                self._add_accounts_record()
            if self.is_config_option_true("fill_people"):
                self._add_people_records()
            if self.is_config_option_true("fill_calculations"):
                self._add_heating()
                self._add_gvs()
                self._add_gvs_elevated()
                self._add_gvs_reaccural(CalculationRecordType.GVS_REACCURAL)
                self._add_gvs_reaccural(CalculationRecordType.GVS_REACCURAL_ELEVATED)
                self._add_heating_correction()
        self.osv.close()

    def close(self):
        "Closes all file descriptors that might still be open"
        for attr_name in dir(self):
            if attr_name.startswith("_"):
                continue
            attr = getattr(self, attr_name)
            if not attr:
                continue
            if not isinstance(attr, BaseWorkBook):
                continue
            attr.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, exc_traceback):
        self.close()


if __name__ == "__main__":
    config = configparser.ConfigParser(inline_comment_prefixes="#")
    config.read(CONFIG_PATH)
    LOGFORMAT = "%(asctime)s %(levelname)s - %(message)s"
    logging.basicConfig(
        level=config["DEFAULT"]["loglevel"],
        format=LOGFORMAT,
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    for exp in OSVDATA_REGEXP:
        osvdata_regexp_compiled.append(re.compile(exp))
    for section in config.sections():
        region: RegionDir
        region_path = os.path.join(config["DEFAULT"]["base_dir"], section)
        with RegionDir(region_path, config[section]) as region:
            try:
                [region.process_osv(osv) for osv in region.osv_files]
            except Exception as err:  # pylint: disable=W0718
                logging.critical("General exception: %s.", err.args)
                raise
            if not region.is_config_option_true("fill_calculations"):
                region.results.save()
                continue
            filled_table = WorkBookDataUpdater(region.results, ResultSheet.CALCULATIONS)
            filled_table.prepare_records_cache(
                GvsIpuMetric,
                filter_func=lambda s: s.type_name
                == CalculationRecordType.GVS_ACCURAL.name,
            )
            filled_table.find_gvs_ipu_replacements()
            filled_table.prepare_records_cache(
                AccountClosingBalance,
                filter_func=lambda s: s.type_name
                in (
                    CalculationRecordType.HEATING_ACCURAL.name,
                    CalculationRecordType.HEATING_POSITIVE_CORRECTION.name,
                ),
            )
            filled_table.decrease_closing_balance()
            logging.info("Total changes: %s", filled_table.changes_counter)
            region.results.save()
