" Energobilling data calculation/formatting "
import calendar
import configparser
from decimal import Decimal
import logging
import os
import re
import sys
from dataclasses import dataclass
from os.path import basename
from typing import Mapping

from lib.addressfile import AddressFile
from lib.datatypes import MonthYear
from lib.detailsfile import (
    AccountDetailsFileSingleton,
    GvsDetailsFileSingleton,
    GvsDetailsRecord,
)
from lib.exceptions import NoServiceRow, ZeroDataResultRow
from lib.gvsipuchange import IpuReplacementFinder
from lib.heatingcorrections import (
    HeatingCorrectionRecord,
    HeatingCorrectionsFile,
    HeatingVolumesOdpuFile,
    HeatingVolumesOdpuRecord,
)
from lib.helpers import ExcelHelpers
from lib.osvfile import (
    OSVDATA_REGEXP,
    OsvAccuralRecord,
    OsvAddressRecord,
    OsvFile,
    OsvRecord,
    osvdata_regexp_compiled,
)
from lib.reaccural import Reaccural
from lib.resultfile import (
    GvsElevatedResultRow,
    GvsMultipleResultFirstRow,
    GvsMultipleResultSecondRow,
    GvsReaccuralResultRow,
    GvsSingleResultRow,
    HeatingLastYearCorrectionZeroResultRow,
    HeatingLastYearNegativeCorrectionResultRow,
    HeatingResultRow,
    ResultFile,
    ResultRecordType,
)


CONFIG_PATH = "./config.ini"


@dataclass
class ColumnIndex:
    "Indexes of fields in the table, zero-based"
    address: int
    heating: int
    gvs: int
    reaccurance: int
    total: int
    gvs_elevated_percent: int


class RegionDir:
    "A directory with the data of a particular region"

    account_details: AccountDetailsFileSingleton
    osv_file: OsvFile
    osv: OsvRecord
    account: str

    def __init__(self, base_dir: str, conf: Mapping[str, str]) -> None:
        logging.info("Initialazing %s region data...", base_dir)
        self.conf = {k: v.strip() for k, v in conf.items()}
        self.base_dir = base_dir
        self.osv_path = os.path.join(self.base_dir, self.conf["osv.dir"])
        self.heating_yearly_correction_dates = [
            MonthYear(*[int(x) for x in reversed(date.split("-"))])
            for date in self.conf["heating.yearly_corrections_dates"].split(",")
        ]
        self.odpus = AddressFile(
            os.path.join(self.base_dir, conf["file.odpu_address"]),
            self.conf,
        )
        logging.info(
            "Read ODPU address data: %s sheets, %s elements total",
            self.odpus.get_sheets_count(),
            self.odpus.get_strings_count(),
        )
        self.heating_average = AddressFile(
            os.path.join(self.base_dir, conf["file.heating_average"]),
            self.conf,
        )
        logging.info(
            "Read heating average: %s sheets, %s elements total",
            self.heating_average.get_sheets_count(),
            self.heating_average.get_strings_count(),
        )
        self.buildings = AddressFile(
            os.path.join(self.base_dir, conf["file.building_address"]),
            self.conf,
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
            os.path.join(self.osv_path, f)
            for f in os.listdir(self.osv_path)
            if os.path.isfile(os.path.join(self.osv_path, f)) and not f.startswith(".")
        ]
        self.osv_files: list[str] = sorted(
            osv_files, key=lambda s: (basename(s)[2:6], basename(s)[0:2])
        )
        self.osv_files = self.osv_files[: int(self.conf["max_osv_files"])]
        for file in self.osv_files:
            if file.endswith(".xlsx"):
                continue
            else:
                logging.critical("Non *.xlsx found in OSV_DIR, exiting")
                sys.exit(1)
        logging.info("Initialazing %s region data done", base_dir)
        self.results = ResultFile(self.base_dir, self.conf)

    def _get_osv_column_indexes(self) -> ColumnIndex:
        "Calculates indexes of columns in the table"
        if self.osv_file is None:
            raise ValueError("OSV file was not initialized yet")
        try:
            header_row = int(self.conf["osv.header_row"])
            column_index = ColumnIndex(
                ExcelHelpers.get_col_by_name(self.osv_file.sheet, "Адрес", header_row)
                - 1,
                ExcelHelpers.get_col_by_name(
                    self.osv_file.sheet, "Отопление", header_row
                )
                - 1,
                ExcelHelpers.get_col_by_name(
                    self.osv_file.sheet,
                    "Тепловая энергия для подогрева воды",
                    header_row,
                )
                - 1,
                ExcelHelpers.get_col_by_name(
                    self.osv_file.sheet, "Перерасчеты", header_row
                )
                - 1,
                ExcelHelpers.get_col_by_name(self.osv_file.sheet, "Всего", header_row)
                - 1,
                ExcelHelpers.get_col_by_name(
                    self.osv_file.sheet,
                    "Тепловая энергия для подогрева воды (повышенный %)",
                    header_row,
                )
                - 1,
            )
        except ValueError as err:
            logging.warning("Check column names: %s", err)
            raise
        return column_index

    def _init_current_osv_row(self, row, column_index_data) -> OsvRecord | None:
        "Parses OSV row and returns OvsRecord instance"
        address_cell = row[column_index_data.address]
        try:
            osv_address_rec = OsvAddressRecord.get_instance(address_cell)
            logging.debug("Address record %s understood as %s", row[0], osv_address_rec)
            if not ExcelHelpers.is_address_in_list(
                osv_address_rec.address,
                self.buildings.get_sheet_data_formatted(str(self.osv_file.date.year)),
            ):
                return None
            osv_accural_rec = OsvAccuralRecord(
                float(row[column_index_data.heating]),
                float(row[column_index_data.gvs]),
                float(row[column_index_data.reaccurance]),
                float(row[column_index_data.total]),
                float(row[column_index_data.gvs_elevated_percent]),
            )
            logging.debug("Accural record %s understood as %s", row[0], osv_accural_rec)
        except AttributeError as err:
            logging.warning("%s. Malformed record: %s", err, row)
            return None
        return OsvRecord(osv_address_rec, osv_accural_rec)

    def _is_debugging_current_account(self):
        "Checks of single account debugging is enabled in config file"
        try:
            if config["DEFAULT"]["account"] == self.osv.address_record.account:
                return True
        except KeyError:
            return True
        return False

    def _process_heating_data(self):
        if not any(
            (
                self.osv.accural_record.heating,
                self.osv.accural_record.reaccural,
                self.osv.accural_record.payment,
            )
        ):
            return
        try:
            heating_row = HeatingResultRow(
                self.osv_file.date,
                self.osv.address_record,
                self.osv.accural_record,
                self.odpus,
                self.heating_average,
                self.account_details,
            )
            self.results.add_row(heating_row)
        except NoServiceRow:
            pass

    def _process_gvs_data(self):
        if not any(
            (
                self.osv.accural_record.gvs,
                self.osv.accural_record.reaccural,
                self.osv.accural_record.payment,
            )
        ):
            return
        gvs_details = GvsDetailsFileSingleton(
            os.path.join(
                self.base_dir,
                self.conf["gvs.dir"],
                f"{self.osv_file.date.month:02d}.{self.osv_file.date.year}.xlsx",
            ),
            int(self.conf["gvs_details.header_row"]),
            lambda x: x.account,
        )
        gvs_details_rows: list[GvsDetailsRecord] = gvs_details.as_filtered_list(
            ("account",), (self.osv.address_record.account,)
        )
        if len(gvs_details_rows) > 2:
            gvs_details_rows = [gvs_details_rows[0], gvs_details_rows[-1]]
            logging.warning(
                "Too many GVS details records for account %s in %s, skipped",
                gvs_details_rows[0].account,
                self.osv_file.date,
            )
        match len(gvs_details_rows):
            case 0:
                pass
            case 1:
                gvs_row = GvsSingleResultRow(
                    self.osv_file.date,
                    self.osv.address_record,
                    self.osv.accural_record,
                    self.account_details,
                    gvs_details_rows[0],
                )
                self.results.add_row(gvs_row)
            case 2:
                for num, gvs_details_row in enumerate(gvs_details_rows):
                    if not num:
                        gvs_row = GvsMultipleResultFirstRow(
                            self.osv_file.date,
                            self.osv.address_record,
                            self.osv.accural_record,
                            self.account_details,
                            gvs_details_row,
                        )
                    else:
                        gvs_row = GvsMultipleResultSecondRow(
                            self.osv_file.date,
                            self.osv.address_record,
                            self.osv.accural_record,
                            self.account_details,
                            gvs_details_row,
                        )
                    self.results.add_row(gvs_row)

    def _process_gvs_reaccural_data(self, record_type: ResultRecordType):
        match record_type:
            case ResultRecordType.GVS_REACCURAL:
                service = "Тепловая энергия для подогрева воды"
            case ResultRecordType.GVS_REACCURAL_ELEVATED:
                service = "Тепловая энергия для подогрева воды (повышенный %)"
            case _:
                raise ValueError("Unknown result record type")
        try:
            reaccural_sum = self.account_details.get_service_month_reaccural(
                self.osv_file.date,
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
                f"{self.osv_file.date.month:02d}.{self.osv_file.date.year}.xlsx",
            ),
            int(self.conf["gvs_details.header_row"]),
            filter_func=lambda x: x.account,
        )
        gvs_details_rows: list[GvsDetailsRecord] = gvs_details.as_filtered_list(
            ("account",), (self.osv.address_record.account,)
        )
        try:
            gvs_details_row: GvsDetailsRecord = gvs_details_rows[0]
        except IndexError:
            gvs_details_row = GvsDetailsRecord.get_dummy_instance()
        reaccural_details = Reaccural(
            self.account_details, self.osv_file.date, reaccural_sum, service
        )
        reaccural_details.init_type(
            os.path.join(self.base_dir, self.conf["gvs.dir"]),
            int(self.conf["gvs_details.header_row"]),
        )
        for rec in reaccural_details.records:
            gvs_reaccural_row = GvsReaccuralResultRow(
                self.osv_file.date,
                self.osv.address_record,
                gvs_details_row,
                rec.date,
                rec.sum,
                reaccural_details.type,
                service,
                record_type,
            )
            if not reaccural_details.valid:
                gvs_reaccural_row.set_field(
                    38, "Не удалось разложить начисление на месяцы"
                )
            self.results.add_row(gvs_reaccural_row)

    def _process_gvs_elevated_data(self):
        service = "Тепловая энергия для подогрева воды (повышенный %)"
        gvs_details = GvsDetailsFileSingleton(
            os.path.join(
                self.base_dir,
                self.conf["gvs.dir"],
                f"{self.osv_file.date.month:02d}.{self.osv_file.date.year}.xlsx",
            ),
            int(self.conf["gvs_details.header_row"]),
            lambda x: x.account,
        )
        gvs_details_rows: list[GvsDetailsRecord] = gvs_details.as_filtered_list(
            ("account",), (self.osv.address_record.account,)
        )
        if not gvs_details_rows:
            return
        try:
            row = GvsElevatedResultRow(
                self.osv_file.date,
                self.osv.address_record,
                self.osv.accural_record,
                self.account_details,
                gvs_details_rows[0],
                service,
            )
            self.results.add_row(row)
        except (NoServiceRow, ZeroDataResultRow):
            pass

    def _process_last_year_negative_heating_correction(self):
        service = "Отопление"

        if self.osv_file.date not in self.heating_yearly_correction_dates:
            return
        try:
            reaccural = self.account_details.get_service_month_reaccural(
                self.osv_file.date, service
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
                    f"{self.osv_file.date.year-1}",
                )
            )
        except ValueError:
            return
        if correction_record.year_correction >= 0:
            return
        if reaccural != correction_record.year_correction:
            logging.warning(
                "Could not determine last year heating correction for %s in %s. "
                "Reaccural: %s; last year correction: %s",
                self.account,
                self.osv_file.date,
                reaccural,
                correction_record.year_correction,
            )
            return
        for month_num, month_abbr in enumerate(
            [m.lower() for m in calendar.month_abbr if m], start=1
        ):
            correction_sum = getattr(correction_record, month_abbr)
            correction_date = MonthYear(month_num, self.osv_file.date.year - 1)
            if not correction_sum:
                continue
            correction_volume = getattr(correction_record, f"vkv_{month_abbr}")
            odpu_records: HeatingVolumesOdpuRecord = (
                self.heating_volumes_odpu.as_filtered_list(
                    ("street", "house"),
                    (correction_record.street, correction_record.house),
                    f"{correction_date.year}",
                )
            )
            if len(odpu_records) != 1:
                raise ValueError(
                    f"Can't correctly determine address for last year correction: \
                    {correction_record.street} {correction_record.house}"
                )
            odpu_volume = getattr(odpu_records[0], month_abbr)
            row = HeatingLastYearNegativeCorrectionResultRow(
                self.osv_file.date,
                self.osv.address_record,
                correction_date,
                correction_sum,
                correction_volume,
                odpu_volume,
            )
            self.results.add_row(row)
        # add zero records closing balance records:
        for cur_year, start_month in [
            (self.osv_file.date.year - 1, 12),
            (self.osv_file.date.year, self.osv_file.date.month),
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
                row = HeatingLastYearCorrectionZeroResultRow(
                    self.osv_file.date,
                    self.osv.address_record,
                    correction_date,
                    self.account_details,
                    service,
                )
                self.results.add_row(row)

    def _process_osv(self, osv_file_name) -> None:
        "Process OSV file currently set as self.osv_file"
        self.osv_file = OsvFile(osv_file_name, self.conf)
        column_index_data = self._get_osv_column_indexes()
        for row in self.osv_file.get_data_row():
            osv = self._init_current_osv_row(row, column_index_data)
            if not osv:
                continue
            self.osv = osv
            if not self._is_debugging_current_account():
                continue
            self.account = self.osv.address_record.account
            self.account_details = AccountDetailsFileSingleton(
                self.account,
                os.path.join(
                    self.base_dir,
                    self.conf["account_details.dir"],
                    f"{self.osv.address_record.account}.xlsx",
                ),
                int(self.conf["account_details.header_row"]),
            )
            self._process_heating_data()
            self._process_gvs_data()
            self._process_gvs_reaccural_data(ResultRecordType.GVS_REACCURAL)
            self._process_gvs_elevated_data()
            self._process_gvs_reaccural_data(ResultRecordType.GVS_REACCURAL_ELEVATED)
            self._process_last_year_negative_heating_correction()

    def read_osvs(self) -> None:
        "Reads OSV files row by row and writes data to result table"
        for file_name in self.osv_files:
            try:
                self._process_osv(file_name)
            except Exception as err:  # pylint: disable=W0718
                logging.critical("General exception: %s.", err.args)
                raise
            finally:
                self.close()

    def close(self):
        "Closes all file descriptors that might still be open"
        try:
            self.account_details.close()
        except AttributeError:
            pass
        try:
            self.buildings.close()
        except AttributeError:
            pass
        try:
            self.heating_average.close()
        except AttributeError:
            pass
        try:
            self.odpus.close()
        except AttributeError:
            pass
        try:
            self.osv_file.close()
        except AttributeError:
            pass
        try:
            self.results.close()
        except AttributeError:
            pass


if __name__ == "__main__":
    config = configparser.ConfigParser(inline_comment_prefixes="#")
    config.read(CONFIG_PATH)
    LOGFORMAT = "%(asctime)s %(levelname)s - %(message)s"
    logging.basicConfig(level=config["DEFAULT"]["loglevel"], format=LOGFORMAT)
    for exp in OSVDATA_REGEXP:
        osvdata_regexp_compiled.append(re.compile(exp))
    for section in config.sections():
        try:
            region = RegionDir(
                os.path.join(config["DEFAULT"]["base_dir"], section), config[section]
            )
            region.read_osvs()
            region.results.save()
        finally:
            try:
                region.close()
            except NameError:
                pass
        gvs_ipu_change = IpuReplacementFinder(
            os.path.join(
                config["DEFAULT"]["base_dir"],
                section,
                config["DEFAULT"]["result_file"].split("@", 1)[0],
            )
        )
        gvs_ipu_change.find_replacements()
        gvs_ipu_change.save()
