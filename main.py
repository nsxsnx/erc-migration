" Energobilling data calculation/formatting "
import configparser
import logging
import os
import re
import sys
from collections import Counter
from dataclasses import dataclass
from os.path import basename

from lib.addressfile import AddressFile
from lib.detailsfile import (
    AccountDetailsFileSingleton,
    AccountDetailsRecord,
    GvsDetailsFileSingleton,
    GvsDetailsRecord,
)
from lib.exceptions import NoServiceRow, ZeroDataResultRow
from lib.gvsipuchange import IpuReplacementFinder
from lib.helpers import ExcelHelpers
from lib.osvfile import (
    OSVDATA_REGEXP,
    OsvAccuralRecord,
    OsvAddressRecord,
    OsvFile,
    OsvRecord,
    osvdata_regexp_compiled,
)
from lib.reaccural import Reaccural, ReaccuralType
from lib.resultfile import (
    GvsElevatedResultRow,
    GvsMultipleResultFirstRow,
    GvsMultipleResultSecondRow,
    GvsReaccuralResultRow,
    GvsSingleResultRow,
    HeatingResultRow,
    ResultFile,
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

    def __init__(self, base_dir: str, conf: dict) -> None:
        self.reaccural_counter = Counter()
        logging.info("Initialazing %s region data...", base_dir)
        self.conf = {k: v.strip() for k, v in conf.items()}
        self.base_dir = base_dir
        self.osv_path = os.path.join(self.base_dir, self.conf["osv.dir"])
        self.osv_file: OsvFile | None = None
        self.osv: OsvRecord | None = None
        self.account: str | None = None
        self.account_details: AccountDetailsFileSingleton | None = None
        logging.info("Reading ODPU addresses...")
        self.odpus = AddressFile(
            os.path.join(self.base_dir, conf["file.odpu_address"]),
            self.conf,
        )
        logging.info(
            "Reading ODPU address data done, found %s sheets with total of %s elements",
            self.odpus.get_sheets_count(),
            self.odpus.get_strings_count(),
        )
        logging.info("Reading heating average...")
        self.heating_average = AddressFile(
            os.path.join(self.base_dir, conf["file.heating_average"]),
            self.conf,
        )
        logging.info(
            "Reading heating average done, found %s sheets with total of %s elements",
            self.heating_average.get_sheets_count(),
            self.heating_average.get_strings_count(),
        )
        logging.info("Reading buildings ...")
        self.buildings = AddressFile(
            os.path.join(self.base_dir, conf["file.building_address"]),
            self.conf,
        )
        logging.info(
            "Reading buildings done, found %s sheets with total of %s elements",
            self.buildings.get_sheets_count(),
            self.buildings.get_strings_count(),
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

    def _get_column_indexes(self) -> ColumnIndex:
        "Calculates indexes of columns in the table"
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
            if not ExcelHelpers.address_in_list(
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
        if not (
            self.osv.accural_record.heating
            or self.osv.accural_record.reaccural
            or self.osv.accural_record.payment
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
        if not (
            self.osv.accural_record.gvs
            or self.osv.accural_record.reaccural
            or self.osv.accural_record.payment
        ):
            return
        gvs_details = GvsDetailsFileSingleton(
            os.path.join(
                self.base_dir,
                self.conf["gvs.dir"],
                f"{self.osv_file.date.month:02d}.{self.osv_file.date.year}.xlsx",
            ),
            int(self.conf["gvs_details.header_row"]),
            GvsDetailsRecord,
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

    def _process_gvs_reaccural_data(self, service: str):
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
            GvsDetailsRecord,
            lambda x: x.account,
        )
        gvs_details_rows: list[GvsDetailsRecord] = gvs_details.as_filtered_list(
            ("account",), (self.osv.address_record.account,)
        )
        try:
            gvs_details_row: GvsDetailsRecord = gvs_details_rows[0]
        except IndexError:
            gvs_details_row: GvsDetailsRecord = GvsDetailsRecord.get_dummy_instance()
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
            )
            if not reaccural_details.valid:
                gvs_reaccural_row.set_field(
                    38, "Не удалось разложить начисление на месяцы"
                )
            self.results.add_row(gvs_reaccural_row)
        self.reaccural_counter.update([reaccural_details.valid])

    def _process_gvs_elevated_data(self):
        service = "Тепловая энергия для подогрева воды (повышенный %)"
        gvs_details = GvsDetailsFileSingleton(
            os.path.join(
                self.base_dir,
                self.conf["gvs.dir"],
                f"{self.osv_file.date.month:02d}.{self.osv_file.date.year}.xlsx",
            ),
            int(self.conf["gvs_details.header_row"]),
            GvsDetailsRecord,
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

    def _process_osv(self, osv_file_name) -> None:
        "Process OSV file currently set as self.osv_file"
        self.osv_file = OsvFile(osv_file_name, self.conf)
        column_index_data = self._get_column_indexes()
        for row in self.osv_file.get_data_row():
            self.osv = self._init_current_osv_row(row, column_index_data)
            if not self.osv:
                continue
            if not self._is_debugging_current_account:
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
                AccountDetailsRecord,
            )
            self._process_heating_data()
            self._process_gvs_data()
            self._process_gvs_reaccural_data("Тепловая энергия для подогрева воды")
            self._process_gvs_elevated_data()
            self._process_gvs_reaccural_data(
                "Тепловая энергия для подогрева воды (повышенный %)"
            )
        # logging.info("Total valid/invalid reacurals: %s", self.reaccural_counter)

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
            region.close()
        gvs_ipu_change = IpuReplacementFinder(
            os.path.join(
                config["DEFAULT"]["base_dir"],
                section,
                config["DEFAULT"]["result_file"].split("@", 1)[0],
            )
        )
        gvs_ipu_change.find_replacements()
        gvs_ipu_change.save()
