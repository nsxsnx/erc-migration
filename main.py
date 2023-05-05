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
from lib.exceptions import NoServiceRow
from lib.gvsipuchange import FilledResultRecord, FilledResultTable, GvsIpuMetric
from lib.helpers import ExcelHelpers
from lib.osvfile import (
    OSVDATA_REGEXP,
    OsvAccuralRecord,
    OsvAddressRecord,
    OsvFile,
    osvdata_regexp_compiled,
)
from lib.reaccural import Reaccural, ReaccuralType
from lib.resultfile import (
    GvsIpuInstallDates,
    GvsMultipleResultFirstRow,
    GvsMultipleResultSecondRow,
    GvsReaccuralResultRow,
    GvsSingleResultRow,
    HeatingResultRow,
    ResultFile,
    ResultRecordType,
)

CONFIG_PATH = "./config.ini"


@dataclass
class RowIndex:
    "Indexes of fields in rhe table, zero-based"
    address: int
    heating: int
    gvs: int
    reaccurance: int
    total: int


class RegionDir:
    "A directory with the data of a particular region"

    def __init__(self, base_dir: str, conf: dict) -> None:
        self.reaccural_counter = Counter()
        logging.info("Initialazing %s region data...", base_dir)
        self.conf = {k: v.strip() for k, v in conf.items()}
        self.base_dir = base_dir
        self.osv_path = os.path.join(self.base_dir, self.conf["osv.dir"])
        self.osv_file: OsvFile | None = None
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

    def process_osv(self) -> None:
        "Process OSV file currently set as self.osv_file"
        try:
            header_row = int(self.conf["osv.header_row"])
            row_index = RowIndex(
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
            )
        except ValueError as err:
            logging.warning("Check column names: %s", err)
            raise
        for row in self.osv_file.get_data_row():
            address_cell = row[row_index.address]
            try:
                osv_address_rec = OsvAddressRecord.get_instance(address_cell)
                logging.debug(
                    "Address record %s understood as %s", row[0], osv_address_rec
                )
                if not ExcelHelpers.address_in_list(
                    osv_address_rec.address,
                    self.buildings.get_sheet_data_formatted(
                        str(self.osv_file.date.year)
                    ),
                ):
                    continue
                osv_accural_rec = OsvAccuralRecord(
                    float(row[row_index.heating]),
                    float(row[row_index.gvs]),
                    float(row[row_index.reaccurance]),
                    float(row[row_index.total]),
                )
                logging.debug(
                    "Accural record %s understood as %s", row[0], osv_accural_rec
                )
            except AttributeError as err:
                logging.warning("%s. Malformed record: %s", err, row)
                continue
            try:
                if config["DEFAULT"]["account"] != osv_address_rec.account:
                    continue
            except KeyError:
                pass
            # region heating
            if not (
                osv_accural_rec.heating
                or osv_accural_rec.reaccural
                or osv_accural_rec.payment
            ):
                continue
            account = osv_address_rec.account
            account_details = AccountDetailsFileSingleton(
                account,
                os.path.join(
                    self.base_dir,
                    self.conf["account_details.dir"],
                    f"{account}.xlsx",
                ),
                int(self.conf["account_details.header_row"]),
                AccountDetailsRecord,
            )
            try:
                heating_row = HeatingResultRow(
                    self.osv_file.date,
                    osv_address_rec,
                    osv_accural_rec,
                    self.odpus,
                    self.heating_average,
                    account_details,
                )
                self.results.add_row(heating_row)
            except NoServiceRow:
                pass
            # endregion
            # region gsv
            if not (
                osv_accural_rec.gvs
                or osv_accural_rec.reaccural
                or osv_accural_rec.payment
            ):
                continue
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
                ("account",), (osv_address_rec.account,)
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
                        osv_address_rec,
                        osv_accural_rec,
                        account_details,
                        gvs_details_rows[0],
                    )
                    self.results.add_row(gvs_row)
                case 2:
                    for num, gvs_details_row in enumerate(gvs_details_rows):
                        if not num:
                            gvs_row = GvsMultipleResultFirstRow(
                                self.osv_file.date,
                                osv_address_rec,
                                osv_accural_rec,
                                account_details,
                                gvs_details_row,
                            )
                        else:
                            gvs_row = GvsMultipleResultSecondRow(
                                self.osv_file.date,
                                osv_address_rec,
                                osv_accural_rec,
                                account_details,
                                gvs_details_row,
                            )
                        self.results.add_row(gvs_row)
            # endregion
            # region reaccural
            try:
                gvs_details_row: GvsDetailsRecord = gvs_details_rows[0]
            except IndexError:
                gvs_details_row: GvsDetailsRecord = (
                    GvsDetailsRecord.get_dummy_instance()
                )
            try:
                reaccural_sum = account_details.get_service_month_reaccural(
                    self.osv_file.date, "Тепловая энергия для подогрева воды"
                )
                if not reaccural_sum:
                    continue
                reaccural_details = Reaccural(
                    account_details, self.osv_file.date, reaccural_sum
                )
                # get type of Reaccural based on the data of a previous GVS file:
                prev_date = self.osv_file.date.previous
                prev_gvs_details = GvsDetailsFileSingleton(
                    os.path.join(
                        self.base_dir,
                        self.conf["gvs.dir"],
                        f"{prev_date.month:02d}.{prev_date.year}.xlsx",
                    ),
                    int(self.conf["gvs_details.header_row"]),
                    GvsDetailsRecord,
                    lambda x: x.account,
                )
                prev_gvs_details_rows: list[
                    GvsDetailsRecord
                ] = prev_gvs_details.as_filtered_list(
                    ("account",), (osv_address_rec.account,)
                )
                try:
                    row: GvsDetailsRecord = prev_gvs_details_rows[0]
                    if row.consumption_average:
                        reaccural_details.set_type(ReaccuralType.AVERAGE)
                    elif row.consumption_ipu:
                        reaccural_details.set_type(ReaccuralType.IPU)
                    else:
                        reaccural_details.set_type(ReaccuralType.NORMATIVE)
                except IndexError:
                    reaccural_details.set_type(ReaccuralType.NORMATIVE)
                for rec in reaccural_details.records:
                    gvs_reaccural_row = GvsReaccuralResultRow(
                        self.osv_file.date,
                        osv_address_rec,
                        gvs_details_row,
                        rec.date,
                        rec.sum,
                        reaccural_details.type,
                    )
                    if not reaccural_details.valid:
                        gvs_reaccural_row.set_field(
                            38, "Не удалось разложить начисление на месяцы"
                        )
                    self.results.add_row(gvs_reaccural_row)
                self.reaccural_counter.update([reaccural_details.valid])
            except NoServiceRow:
                pass
            # endregion
        logging.info("Total valid/invalid reaacurals: %s", self.reaccural_counter)

    def read_osvs(self) -> None:
        "Reads OSV files row by row and writes data to result table"
        for file in self.osv_files:
            try:
                self.osv_file = OsvFile(file, self.conf)
                self.process_osv()
            except Exception as err:  # pylint: disable=W0718
                logging.critical("General exception: %s.", err.args)
                self.close()
                raise
            else:
                self.osv_file.workbook.close()

    def close(self):
        "Closes all opened descriptors"
        try:
            self.odpus.close()
            self.heating_average.close()
            self.buildings.close()
            self.results.close()
            self.osv_file.close()
        except (NameError, AttributeError):
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
            region.results.close()
        finally:
            try:
                region.close()
            except NameError:
                pass
        logging.info("Reading GVS IPU data from results table...")
        filename = config["DEFAULT"]["result_file"].split("@", 1)[0]
        gvsipus = FilledResultTable(
            os.path.join(config["DEFAULT"]["base_dir"], section, filename),
            3,
            FilledResultRecord,
            filter_func=lambda s: s.type == ResultRecordType.GVS_ACCURAL.name,
            max_col=22,
        )
        logging.info("Reading GVS IPU data from results table done")
        logging.info("Looking for IPU replacement...")
        active_sheet = gvsipus.workbook.active
        gvs_accounts = gvsipus.get_field_values("account")
        TOTAL_IPU_REPLACEMENTS = 0
        for gvs_account in gvs_accounts:
            counters: list[GvsIpuMetric] = gvsipus.as_filtered_list(
                ("account",), (gvs_account,)
            )
            IGNORE_NEXT = False
            for i, current_el in enumerate(counters[:-1]):
                next_el = counters[i + 1]
                if IGNORE_NEXT:
                    IGNORE_NEXT = False
                    continue
                if current_el.date == next_el.date:
                    IGNORE_NEXT = True
                    continue
                if current_el.counter_number != next_el.counter_number:
                    row_num = current_el.row
                    type_cell = active_sheet[f"U{row_num}"]
                    type_cell.value = "При снятии прибора"
                    row_num = next_el.row
                    type_cell = active_sheet[f"U{row_num}"]
                    type_cell.value = "При установке"
                    GvsIpuInstallDates[gvs_account] = next_el.metric_date
                    logging.debug(current_el)
                    logging.debug(next_el)
                    TOTAL_IPU_REPLACEMENTS += 1
                if gvs_account in GvsIpuInstallDates:
                    row_num = next_el.row
                    date_cell = active_sheet[f"K{row_num}"]
                    date_cell.value = GvsIpuInstallDates[gvs_account]
        logging.info(
            "Total additional IPU replacements found: %s", TOTAL_IPU_REPLACEMENTS
        )
        logging.info("Saving results...")
        gvsipus.save()
        gvsipus.close()
        logging.info("Saving results done")
