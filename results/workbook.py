"Workbook of results"

import logging
import os
import shutil

from openpyxl import load_workbook

from lib.helpers import BaseWorkBook
from results import ResultSheet
from results.calculations import BaseResultRow


class ResultWorkBookSheet:
    "Single workbook sheet of the result table"

    def __init__(self, sheet) -> None:
        self.sheet = sheet

    def add_row(self, row: BaseResultRow):
        "Adds row to table"
        self.sheet.append(row.as_list())


class ResultWorkBook(BaseWorkBook):
    """Table of results"""

    def __init__(self, base_dir: str, conf: dict) -> None:
        self.base_dir = base_dir
        self.conf = conf
        file_name = conf["result_file"]
        self.header_row = int(conf["calculations_header_row"])
        self.file_name_full = os.path.join(self.base_dir, file_name)
        template_name_full = os.path.join(
            os.path.dirname(self.base_dir), conf["result_template"]
        )
        logging.info("Initialazing result table %s ...", self.file_name_full)
        shutil.copyfile(template_name_full, self.file_name_full)
        self.workbook = load_workbook(filename=self.file_name_full)
        self.calculations = ResultWorkBookSheet(self.workbook[ResultSheet.CALCULATIONS])
        self.accounts = ResultWorkBookSheet(self.workbook[ResultSheet.ACCOUNTS])

    def save(self) -> None:
        """Saves result table data to disk"""
        logging.info("Saving results table...")
        self.workbook.save(filename=self.file_name_full)
        logging.info("All done")
