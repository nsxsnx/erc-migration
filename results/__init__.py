"Generation of workbook with the results of all calculations"

from enum import StrEnum


class ResultSheet(StrEnum):
    "Names of sheets in the results workbook"
    ACCOUNTS = "Лицевые счета и помещения"
    CALCULATIONS = "Расчеты"
    PEOPLE = "Жильцы и собственники"
