"Heating tariffs for different periods"
from enum import Enum

from lib.datatypes import MonthYear


class HeatingTariff(float, Enum):
    """Heating prices for different periods"""

    T2020_1 = 3_217.63
    T2020_2 = 3_217.63
    T2021_1 = 3_276.63
    T2021_2 = 3_300
    T2022_1 = 3_457.63
    T2022_2 = 3_589.02

    @classmethod
    def get_tariff(cls, date: MonthYear):
        """Returns tariff for a particulat month/year"""
        match date.year:
            case 2020:
                if date.month < 7:
                    return HeatingTariff.T2020_1
                else:
                    return HeatingTariff.T2020_2
            case 2021:
                if date.month < 7:
                    return HeatingTariff.T2021_1
                else:
                    return HeatingTariff.T2021_2
            case 2022:
                if date.month < 7:
                    return HeatingTariff.T2022_1
                else:
                    return HeatingTariff.T2022_2
            case _:
                raise ValueError(f"No tariff specified for date: {date}")
