"""Custom exceptions"""


class NoServiceRow(Exception):
    """
    Raised when no row for a given month found service (heating or GVS)
    in accounts details table
    """


class NoAccountGvsRow(Exception):
    """
    Raised when no data found for a given account in GVS file
    """


class ZeroServiceReacuralRow(Exception):
    """
    Raised when reaccural for a service is zero
    """


class ZeroDataResultRow(Exception):
    """
    Raised when result row has only Zero values
    """


class NoAddressRow(Exception):
    """
    Raised when address is understood but not found in table
    """
