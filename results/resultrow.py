"Single row of results"


from decimal import Decimal
from typing import Any


class ResultRow:
    "Base class for all results"
    max_fields: int

    def __init__(self, max_fields) -> None:
        self.max_fields = max_fields
        for ind in range(self.max_fields):
            setattr(self, f"f{ind:02d}", None)

    def set_field(self, ind: int, value: str | int | float | Decimal | None = None):
        "Field setter by field number"
        setattr(self, f"f{ind:02d}", value)

    def get_field(self, ind: int) -> str | None:
        "Field getter by field number"
        return getattr(self, f"f{ind:02d}")

    def as_list(self) -> list[Any]:
        "Returns list of all fields"
        result = []
        for ind in range(self.max_fields):
            result.append(getattr(self, f"f{ind:02d}"))
        return result
