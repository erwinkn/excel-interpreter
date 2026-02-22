from enum import IntEnum, auto
from typing import Iterable, Union
from datetime import date, datetime, time, timedelta


from excel_interpreter.errors import CoercionError
from openpyxl.utils.datetime import (
    to_ISO8601,
    from_ISO8601,
    from_excel,
    to_excel,
    WINDOWS_EPOCH,
)


# The order of the flags is very deliberate here: in comparisons, booleans >
# text > numbers | dates. This ordering allows us to use ExcelType to directly handle that
class ExcelType(IntEnum):
    EMPTY = auto()
    NUMBER = auto()
    DATE = auto()
    TEXT = auto()
    BOOLEAN = auto()
    ARRAY = auto()


ScalarExcelValue = (
    None | int | float | str | bool | date | datetime | time | timedelta
)  # | time
ExcelValue = Union[ScalarExcelValue, "list[ExcelValue]"]


def excel_type(value: ExcelValue) -> ExcelType:
    """Return the ExcelType for a given ExcelValue."""
    if value is None:
        return ExcelType.EMPTY
    if isinstance(value, bool):
        return ExcelType.BOOLEAN
    if isinstance(value, (int, float)):
        return ExcelType.NUMBER
    if isinstance(value, str):
        return ExcelType.TEXT
    if isinstance(value, (datetime, time, timedelta)):
        return ExcelType.DATE
    if isinstance(value, list):
        return ExcelType.ARRAY
    raise CoercionError(f"Unknown Excel type for value: {value}")


def parse_number(val: str):
    is_float = ("." in val) or ("e" in val) or ("E" in val)
    return float(val) if is_float else int(val)


def coerce_to_number(val: ExcelValue, epoch=WINDOWS_EPOCH) -> float:
    """Convert an ExcelValue to a number following Excel semantics."""
    if val is None:
        return 0
    if isinstance(val, bool):
        return int(val)
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, str):
        try:
            return parse_number(val)
        except ValueError:
            raise CoercionError(f"Cannot convert text '{val}' to number")
    if isinstance(val, (datetime, date, time, timedelta)):
        return to_excel(val, epoch=epoch)
    if isinstance(val, list):
        raise CoercionError("Cannot convert array to number")
    raise CoercionError(f"Cannot convert {val} to number")


# Useful for math operations where we let Python handle the compatibility between the different types
def coerce_to_date_or_number(
    val: ExcelValue,
) -> int | float | datetime | time | timedelta:
    if val is None:
        return 0
    if isinstance(val, bool):
        return int(val)
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, (datetime, time, timedelta)):
        return val
    if isinstance(val, str):
        # from_ISO8601 returns None for the empty string & doesn't fail
        if not val:
            raise CoercionError("Cannot convert empty string to number or date")
        # Try as ISO format date
        try:
            return from_ISO8601(val)
        except ValueError:
            try:
                return parse_number(val)
            except ValueError:
                raise CoercionError(f"Cannot convert text '{val}' to number or date")
    if isinstance(val, list):
        raise CoercionError("Cannot convert array to number or date")
    raise CoercionError(f"Cannot convert {val} to number or date")


def coerce_to_bool(value: ExcelValue) -> bool:
    """Convert an ExcelValue to boolean following Excel semantics."""
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        if value.upper() in ("TRUE", "FALSE"):
            return value.upper() == "TRUE"
        raise CoercionError(f"Cannot convert text '{value}' to boolean")
    if isinstance(value, list):
        raise CoercionError("Cannot convert array to boolean")
    raise CoercionError(f"Cannot convert {value} to boolean")


def coerce_to_text(value: ExcelValue) -> str:
    """Convert an ExcelValue to text following Excel semantics."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value).upper()
    if isinstance(value, (int, float)):
        # TODO: Apply Excel's number formatting rules
        return str(value)
    if isinstance(value, str):
        return value
    if isinstance(value, (date, datetime)):
        # It's common to use a space for readability, Excel also does it
        return to_ISO8601(value).replace("T", " ")
    if isinstance(value, timedelta):
        hours = 24 * value.days + value.seconds // 3600
        minutes = (value.seconds // 60) % 60
        seconds = value.seconds % 60
        return f"{hours}:{minutes:02d}:{seconds:02d}"
    if isinstance(value, list):
        raise CoercionError("Cannot convert array to text")
    raise CoercionError(f"Cannot convert {value} to text")


def coerce_to_date(value: ExcelValue, epoch=WINDOWS_EPOCH) -> datetime:
    """Convert an ExcelValue to a date/time value following Excel semantics."""
    if isinstance(value, timedelta):
        # Convert timedelta to datetime using the epoch
        return epoch + value
    if isinstance(value, datetime):
        return value
    if isinstance(value, time):
        # Combine with epoch date
        return datetime.combine(epoch.date(), value)
    if isinstance(value, (int, float)):
        # Convert Excel serial number to datetime
        return from_excel(float(value), epoch=epoch)
    if isinstance(value, str):
        try:
            # Try as ISO format date
            return from_ISO8601(value)
        except ValueError:
            raise CoercionError(f"Cannot convert text '{value}' to date")
    raise CoercionError(f"Cannot convert {value} to date")


# Aggregation functions, which try to coerce and ignore invalid or empty values
def aggregate_numbers(
    values: Iterable[ExcelValue], skip_empty=True, epoch=WINDOWS_EPOCH
) -> list[float]:
    result: list[float] = []
    for val in values:
        is_empty = val is None or (isinstance(val, str) and len(val) == 0)
        if skip_empty and is_empty:
            continue
        try:
            result.append(coerce_to_number(val, epoch=epoch))
        except CoercionError:
            continue
    return result


def aggregate_booleans(values: Iterable[ExcelValue], skip_empty=True) -> list[bool]:
    result: list[bool] = []
    for val in values:
        is_empty = val is None or (isinstance(val, str) and len(val) == 0)
        if skip_empty and is_empty:
            continue
        try:
            result.append(coerce_to_bool(val))
        except CoercionError:
            continue
    return result


def is_error(value: ExcelValue) -> bool:
    """Return True if the value is an Excel error value."""
    return isinstance(value, str) and value.startswith("#")
