import logging
from datetime import date, datetime, time, timedelta
from enum import IntEnum
from typing import Any, Callable, Optional, Type, TypeVar, Union, cast

from openpyxl.utils.datetime import (
    WINDOWS_EPOCH,
    days_to_time,
    from_excel,
    to_excel,
)

from excel_interpreter.types import (
    ExcelType,
    ExcelValue,
    ScalarExcelValue,
    coerce_to_date_or_number,
    coerce_to_number,
    coerce_to_text,
    excel_type,
    is_error,
)

T = TypeVar("T", bound=ScalarExcelValue)
ListOrScalar = Union[T, list[T], list[list[T]]]


def vectorize(
    op: Callable[[ScalarExcelValue, ScalarExcelValue, datetime], Any],
    left: ExcelValue,
    right: ExcelValue,
    epoch: datetime = WINDOWS_EPOCH,
) -> ExcelValue:
    """
    Apply a scalar operation element-wise if either operand is an array.
    Handles broadcasting between 2D and 1D arrays.
    """
    # Handle 2D array operations with broadcasting
    if isinstance(left, list) and isinstance(left[0], list):
        if isinstance(right, list):
            if isinstance(right[0], list):
                # Both 2D arrays - dimensions must match exactly
                if len(left) != len(right) or len(left[0]) != len(right[0]):
                    raise ValueError("2D array dimensions do not match")
                return [
                    [vectorize(op, l, r, epoch) for l, r in zip(lrow, rrow)]
                    for lrow, rrow in zip(left, right)
                ]
            else:
                # 2D array * 1D array - try broadcasting
                if len(left) == len(right):
                    # Broadcast 1D array across columns
                    return [
                        [vectorize(op, r, col, epoch) for col in row]
                        for r, row in zip(right, left)
                    ]
                elif len(left[0]) == len(right):
                    # Broadcast 1D array across rows
                    return [
                        [vectorize(op, l, r, epoch) for l, r in zip(lrow, right)]
                        for lrow in left
                    ]
                else:
                    raise ValueError("Array dimensions do not match for broadcasting")
        else:
            # 2D array * scalar - broadcast scalar to all elements
            return [[vectorize(op, l, right, epoch) for l in lrow] for lrow in left]

    elif isinstance(right, list) and isinstance(right[0], list):
        if isinstance(left, list):
            # 1D array * 2D array - try broadcasting
            if len(right) == len(left):
                # Broadcast 1D array across columns
                return [
                    [vectorize(op, l, col, epoch) for col in row]
                    for l, row in zip(left, right)
                ]
            elif len(right[0]) == len(left):
                # Broadcast 1D array across rows
                return [
                    [vectorize(op, l, r, epoch) for l, r in zip(left, rrow)]
                    for rrow in right
                ]
            else:
                raise ValueError("Array dimensions do not match for broadcasting")
        else:
            # scalar * 2D array - broadcast scalar to all elements
            return [[vectorize(op, left, r, epoch) for r in rrow] for rrow in right]

    # Handle 1D array operations (no change to existing logic)
    elif isinstance(left, list) and isinstance(right, list):
        if len(left) != len(right):
            raise ValueError("Array dimensions do not match")
        return [vectorize(op, l, r, epoch) for l, r in zip(left, right)]
    elif isinstance(left, list):
        return [vectorize(op, l, right, epoch) for l in left]
    elif isinstance(right, list):
        return [vectorize(op, left, r, epoch) for r in right]
    else:
        return op(left, right, epoch)


# Define value categories for date operations
class DateCategory(IntEnum):
    OTHER = 0
    INTERVAL = 1  # time, timedelta, int, float
    DATE = 2  # date, datetime


def ensure_is_number(
    value: int | float | datetime | date | time | timedelta, epoch=WINDOWS_EPOCH
) -> int | float:
    if isinstance(value, (date, datetime, time, timedelta)):
        value = cast(int | float, to_excel(value, epoch=epoch))
    return value


def get_value_category(value: Any) -> DateCategory:
    """
    Categorize a value based on its type.
    - Dates: date, datetime
    - Intervals: time, timedelta, int, float
    - Other: anything else
    """
    if isinstance(value, (datetime, date)):
        return DateCategory.DATE
    elif isinstance(value, (time, timedelta, int, float)):
        return DateCategory.INTERVAL
    else:
        return DateCategory.OTHER


add_output_type = {
    (DateCategory.DATE, DateCategory.INTERVAL): DateCategory.DATE,
    (DateCategory.INTERVAL, DateCategory.DATE): DateCategory.DATE,
    (DateCategory.INTERVAL, DateCategory.INTERVAL): DateCategory.INTERVAL,
}


# Useful for math operations where we let Python handle the compatibility between the different types
def add_scalar(
    left: ScalarExcelValue,
    right: ScalarExcelValue,
    epoch=WINDOWS_EPOCH,
) -> ScalarExcelValue:
    # Propagate errors
    if is_error(left):
        return left
    if is_error(right):
        return right

    # Fast case
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return left + right

    l = coerce_to_date_or_number(left)
    r = coerce_to_date_or_number(right)

    # Fast case nb2
    if isinstance(l, (int, float)) and isinstance(r, (int, float)):
        return l + r

    # If we're here, we're dealing with date operations.
    ltype = get_value_category(l)
    rtype = get_value_category(r)

    # Use default epoch if None is provided
    result = ensure_is_number(l, epoch) + ensure_is_number(r, epoch)

    output_type = add_output_type.get((ltype, rtype))
    if output_type == DateCategory.DATE:
        return from_excel(result, epoch=epoch)
    if output_type == DateCategory.INTERVAL:
        return from_excel(result, epoch=epoch, timedelta=True)

    # If no conversion, we're dealing with something weird
    logging.warning(
        "Returning numerical result from date operations. Due to Excel's peculiarities, the result is not guaranteed to be correct. You should triple check whether you really intend to perform this operation."
    )
    return result


sub_output_type = {
    (DateCategory.DATE, DateCategory.DATE): DateCategory.INTERVAL,
    (DateCategory.DATE, DateCategory.INTERVAL): DateCategory.DATE,
    (DateCategory.INTERVAL, DateCategory.INTERVAL): DateCategory.INTERVAL,
}


def subtract_scalar(
    left: ScalarExcelValue,
    right: ScalarExcelValue,
    epoch: datetime = WINDOWS_EPOCH,
) -> ScalarExcelValue:
    # Propagate errors
    if is_error(left):
        return left
    if is_error(right):
        return right

    # Fast case
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return left - right

    l = coerce_to_date_or_number(left)
    r = coerce_to_date_or_number(right)

    # Fast case nb2
    if isinstance(l, (int, float)) and isinstance(r, (int, float)):
        return l - r

    # If we're here, we're dealing with date operations.
    ltype = get_value_category(l)
    rtype = get_value_category(r)

    result = ensure_is_number(l, epoch) - ensure_is_number(r, epoch)

    output_type = sub_output_type.get((ltype, rtype))
    if output_type == DateCategory.DATE:
        return from_excel(result, epoch=epoch)
    if output_type == DateCategory.INTERVAL:
        return from_excel(result, epoch=epoch, timedelta=True)

    # If no conversion, we're dealing with something weird
    logging.warning(
        "Returning numerical result from date operations. Due to Excel's peculiarities, the result is not guaranteed to be correct. You should triple check whether you really intend to perform this operation."
    )
    return result


mul_output_type = {
    (DateCategory.INTERVAL, DateCategory.INTERVAL): DateCategory.INTERVAL,
}


# Multiplication and division convert dates to numbers
def multiply_scalar(
    left: ScalarExcelValue,
    right: ScalarExcelValue,
    epoch: datetime = WINDOWS_EPOCH,
) -> ScalarExcelValue:
    # Propagate errors
    if is_error(left):
        return left
    if is_error(right):
        return right

    # Fast case
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return left * right

    l = coerce_to_date_or_number(left)
    r = coerce_to_date_or_number(right)

    # Fast case nb2
    if isinstance(l, (int, float)) and isinstance(r, (int, float)):
        return l * r

    # If we're here, we're dealing with date operations.
    ltype = get_value_category(l)
    rtype = get_value_category(r)

    result = ensure_is_number(l, epoch) * ensure_is_number(r, epoch)

    output_type = mul_output_type.get((ltype, rtype))

    if output_type == DateCategory.INTERVAL:
        return from_excel(result, epoch=epoch, timedelta=True)

    # If no conversion, we're dealing with something weird
    logging.warning(
        "Returning numerical result from date operations. Due to Excel's peculiarities, the result is not guaranteed to be correct. You should triple check whether you really intend to perform this operation."
    )
    return result


div_output_type = {
    (
        DateCategory.INTERVAL,
        DateCategory.INTERVAL,
    ): None,  # Special case, may return scalar
}


def divide_scalar(
    left: ScalarExcelValue,
    right: ScalarExcelValue,
    epoch: datetime = WINDOWS_EPOCH,
) -> ScalarExcelValue:
    # Propagate errors
    if is_error(left):
        return left
    if is_error(right):
        return right

    # Fast case
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return left / right

    l = coerce_to_date_or_number(left)
    r = coerce_to_date_or_number(right)

    # Fast case nb2
    if isinstance(l, (int, float)) and isinstance(r, (int, float)):
        return l / r

    # If we're here, we're dealing with date operations.
    ltype = get_value_category(l)
    rtype = get_value_category(r)

    result = ensure_is_number(l, epoch) / ensure_is_number(r, epoch)

    # Special case for division of same types (returns a scalar)
    if isinstance(l, time) and isinstance(r, time):
        return result
    if isinstance(l, timedelta) and isinstance(r, timedelta):
        return result

    # For interval / number, preserve the interval type
    if ltype == DateCategory.INTERVAL and isinstance(l, time):
        return days_to_time(timedelta(days=result % 1))
    elif ltype == DateCategory.INTERVAL and isinstance(l, timedelta):
        return from_excel(result, epoch=epoch, timedelta=True)

    # If no conversion, we're dealing with something weird
    logging.warning(
        "Returning numerical result from date operations. Due to Excel's peculiarities, the result is not guaranteed to be correct. You should triple check whether you really intend to perform this operation."
    )
    return result


def power_scalar(
    left: ScalarExcelValue,
    right: ScalarExcelValue,
    epoch: Optional[datetime] = WINDOWS_EPOCH,
) -> ScalarExcelValue:
    # Propagate errors
    if is_error(left):
        return left
    if is_error(right):
        return right

    actual_epoch = epoch if epoch is not None else WINDOWS_EPOCH
    return coerce_to_number(left, actual_epoch) ** coerce_to_number(right, actual_epoch)


def eq_scalar(
    left: ScalarExcelValue,
    right: ScalarExcelValue,
    epoch: Optional[datetime] = WINDOWS_EPOCH,
) -> bool:
    # Empty values are treated as zeros for comparisons
    if left is None:
        left = 0
    if right is None:
        right = 0

    actual_epoch = epoch if epoch is not None else WINDOWS_EPOCH
    if isinstance(left, (datetime, date, time, timedelta)):
        left = to_excel(left, epoch=actual_epoch)
    if isinstance(right, (datetime, date, time, timedelta)):
        right = to_excel(right, epoch=actual_epoch)

    # String comparisons are case-insensitive in Excel
    if isinstance(left, str) and isinstance(right, str):
        return left.lower() == right.lower()

    return excel_type(left) == excel_type(right) and left == right


def neq_scalar(
    left: ScalarExcelValue,
    right: ScalarExcelValue,
    epoch: Optional[datetime] = WINDOWS_EPOCH,
) -> bool:
    return not eq_scalar(left, right, epoch)


COMPARISON_TYPE_PRIORITY: dict[ExcelType, int] = {
    ExcelType.BOOLEAN: 3,
    ExcelType.TEXT: 2,
    ExcelType.NUMBER: 1,
    ExcelType.DATE: 1,
}


def lt_scalar(
    left: ScalarExcelValue,
    right: ScalarExcelValue,
    epoch: Optional[datetime] = WINDOWS_EPOCH,
) -> bool:
    # Empty values are treated as zeros or false for comparisons
    if left is None:
        left = 0
    if right is None:
        right = 0

    # Optimization, early exit
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return left < right

    actual_epoch = epoch if epoch is not None else WINDOWS_EPOCH
    # Dates can be treated as numbers
    if isinstance(left, (datetime, date, time, timedelta)):
        left = to_excel(left, epoch=actual_epoch)
    if isinstance(right, (datetime, date, time, timedelta)):
        right = to_excel(right, epoch=actual_epoch)

    # Priority order:
    # Booleans > strings > numbers | dates
    # (don't ask me why, I just spent way too much time testing this)
    lpriority = COMPARISON_TYPE_PRIORITY[excel_type(left)]
    rpriority = COMPARISON_TYPE_PRIORITY[excel_type(right)]
    if lpriority != rpriority:
        return lpriority < rpriority

    # String comparisons are case insensitive
    if isinstance(left, str) and isinstance(right, str):
        return left.lower() < right.lower()

    # Otherwise, compare values (only numbers and booleans left here)
    # At this point we know both values are of the same type and are comparable
    return cast(bool | float, left) < cast(bool | float, right)


def gt_scalar(
    left: ScalarExcelValue,
    right: ScalarExcelValue,
    epoch: Optional[datetime] = WINDOWS_EPOCH,
) -> bool:
    return lt_scalar(right, left, epoch)  # flip the arguments


def lte_scalar(
    left: ScalarExcelValue,
    right: ScalarExcelValue,
    epoch: Optional[datetime] = WINDOWS_EPOCH,
) -> bool:
    lt_result = lt_scalar(left, right, epoch)
    eq_result = eq_scalar(left, right, epoch)
    return lt_result or eq_result


def gte_scalar(
    left: ScalarExcelValue,
    right: ScalarExcelValue,
    epoch: Optional[datetime] = WINDOWS_EPOCH,
) -> bool:
    return not lt_scalar(left, right, epoch)


def concatenate_scalar(
    left: ScalarExcelValue, right: ScalarExcelValue, epoch: Optional[datetime] = None
) -> str:
    return coerce_to_text(left) + coerce_to_text(right)


# Operator functions that use vectorize to handle arrays


def add(left: ExcelValue, right: ExcelValue, epoch=WINDOWS_EPOCH) -> ExcelValue:
    return vectorize(add_scalar, left, right, epoch)


def subtract(left: ExcelValue, right: ExcelValue, epoch=WINDOWS_EPOCH) -> ExcelValue:
    return vectorize(subtract_scalar, left, right, epoch)


def multiply(left: ExcelValue, right: ExcelValue, epoch=WINDOWS_EPOCH) -> ExcelValue:
    return vectorize(multiply_scalar, left, right, epoch)


def divide(left: ExcelValue, right: ExcelValue, epoch=WINDOWS_EPOCH) -> ExcelValue:
    return vectorize(divide_scalar, left, right, epoch)


def power(left: ExcelValue, right: ExcelValue, epoch=WINDOWS_EPOCH) -> ExcelValue:
    """Exponentiation operator (^) that handles arrays."""
    return vectorize(power_scalar, left, right, epoch)


def eq(left: ExcelValue, right: ExcelValue, epoch=WINDOWS_EPOCH) -> ExcelValue:
    return vectorize(eq_scalar, left, right, epoch)


def neq(left: ExcelValue, right: ExcelValue, epoch=WINDOWS_EPOCH) -> ExcelValue:
    return vectorize(neq_scalar, left, right, epoch)


def lt(left: ExcelValue, right: ExcelValue, epoch=WINDOWS_EPOCH) -> ExcelValue:
    return vectorize(lt_scalar, left, right, epoch)


def gt(left: ExcelValue, right: ExcelValue, epoch=WINDOWS_EPOCH) -> ExcelValue:
    return vectorize(gt_scalar, left, right, epoch)


def lte(left: ExcelValue, right: ExcelValue, epoch=WINDOWS_EPOCH) -> ExcelValue:
    return vectorize(lte_scalar, left, right, epoch)


def gte(left: ExcelValue, right: ExcelValue, epoch=WINDOWS_EPOCH) -> ExcelValue:
    return vectorize(gte_scalar, left, right, epoch)


def concatenate(left: ExcelValue, right: ExcelValue) -> ExcelValue:
    return vectorize(
        concatenate_scalar, left, right, WINDOWS_EPOCH
    )  # epoch doesn't matter here
