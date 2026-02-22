import math
import re
from datetime import datetime, time, timedelta
from typing import Any, Callable, Optional, ParamSpec, cast, overload

import numpy as np

from excel_interpreter.errors import CoercionError, ExcelFunctionError
from excel_interpreter.operators import (
    eq_scalar,
    gt_scalar,
    gte_scalar,
    lt_scalar,
    lte_scalar,
    neq_scalar,
    power,
)
from excel_interpreter.types import (
    ExcelValue,
    ScalarExcelValue,
    aggregate_booleans,
    aggregate_numbers,
    coerce_to_bool,
    coerce_to_date,
    coerce_to_number,
    coerce_to_text,
    is_error,
)
from excel_interpreter.utils import array_shape

P = ParamSpec("P")

EXCEL_FUNCTIONS: dict[str, Callable[..., ExcelValue]] = {}


@overload
def excel_fn(
    fn: Callable[P, ExcelValue], *, name: Optional[str] = None
) -> Callable[P, ExcelValue]: ...
@overload
def excel_fn(
    fn: None = None, *, name: Optional[str] = None
) -> Callable[[Callable[P, ExcelValue]], Callable[P, ExcelValue]]: ...


def excel_fn(
    fn: Callable[P, ExcelValue] | None = None,
    *,
    name: Optional[str] = None,
) -> Any:
    """Decorator to register a function as an Excel function."""

    def decorator(fn: Callable[P, ExcelValue]) -> Callable[P, ExcelValue]:
        # If used on a staticmethod/classmethod, unwrap for registration but
        # return the original descriptor to preserve method semantics.
        if isinstance(fn, staticmethod):
            underlying = fn.__func__
            reg_name = name or underlying.__name__
            EXCEL_FUNCTIONS[reg_name] = underlying
            setattr(underlying, "_excel_fn_registered", True)
            setattr(underlying, "_excel_fn_name", reg_name)
            return fn  # type: ignore
        if isinstance(fn, classmethod):
            underlying = fn.__func__
            reg_name = name or underlying.__name__
            EXCEL_FUNCTIONS[reg_name] = underlying
            setattr(underlying, "_excel_fn_registered", True)
            setattr(underlying, "_excel_fn_name", reg_name)
            return fn  # type: ignore

        reg_name = name or fn.__name__
        EXCEL_FUNCTIONS[reg_name] = fn
        setattr(fn, "_excel_fn_registered", True)
        setattr(fn, "_excel_fn_name", reg_name)
        return fn

    if fn:
        return decorator(fn)
    else:
        return decorator


def to_1d_array(value: ExcelValue) -> list:
    """Convert an Excel value to a 1D array.

    Handles:
    - Single values (converted to [value])
    - 1D arrays (returned as is)
    - 2D arrays (flattened to 1D)
    """
    if not isinstance(value, list):
        return [value]
    if not value:
        return []
    if isinstance(value[0], list):
        # Flatten 2D array
        return [item for sublist in value for item in cast(list[ExcelValue], sublist)]
    return value


def _match_pattern(value: ScalarExcelValue, pattern: str) -> bool:
    """Match a value against an Excel pattern, supporting wildcards and comparisons."""
    # Handle empty values
    if value is None:
        return pattern == "" or pattern == '""' or pattern == "0"

    value = coerce_to_text(value)

    # Map of operators to their scalar comparison functions
    op_funcs = {
        ">=": gte_scalar,
        "<=": lte_scalar,
        "<>": neq_scalar,
        ">": gt_scalar,
        "<": lt_scalar,
        "=": eq_scalar,
    }

    # Handle comparison operators
    for op, func in op_funcs.items():
        if pattern.startswith(op):
            try:
                pattern = pattern[len(op) :].strip()
                if len(pattern) == 0:
                    raise NotImplementedError("Unsupported: empty pattern")
                return func(value, pattern)
            except (ValueError, CoercionError):
                # If comparison fails, fall back to string comparison
                break

    # Handle wildcards (* and ?)
    if "*" in pattern or "?" in pattern:
        # Convert Excel wildcards to regex patterns
        regex_pattern = (
            "^" + re.escape(pattern).replace("\\*", ".*").replace("\\?", ".") + "$"
        )
        return bool(re.match(regex_pattern, coerce_to_text(value), re.IGNORECASE))

    # For exact match, use the equality operator
    result = eq_scalar(value, pattern)
    return result


def flatten_args(*args: ExcelValue) -> list[ScalarExcelValue]:
    """Flatten multiple Excel function arguments into a single list of non-array values."""
    result: list[ScalarExcelValue] = []
    for arg in args:
        if isinstance(arg, list):
            result.extend(flatten_args(*arg))
        else:
            # Flat arg
            result.append(arg)
    return result


class ExcelFunctions:
    """Collection of Excel function implementations."""

    @staticmethod
    def SUM(*args: ExcelValue) -> ExcelValue:
        """Sum of arguments, handling arrays and ranges."""
        total = sum(aggregate_numbers(flatten_args(*args)))
        return total

    @staticmethod
    def AVERAGE(*args: ExcelValue) -> ExcelValue:
        """Average of arguments, ignoring empty cells."""
        nums = aggregate_numbers(flatten_args(*args))
        return (sum(nums) / len(nums)) if nums else 0

    @staticmethod
    def MAX(*args: ExcelValue) -> ExcelValue:
        """Return the maximum value, ignoring empty cells."""
        nums = aggregate_numbers(flatten_args(*args))
        return max(nums) if nums else 0

    @staticmethod
    def MIN(*args: ExcelValue) -> ExcelValue:
        """Return the minimum value, ignoring empty cells."""
        nums = aggregate_numbers(flatten_args(*args))
        return min(nums) if nums else 0

    @staticmethod
    def IF(
        condition: ExcelValue, true_value: ExcelValue, false_value: ExcelValue
    ) -> ExcelValue:
        """Return true_value if condition is True, false_value otherwise."""
        # First check if condition is an error
        if is_error(condition):
            return condition

        # Handle array inputs by vectorizing the operation
        if isinstance(condition, list):
            # For 2D arrays
            if isinstance(condition[0], list):
                return [
                    [
                        (
                            true_value[i][j]
                            if isinstance(true_value, list)
                            else true_value
                        )
                        if coerce_to_bool(cond)
                        else (
                            false_value[i][j]
                            if isinstance(false_value, list)
                            else false_value
                        )
                        for j, cond in enumerate(row)
                    ]
                    for i, row in enumerate(condition)
                ]
            # For 1D arrays
            return [
                (true_value[i] if isinstance(true_value, list) else true_value)
                if coerce_to_bool(cond)
                else (false_value[i] if isinstance(false_value, list) else false_value)
                for i, cond in enumerate(condition)
            ]

        # Handle scalar case
        try:
            result = coerce_to_bool(condition)
            return true_value if result else false_value
        except (ValueError, TypeError):
            return "#VALUE!"

    @staticmethod
    def IFERROR(value: ExcelValue, error_value: ExcelValue) -> ExcelValue:
        """Return error_value if value is an error, value otherwise."""
        return (
            error_value if isinstance(value, str) and value.startswith("#") else value
        )

    @staticmethod
    def AND(*args: ExcelValue) -> ExcelValue:
        """Return True if all arguments are True."""
        return all(aggregate_booleans(flatten_args(*args)))

    @staticmethod
    def OR(*args: ExcelValue) -> ExcelValue:
        """Return True if any argument is True."""
        return any(aggregate_booleans(flatten_args(*args)))

    @staticmethod
    def TEXT(value: ExcelValue, format_text: ExcelValue) -> ExcelValue:
        """Convert a value to text using a (limited) Excel format string.

        This is intentionally minimal and currently supports the patterns we
        commonly see in SharePoint-hosted workbooks where cached formula results
        are not present (e.g. TEXT(A1, "0"), TEXT(A1, "00"), TEXT(A1, "0.00"),
        TEXT(A1, "mmdd"), TEXT(A1, "hh")).
        """
        if is_error(value):
            return value

        if value is None:
            return ""

        fmt = coerce_to_text(format_text)
        fmt = fmt.strip().strip('"')
        fmt_lower = fmt.lower()

        # Support the date/time formats used by pilot workbook sample-id formulas.
        if fmt_lower == "mmdd":
            if isinstance(value, datetime):
                return value.strftime("%m%d")
            return coerce_to_date(value).strftime("%m%d")

        if fmt_lower == "hh":
            if isinstance(value, datetime):
                return value.strftime("%H")
            if isinstance(value, time):
                return f"{value.hour:02d}"
            return coerce_to_date(value).strftime("%H")

        # Date/time formatting is not implemented; fall back to Excel-like text.
        if isinstance(value, (datetime, time)):
            return coerce_to_text(value)

        num = coerce_to_number(value)

        if not fmt or fmt == "general":
            return coerce_to_text(num)

        # Handle simple integer/decimal zero formats like 0, 00, 0.0, 0.00
        if all(ch in "0." for ch in fmt) and fmt.count(".") <= 1:
            if "." in fmt:
                int_part, dec_part = fmt.split(".", 1)
                decimals = len(dec_part)
                width = len(int_part)
                rounded = round(float(num), decimals)
                # Ensure we keep trailing zeros per decimals
                s = f"{rounded:.{decimals}f}"
                if width > 1:
                    if "." in s:
                        left, right = s.split(".", 1)
                        left = left.zfill(width)
                        s = f"{left}.{right}"
                    else:
                        s = s.zfill(width)
                return s

            width = len(fmt)
            as_int = int(round(float(num)))
            return str(as_int).zfill(width) if width > 1 else str(as_int)

        # Fallback: best-effort string coercion.
        return coerce_to_text(value)

    @staticmethod
    def NOT(value: ExcelValue) -> ExcelValue:
        """Return the logical NOT of the argument."""
        try:
            return not coerce_to_bool(value)
        except ValueError:
            return True  # NOT of invalid value is TRUE in Excel

    @staticmethod
    def PI() -> ExcelValue:
        """Return the value of π."""
        return math.pi

    @staticmethod
    def ABS(x: ExcelValue) -> ExcelValue:
        """Return the absolute value."""
        return abs(coerce_to_number(x))

    @staticmethod
    def SQRT(x: ExcelValue) -> ExcelValue:
        """Return the square root."""
        num = coerce_to_number(x)
        return math.sqrt(num)

    @staticmethod
    def POWER(base: ExcelValue, exponent: ExcelValue) -> ExcelValue:
        """Excel POWER function that matches ^ operator behavior."""
        return power(base, exponent)

    @staticmethod
    def TRUNC(value: ExcelValue, num_digits: ExcelValue = 0) -> ExcelValue:
        """Truncate a number to the specified number of digits."""
        num = coerce_to_number(value)
        digits = int(coerce_to_number(num_digits))
        multiplier = 10**digits
        return int(num * multiplier) / multiplier

    @staticmethod
    def CEILING(number: ExcelValue, significance: ExcelValue = 1) -> ExcelValue:
        """Round up to the nearest multiple of significance."""
        num = coerce_to_number(number)
        sig = coerce_to_number(significance)

        if num == 0:
            return 0
        if sig == 0:
            raise ExcelFunctionError("Significance cannot be zero")
        if num < 0 and sig > 0:
            return -math.floor(-num / sig) * sig
        if num > 0 and sig < 0:
            raise ExcelFunctionError(
                "Positive number with negative significance is not allowed"
            )
        return sig * math.ceil(num / sig)

    @staticmethod
    def EXP(x: ExcelValue) -> ExcelValue:
        """Return e raised to the power of x."""
        return math.exp(coerce_to_number(x))

    @staticmethod
    def LN(x: ExcelValue) -> ExcelValue:
        """Return the natural logarithm of x."""
        num = coerce_to_number(x)
        if num <= 0:
            raise ExcelFunctionError("LN requires positive input")
        return math.log(num)

    @staticmethod
    def NPV(rate: ExcelValue, *values: ExcelValue) -> ExcelValue:
        """Calculate Net Present Value from a rate and series of payments."""
        r = coerce_to_number(rate)
        if r == -1:
            raise ExcelFunctionError("Rate cannot be -100%")

        flatargs = flatten_args(*values)
        if len(flatargs) < 2:
            raise ExcelFunctionError("Not enough arguments for NPV")
        npv = 0
        for i, val in enumerate(flatargs):
            try:
                num = coerce_to_number(val)
                npv += num / ((1 + r) ** (i + 1))
            except CoercionError:
                # Excel ignores text values in NPV
                continue
        return npv

    @staticmethod
    def IRR(*values: ExcelValue):
        import numpy_financial as npf

        return npf.irr(aggregate_numbers(flatten_args(*values)))

    @staticmethod
    def DATE(year: ExcelValue, month: ExcelValue, day: ExcelValue) -> ExcelValue:
        """Create a date from year, month, and day components."""
        return datetime(
            int(coerce_to_number(year)),
            int(coerce_to_number(month)),
            int(coerce_to_number(day)),
        )

    @staticmethod
    def TIME(hour: ExcelValue, minute: ExcelValue, second: ExcelValue) -> ExcelValue:
        """Create a time from hour, minute, and second components."""
        return time(
            hour=int(coerce_to_number(hour)) % 24,
            minute=int(coerce_to_number(minute)) % 60,
            second=int(coerce_to_number(second)) % 60,
        )

    @staticmethod
    def YEAR(date_value: ExcelValue) -> ExcelValue:
        """Return the year component of a date."""
        date = coerce_to_date(date_value)
        return date.year

    @staticmethod
    def MONTH(date_value: ExcelValue) -> ExcelValue:
        """Return the month component of a date."""
        date = coerce_to_date(date_value)
        return date.month

    @staticmethod
    def DAY(date_value: ExcelValue) -> ExcelValue:
        """Return the day component of a date."""
        date = coerce_to_date(date_value)
        return date.day

    @staticmethod
    def HOUR(date_value: ExcelValue) -> ExcelValue:
        """Return the hour component of a time."""
        date = coerce_to_date(date_value)
        return date.hour

    @staticmethod
    def MINUTE(date_value: ExcelValue) -> ExcelValue:
        """Return the minute component of a time."""
        date = coerce_to_date(date_value)
        return date.minute

    @staticmethod
    def ISBLANK(value: ExcelValue) -> ExcelValue:
        return value is None

    @staticmethod
    def CONCAT(*args: ExcelValue) -> ExcelValue:
        return "".join(coerce_to_text(val) for val in flatten_args(*args))

    @staticmethod
    def CONCATENATE(*args: ExcelValue) -> ExcelValue:
        return ExcelFunctions.CONCAT(*args)

    @staticmethod
    def REGEXEXTRACT(
        text: ExcelValue,
        pattern: ExcelValue,
        return_mode: ExcelValue = 0,
        case_sensitivity: ExcelValue = 0,
    ) -> ExcelValue:
        """Extract text using a regular expression.

        Args:
            text: The text to search within.
            pattern: The regex pattern to match (PCRE-like; Python `re`).
            return_mode: 0 (default) = first match; 1 = all matches; 2 = capturing groups from first match.
            case_sensitivity: 0 (default) case sensitive, 1 = case insensitive.

        Returns:
            A text value (for return_mode 0) or an array of text values (for modes 1 and 2).
            Returns "#N/A" if no match is found. Returns "#VALUE!" for invalid inputs.
        """
        # Reject array inputs for text to avoid ambiguous shapes
        if isinstance(text, list):
            raise ExcelFunctionError("REGEXTRACT doesn't support ranges")

        try:
            text_str = coerce_to_text(text)
            pattern_str = coerce_to_text(pattern)
        except Exception:
            return "#VALUE!"

        # Parse optional parameters
        try:
            rm = int(coerce_to_number(return_mode)) if return_mode is not None else 0
        except Exception:
            return "#VALUE!"

        try:
            cs = (
                int(coerce_to_number(case_sensitivity))
                if case_sensitivity is not None
                else 0
            )
        except Exception:
            return "#VALUE!"

        flags = re.IGNORECASE if cs == 1 else 0

        try:
            regex = re.compile(pattern_str, flags)
        except re.error:
            return "#VALUE!"

        # Mode 0: first match
        if rm == 0:
            match = regex.search(text_str)
            if not match:
                return "#N/A"
            return match.group(0)

        # Mode 1: all matches (as array of full-match strings)
        if rm == 1:
            matches: list[str] = [m.group(0) for m in regex.finditer(text_str)]
            if not matches:
                return "#N/A"
            return cast("list[ExcelValue]", matches)

        # Mode 2: capturing groups from the first match
        if rm == 2:
            match = regex.search(text_str)
            if not match:
                return "#N/A"
            groups: list[str | None] = list(match.groups())
            if not groups:
                return "#N/A"
            normalized: list[str] = [g if g is not None else "" for g in groups]
            return cast("list[ExcelValue]", normalized)

        return "#VALUE!"

    @staticmethod
    def ROUNDUP(number: ExcelValue, num_digits: ExcelValue) -> ExcelValue:
        """Round a number up to the specified number of decimal places, with smart rounding to integer."""
        factor = 10**num_digits
        rounded = math.ceil(number * factor) / factor
        rounded = round(rounded, num_digits)

        # Smart rounding: if within tolerance of a whole number, return int
        if abs(rounded - round(rounded)) < (1 / factor):
            return int(round(rounded))

        return rounded

    @staticmethod
    def COUNTIF(values: ExcelValue, criteria: ExcelValue) -> ExcelValue:
        """
        Count cells in a range that meet the given criteria.

        Supports:
        - Exact matches: COUNTIF(range, "value")
        - Numeric comparisons: COUNTIF(range, ">10")
        - Wildcard text matching: COUNTIF(range, "a*")
        - Direct cell reference: COUNTIF(range, A1)
        """
        criteria_str = coerce_to_text(criteria)
        count = 0
        for value in flatten_args(values):
            if _match_pattern(value, criteria_str):
                count += 1

        return count

    @staticmethod
    def COUNTIFS(*args: ExcelValue) -> ExcelValue:
        assert len(args) % 2 == 0, "COUNTIFS requires an even number of arguments"
        if len(args) == 0:
            return 0
        values = []
        criteria = []
        shapes = []
        for n in range(0, len(args), 2):
            values.append(flatten_args(*args[n]))
            shapes.append(array_shape(args[n]))
            criteria.append(coerce_to_text(args[n + 1]))
        assert len(set(shapes)) == 1, (
            f"All ranges passed to COUNTIFS must have the same shape, got ranges of shapes: {shapes}"
        )
        total = 0
        for n in range(values[0]):
            matches_all = all(
                _match_pattern(values[i][n], criteria[i]) for i in range(len(values))
            )
            total += matches_all
        return total

    @staticmethod
    def VLOOKUP(
        lookup_value: ExcelValue,
        table_array: ExcelValue,
        col_index_num: ExcelValue,
        range_lookup: ExcelValue = True,
    ) -> ExcelValue:
        """
        Lookup a value in a table and return the corresponding value from a specified column.

        Args:
            lookup_value: The value to search for in the first column of table_array
            table_array: The range of cells to search in
            col_index_num: The column number in table_array from which to return a value
            range_lookup: If False, find exact match. If True, find closest match if exact not found
        """
        # Handle array inputs for lookup_value
        if isinstance(lookup_value, list):
            # If lookup_value is 2D array, convert to 1D
            if isinstance(lookup_value[0], list):
                lookup_value = [val for row in lookup_value for val in row]
            # Apply VLOOKUP to each value in the array
            results = [
                ExcelFunctions.VLOOKUP(val, table_array, col_index_num, range_lookup)
                for val in lookup_value
            ]
            return results

        # Convert lookup_value to text for comparison
        try:
            lookup_text = (
                coerce_to_text(lookup_value)
                if not isinstance(lookup_value, (int, float))
                else lookup_value
            )
        except Exception:
            return "#VALUE!"

        # Ensure table_array is a 2D array
        if not isinstance(table_array, list):
            table_array = [[table_array]]
        elif not isinstance(table_array[0], list):
            table_array = [table_array]

        # Get the column index (1-based)
        try:
            col_idx = int(coerce_to_number(col_index_num))
            if col_idx < 1:
                return "#VALUE!"
            if col_idx > len(table_array[0]):
                return "#REF!"
        except Exception:
            return "#VALUE!"

        # Convert range_lookup to boolean
        try:
            exact_match = not coerce_to_bool(range_lookup)
        except Exception:
            return "#VALUE!"

        # Search for the value in the first column
        match_row = None
        prev_row = None

        try:
            for i, row in enumerate(table_array):
                if not row:  # Skip empty rows
                    continue

                # Get the value in the first column
                current_val = row[0]
                if current_val is None:  # Skip empty cells
                    continue

                # If lookup value is numeric and current value is timedelta, convert timedelta to days
                if isinstance(lookup_value, (int, float)) and isinstance(
                    current_val, timedelta
                ):
                    current_val = current_val.total_seconds() / (24 * 3600)
                    current_text = current_val
                else:
                    try:
                        # Handle numeric comparisons properly
                        if isinstance(lookup_value, (int, float)) and isinstance(
                            current_val, (int, float)
                        ):
                            current_text = current_val
                        else:
                            current_text = coerce_to_text(current_val)
                    except Exception:
                        continue  # Skip values that can't be compared

                if exact_match:
                    # For exact match
                    if current_text == lookup_text:
                        match_row = i
                        break
                else:
                    # For approximate match (assumes sorted ascending)
                    try:
                        if current_text == lookup_text:
                            match_row = i
                            break
                        elif current_text > lookup_text:
                            match_row = (
                                prev_row  # Use previous row if current is too high
                            )
                            break
                        prev_row = i
                    except TypeError:
                        # If comparison fails, try string comparison
                        str_current = str(current_text)
                        str_lookup = str(lookup_text)
                        if str_current == str_lookup:
                            match_row = i
                            break
                        elif str_current > str_lookup:
                            match_row = prev_row
                            break
                        prev_row = i
        except Exception:
            return "#VALUE!"

        # If no match found
        if match_row is None:
            if exact_match or prev_row is None:
                return "#N/A"
            match_row = prev_row  # Use last row for approximate match

        # Return the value from the specified column
        try:
            result = table_array[match_row][col_idx - 1]
            return result if result is not None else "#N/A"
        except IndexError:
            return "#REF!"

    @staticmethod
    def FILTER(
        array: list | ExcelValue, include: list | ExcelValue, if_empty: ExcelValue = 0
    ) -> ExcelValue:
        """Filter an array based on conditions.

        Args:
            array: A 1D or 2D array to filter
            include: A 1D array of boolean conditions
            if_empty: Value to return if no values match the condition
        """
        # Convert include to 1D array if it isn't already
        include = to_1d_array(include)

        # Handle 2D array input
        is_2d = isinstance(array, list) and array and isinstance(array[0], list)
        if is_2d:
            # For 2D arrays, we filter rows based on the include condition
            if len(array) != len(include):
                raise ExcelFunctionError(
                    f"Array rows ({len(array)}) and include condition length ({len(include)}) do not match"
                )

            # Filter rows based on condition
            filtered_values = []
            for row, condition in zip(array, include):
                try:
                    # Skip any error values in the condition
                    if isinstance(condition, str) and condition.startswith("#"):
                        continue
                    # Only include row if condition is True
                    if condition:
                        filtered_values.append(row)
                except Exception:
                    continue

            # If no values match, return if_empty
            if not filtered_values:
                return if_empty

            # If we're evaluating a single cell within an array formula
            array_context = getattr(array, "array_context", None)
            if array_context is not None:
                row = array_context.get("row", 0)
                col = array_context.get("col", 0)
                # Return the value at the requested position or if_empty if out of range
                if 0 <= row < len(filtered_values) and 0 <= col < len(
                    filtered_values[0]
                ):
                    return filtered_values[row][col]
                return if_empty

            # Otherwise return the full filtered 2D array
            return filtered_values
        else:
            # For 1D arrays, convert to 1D if needed
            array = to_1d_array(array)

            # Validate array and include lengths match
            if len(array) != len(include):
                raise ExcelFunctionError(
                    f"Array and include condition lengths do not match: {len(array)} vs {len(include)}"
                )

            # Filter values based on condition
            filtered_values = []
            for value, condition in zip(array, include):
                try:
                    # Skip any error values in the condition
                    if isinstance(condition, str) and condition.startswith("#"):
                        continue
                    # Only include value if condition is True
                    if condition:
                        filtered_values.append(value)
                except Exception:
                    continue

            # If no values match, return if_empty
            if not filtered_values:
                return if_empty

            # If we're evaluating a single cell within an array formula
            array_context = getattr(array, "array_context", None)
            if array_context is not None:
                row = array_context.get("row", 0)
                # Return the value at the requested position or if_empty if out of range
                if 0 <= row < len(filtered_values):
                    return filtered_values[row]
                return if_empty

            # Otherwise return the full filtered array
            return filtered_values

    @excel_fn(name="MODE.SNGL")
    @staticmethod
    def MODE_SNGL(values: list[float]) -> float:
        """Compute Excel's MODE.SNGL semantics for numeric inputs.

        - Returns the most frequent value.
        - In case of tie, returns the smallest value among those with max frequency (Excel's MODE.SNGL behavior).
        - Raises ExcelFunctionError if the input is empty.
        """
        if not values:
            raise ExcelFunctionError("No valid numeric values for MODE.SNGL")

        counts: dict[float, int] = {}
        for v in values:
            counts[v] = counts.get(v, 0) + 1

        max_count = max(counts.values())
        candidates = [v for v, c in counts.items() if c == max_count]
        return float(min(candidates))

    @staticmethod
    def AGGREGATE(
        function_num: ExcelValue,
        options: ExcelValue,
        ref1: ExcelValue,
        ref2: ExcelValue = None,
    ) -> ExcelValue:
        """
        Applies an aggregate function to a list or database with options to ignore hidden rows and error values.

        Args:
            function_num: Number 1-19 specifying which function to use
            options: Number 0-7 determining which values to ignore
            ref1: First numeric argument or array
            ref2: Optional second argument required for certain functions (k or quart)
        """
        # Convert arguments to appropriate types
        func_num = int(coerce_to_number(function_num))
        opt = int(coerce_to_number(options))

        # Define mapping of function numbers to their implementations
        AGGREGATE_FUNCTIONS = {
            1: ExcelFunctions.AVERAGE,
            2: lambda x: len([v for v in flatten_args(x) if v is not None]),  # COUNT
            3: lambda x: len([v for v in flatten_args(x) if v is not None]),  # COUNTA
            4: ExcelFunctions.MAX,
            5: ExcelFunctions.MIN,
            6: lambda x: np.prod(aggregate_numbers(flatten_args(x))),  # PRODUCT
            7: lambda x: np.std(aggregate_numbers(flatten_args(x)), ddof=1),  # STDEV.S
            8: lambda x: np.std(aggregate_numbers(flatten_args(x)), ddof=0),  # STDEV.P
            9: ExcelFunctions.SUM,
            10: lambda x: np.var(aggregate_numbers(flatten_args(x)), ddof=1),  # VAR.S
            11: lambda x: np.var(aggregate_numbers(flatten_args(x)), ddof=0),  # VAR.P
            12: lambda x: np.median(aggregate_numbers(flatten_args(x))),  # MEDIAN
            13: lambda x: ExcelFunctions.MODE_SNGL(
                aggregate_numbers(flatten_args(x))
            ),  # MODE.SNGL
            14: lambda x, k: sorted(aggregate_numbers(flatten_args(x)), reverse=True)[
                int(k) - 1
            ],  # LARGE
            15: lambda x, k: sorted(aggregate_numbers(flatten_args(x)))[
                int(k) - 1
            ],  # SMALL
            16: lambda x, k: np.percentile(
                aggregate_numbers(flatten_args(x)), float(k) * 100
            ),  # PERCENTILE.INC
            17: lambda x, k: np.percentile(
                aggregate_numbers(flatten_args(x)), float(k) * 25
            ),  # QUARTILE.INC
            18: lambda x, k: np.percentile(
                aggregate_numbers(flatten_args(x)), float(k) * 100, method="higher"
            ),  # PERCENTILE.EXC
            19: lambda x, k: np.percentile(
                aggregate_numbers(flatten_args(x)), float(k) * 25, method="higher"
            ),  # QUARTILE.EXC
        }

        # Validate function number
        if func_num not in AGGREGATE_FUNCTIONS:
            raise ExcelFunctionError(f"Invalid function number: {func_num}")

        # Validate options
        if opt not in range(8):
            raise ExcelFunctionError(f"Invalid options value: {opt}")

        # Get the selected function
        func = AGGREGATE_FUNCTIONS[func_num]

        # Process the data according to options
        data = flatten_args(ref1)
        processed_data = []

        for value in data:
            # Skip based on options
            if opt in [2, 3, 6, 7] and isinstance(value, str) and value.startswith("#"):
                continue
            # Note: options 1, 3, 5, 7 for hidden rows would need worksheet context
            processed_data.append(value)

        # Handle functions that require k parameter
        if func_num in [14, 15, 16, 17, 18, 19]:
            if ref2 is None:
                raise ExcelFunctionError("Second argument required for this function")
            try:
                return func(processed_data, coerce_to_number(ref2))
            except (IndexError, ValueError) as e:
                raise ExcelFunctionError(str(e))

        # Handle single-argument functions
        try:
            return func(processed_data)
        except (ValueError, TypeError) as e:
            raise ExcelFunctionError(str(e))

    @staticmethod
    def SUMPRODUCT(*arrays: ExcelValue) -> ExcelValue:
        """Returns the sum of the products of corresponding ranges or arrays.

        Args:
            *arrays: Two or more arrays whose components you want to multiply and then add.
                    All arrays must have the same dimensions.

        Returns:
            The sum of the products of corresponding array elements.
            Returns #VALUE! if arrays have different dimensions.
        """
        if not arrays:
            return 0

        # Convert all arrays to lists if they aren't already
        processed_arrays = []
        array_shape = None

        for array in arrays:
            # If it's not a list, treat it as a scalar and create a single-element list
            if not isinstance(array, list):
                array = [[array]] if array_shape and len(array_shape) > 1 else [array]

            # For 1D arrays, ensure they're all 1D
            if not isinstance(array[0], list):
                if array_shape and len(array_shape) > 1:
                    return "#VALUE!"  # Mixing 1D and 2D arrays
                array_shape = (len(array),)
            else:
                # For 2D arrays, check dimensions
                if not array_shape:
                    array_shape = (len(array), len(array[0]))
                elif array_shape != (len(array), len(array[0])):
                    return "#VALUE!"  # Arrays have different dimensions

            processed_arrays.append(array)

        # If we only got one array, just sum its elements
        if len(processed_arrays) == 1:
            flattened = flatten_args(processed_arrays[0])
            return sum(coerce_to_number(x) if x is not None else 0 for x in flattened)

        # Multiply corresponding elements and sum the results
        result = 0
        if len(array_shape) == 1:
            # 1D arrays
            for elements in zip(*processed_arrays):
                product = 1
                for element in elements:
                    # Treat non-numeric values as 0
                    try:
                        value = coerce_to_number(element) if element is not None else 0
                        product *= value
                    except (ValueError, TypeError):
                        product *= 0
                result += product
        else:
            # 2D arrays
            for i in range(array_shape[0]):
                for j in range(array_shape[1]):
                    product = 1
                    for array in processed_arrays:
                        try:
                            value = (
                                coerce_to_number(array[i][j])
                                if array[i][j] is not None
                                else 0
                            )
                            product *= value
                        except (ValueError, TypeError):
                            product *= 0
                    result += product

        return result

    @staticmethod
    def STDEV_P(*numbers: ExcelValue) -> ExcelValue:
        """Calculates standard deviation based on the entire population (ignores logical values and text).

        The standard deviation is a measure of how widely values are dispersed from the average value (mean).
        Uses the "n" method (population standard deviation).

        Args:
            *numbers: Number arguments corresponding to a population.
                     Can be numbers, arrays, or references containing numbers.
                     Logical values and text representations of numbers directly in arguments are counted.
                     Empty cells, logical values, text, or error values in arrays/references are ignored.

        Returns:
            The population standard deviation of the values.
            Returns #DIV/0! if no valid numbers are found.
            Returns #VALUE! if any argument is an error value or text that can't be translated to a number.
        """
        # Flatten and process all arguments
        try:
            values = []
            for num in flatten_args(*numbers):
                if num is None:  # Skip empty cells
                    continue
                try:
                    values.append(coerce_to_number(num))
                except (ValueError, TypeError):
                    continue  # Skip non-numeric values

            if not values:
                return "#DIV/0!"  # No valid numbers found

            n = len(values)
            mean = sum(values) / n

            # Calculate sum of squared differences from mean
            squared_diff_sum = sum((x - mean) ** 2 for x in values)

            # Population standard deviation formula: sqrt(Σ(x - μ)²/n)
            return (squared_diff_sum / n) ** 0.5

        except (ValueError, TypeError):
            return "#VALUE!"

    @staticmethod
    def IFS(*args: ExcelValue) -> ExcelValue:
        """Returns a value corresponding to the first TRUE condition in a series of condition-value pairs.

        Args:
            *args: A series of condition-value pairs. Must be an even number of arguments.
                  Each pair consists of a logical test and the value to return if true.

        Returns:
            The value corresponding to the first TRUE condition.
            Returns #N/A if no conditions are TRUE.
            Returns #VALUE! if a condition evaluates to a non-boolean value.
            Returns error if too few arguments or odd number of arguments.
        """
        # Check for minimum arguments (at least one condition-value pair)
        if len(args) < 2:
            return "#VALUE!"  # Too few arguments

        # Check for even number of arguments (pairs of condition-value)
        if len(args) % 2 != 0:
            return "#VALUE!"  # Missing value for a condition

        # Process each condition-value pair
        for i in range(0, len(args), 2):
            condition = args[i]
            value = args[i + 1]

            # Handle array inputs by vectorizing the operation
            if isinstance(condition, list):
                # For 2D arrays
                if isinstance(condition[0], list):
                    result = [[None] * len(row) for row in condition]
                    for i, row in enumerate(condition):
                        for j, cond in enumerate(row):
                            try:
                                if coerce_to_bool(cond):
                                    result[i][j] = (
                                        value[i][j]
                                        if isinstance(value, list)
                                        else value
                                    )
                            except (ValueError, TypeError):
                                result[i][j] = "#VALUE!"
                    return result
                # For 1D arrays
                result = [None] * len(condition)
                for i, cond in enumerate(condition):
                    try:
                        if coerce_to_bool(cond):
                            result[i] = value[i] if isinstance(value, list) else value
                    except (ValueError, TypeError):
                        result[i] = "#VALUE!"
                return result

            # Handle scalar case
            try:
                if coerce_to_bool(condition):
                    return value
            except (ValueError, TypeError):
                return "#VALUE!"

        # If no conditions were true
        return "#N/A"

    @staticmethod
    def ATAN(number: ExcelValue) -> ExcelValue:
        """Returns the arctangent, or inverse tangent, of a number.

        The arctangent is the angle whose tangent is number. The returned angle is given
        in radians in the range -pi/2 to pi/2.

        Args:
            number: The tangent of the angle you want.

        Returns:
            The arctangent in radians.
            Returns #VALUE! if the argument is not a number.
        """
        try:
            num = coerce_to_number(number)
            return math.atan(num)
        except (ValueError, TypeError):
            return "#VALUE!"

    @staticmethod
    def DEGREES(angle: ExcelValue) -> ExcelValue:
        """Converts radians into degrees.

        Args:
            angle: The angle in radians that you want to convert.

        Returns:
            The angle in degrees.
            Returns #VALUE! if the argument is not a number.
        """
        try:
            rad = coerce_to_number(angle)
            return math.degrees(rad)
        except (ValueError, TypeError):
            return "#VALUE!"

    @staticmethod
    def COUNTBLANK(range_values: ExcelValue) -> ExcelValue:
        """Counts the number of empty cells in a range of cells.

        Cells with formulas that return "" (empty text) are also counted.
        Cells with zero values are not counted.

        Args:
            range_values: The range from which you want to count the blank cells.

        Returns:
            The number of empty cells in the range.
        """
        count = 0

        # Flatten the range to process all values
        flat_values = flatten_args(range_values)

        for value in flat_values:
            # Count cells that are None (truly empty)
            if value is None:
                count += 1
            # Count cells with empty text (formulas returning "")
            elif isinstance(value, str) and value == "":
                count += 1

        return count

    @staticmethod
    def COUNTA(*args: ExcelValue) -> ExcelValue:
        """Counts the number of cells that are not empty in a range.

        The COUNTA function counts cells containing any type of information, including
        error values and empty text (""). For example, if the range contains a formula
        that returns an empty string, the COUNTA function counts that value.
        The COUNTA function does not count empty cells.

        Args:
            *args: The values that you want to count. Can be individual values,
                   arrays, or references containing values.

        Returns:
            The number of non-empty cells in the range.
        """
        count = 0

        # Flatten all arguments to process all values
        flat_values = flatten_args(*args)

        for value in flat_values:
            # Count all non-None values (including empty strings, errors, etc.)
            if value is not None:
                count += 1

        return count

    @excel_fn(name="NA")
    def NA() -> str:
        """Return the #N/A error value."""
        return "#N/A"


# Register all unregistered static methods on ExcelFunctions by their method names
for _name, _member in ExcelFunctions.__dict__.items():
    if _name.startswith("_"):
        continue
    if isinstance(_member, staticmethod):
        _func = _member.__func__
        if (
            not getattr(_func, "_excel_fn_registered", False)
            and _name not in EXCEL_FUNCTIONS
        ):
            EXCEL_FUNCTIONS[_name] = _func

if __name__ == "__main__":
    print("Registered functions:", list(EXCEL_FUNCTIONS.keys()))
