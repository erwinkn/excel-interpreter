from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable

ExcelScalar = str | int | float | bool | None
CellResolver = Callable[[str, str | None], ExcelScalar]


@dataclass
class ExpressionBuilder:
    _parts: list[str] = field(default_factory=list)

    def ref(self, address: str, sheet: str | None = None) -> "ExpressionBuilder":
        self._parts.append(f"{sheet}!{address}" if sheet else address)
        return self

    def lit(self, value: ExcelScalar) -> "ExpressionBuilder":
        if value is None:
            self._parts.append('""')
        elif isinstance(value, bool):
            self._parts.append("TRUE" if value else "FALSE")
        elif isinstance(value, str):
            escaped = value.replace('"', '""')
            self._parts.append(f'"{escaped}"')
        else:
            self._parts.append(str(value))
        return self

    def fn(self, name: str, *args: str) -> "ExpressionBuilder":
        self._parts.append(f"{name.upper()}({','.join(args)})")
        return self

    def raw(self, fragment: str) -> "ExpressionBuilder":
        self._parts.append(fragment)
        return self

    def build(self) -> str:
        return "".join(self._parts)


def evaluate_formula(formula: str, _resolver: CellResolver) -> ExcelScalar:
    """Evaluate an Excel formula string.

    Placeholder implementation. Workbook-aware evaluation will be added incrementally.
    """
    if not formula.startswith("="):
        return formula
    raise NotImplementedError("evaluate_formula is not implemented yet.")
