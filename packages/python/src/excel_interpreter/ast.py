from typing import NamedTuple


class ExcelFunction(NamedTuple):
    name: str
    arguments: "tuple[ASTNode, ...]"


class ArrayLiteral(NamedTuple):
    elements: "tuple[ASTNode, ...]"
    vertical: bool


class BinaryOperation(NamedTuple):
    left: "ASTNode"
    operator: str
    right: "ASTNode"


class UnaryOperation(NamedTuple):
    operator: str
    operand: "ASTNode"


class CellReference(NamedTuple):
    column: int | str
    row: int
    sheet: str | None = None
    absolute_col: bool = False
    absolute_row: bool = False

    def coords(self) -> str:
        # Avoid circular imports
        from excel_interpreter.utils import column_as_str

        return f"{column_as_str(self.column)}{self.row}"


class CellRange(NamedTuple):
    start: CellReference
    end: CellReference


class Constant(NamedTuple):
    value: float | str | bool


class NamedRange(NamedTuple):
    name: str
    reference: CellReference | CellRange


class NameReference(NamedTuple):
    name: str


# Type alias for all possible AST nodes
ASTNode = (
    ExcelFunction
    | BinaryOperation
    | UnaryOperation
    | CellReference
    | CellRange
    | Constant
    | NameReference
    | ArrayLiteral
)
