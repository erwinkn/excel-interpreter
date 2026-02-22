import pytest

from excel_interpreter.ast import (
    BinaryOperation,
    CellRange,
    CellReference,
    Constant,
    ExcelFunction,
    NameReference,
)
from excel_interpreter.errors import ParseError
from excel_interpreter.parser import parse_formula


class TestExcelParser:
    def test_simple_arithmetic(self):
        """Test parsing of basic arithmetic expressions."""
        # Test addition
        ast = parse_formula("1 + 2")
        assert isinstance(ast, BinaryOperation)
        assert ast.operator == "+"
        assert isinstance(ast.left, Constant) and ast.left.value == 1.0
        assert isinstance(ast.right, Constant) and ast.right.value == 2.0

        # Test multiplication with parentheses
        ast = parse_formula("(2 + 3) * 4")
        assert isinstance(ast, BinaryOperation)
        assert ast.operator == "*"
        assert isinstance(ast.right, Constant) and ast.right.value == 4.0
        assert isinstance(ast.left, BinaryOperation)
        assert ast.left.operator == "+"

    def test_operator_precedence(self):
        """Test that operator precedence is correctly handled."""
        # Multiplication before addition
        ast = parse_formula("1 + 2 * 3")
        assert isinstance(ast, BinaryOperation)
        assert ast.operator == "+"
        assert isinstance(ast.left, Constant) and ast.left.value == 1.0
        assert isinstance(ast.right, BinaryOperation)
        assert ast.right.operator == "*"

        # Complex expression
        ast = parse_formula("1 + 2 * 3 / 4")
        assert isinstance(ast, BinaryOperation)
        assert ast.operator == "+"
        assert isinstance(ast.right, BinaryOperation)
        assert ast.right.operator in ["*", "/"]

    def test_cell_references(self):
        """Test parsing of cell references."""
        # Simple cell reference
        ast = parse_formula("A1")
        assert isinstance(ast, CellReference)
        assert ast.column == "A"
        assert ast.row == 1
        assert not ast.absolute_col
        assert not ast.absolute_row

        # Absolute references
        ast = parse_formula("$A$1")
        assert isinstance(ast, CellReference)
        assert ast.column == "A"
        assert ast.row == 1
        assert ast.absolute_col
        assert ast.absolute_row

        # Mixed absolute/relative
        ast = parse_formula("$A1")
        assert isinstance(ast, CellReference)
        assert ast.absolute_col
        assert not ast.absolute_row

    def test_cell_ranges(self):
        """Test parsing of cell ranges."""
        ast = parse_formula("A1:B2")
        assert isinstance(ast, CellRange)
        assert isinstance(ast.start, CellReference)
        assert isinstance(ast.end, CellReference)
        assert ast.start.column == "A" and ast.start.row == 1
        assert ast.end.column == "B" and ast.end.row == 2

        # With sheet reference
        ast = parse_formula("Sheet1!A1:B2")
        assert isinstance(ast, CellRange)
        assert isinstance(ast, CellRange)
        assert isinstance(ast.start, CellReference)
        assert isinstance(ast.end, CellReference)
        assert ast.start.sheet == "Sheet1"
        assert ast.start.column == "A" and ast.start.row == 1
        assert ast.end.column == "B" and ast.end.row == 2

        # With double sheet reference
        ast = parse_formula("Sheet1!A1:'Sheet1'!B2")
        assert isinstance(ast, CellRange)
        assert isinstance(ast, CellRange)
        assert isinstance(ast.start, CellReference)
        assert isinstance(ast.end, CellReference)
        assert ast.end.sheet == "Sheet1"
        assert ast.start.column == "A" and ast.start.row == 1
        assert ast.end.column == "B" and ast.end.row == 2

    def test_invalid_ranges(self):
        with pytest.raises(ParseError):
            parse_formula("Sheet1!A1:Sheet2!B2")

    def test_sheet_references(self):
        """Test parsing of sheet references."""
        # Simple sheet reference
        ast = parse_formula("Sheet1!A1")
        assert isinstance(ast, CellReference)
        assert ast.sheet == "Sheet1"
        assert ast.column == "A"
        assert ast.row == 1

        # Quoted sheet name
        ast = parse_formula("'Sheet 1'!A1")
        assert isinstance(ast, CellReference)
        assert ast.sheet == "Sheet 1"

    def test_functions(self):
        """Test parsing of function calls."""
        # Simple function
        ast = parse_formula("SUM(1, 2)")
        assert isinstance(ast, ExcelFunction)
        assert ast.name == "SUM"
        assert len(ast.arguments) == 2
        assert all(isinstance(arg, Constant) for arg in ast.arguments)

        # Function with range
        ast = parse_formula("SUM(A1:B2)")
        assert isinstance(ast, ExcelFunction)
        assert len(ast.arguments) == 1
        assert isinstance(ast.arguments[0], CellRange)

        # Nested functions
        ast = parse_formula("SUM(1, MAX(2, 3))")
        assert isinstance(ast, ExcelFunction)
        assert ast.name == "SUM"
        assert len(ast.arguments) == 2
        assert isinstance(ast.arguments[1], ExcelFunction)
        assert ast.arguments[1].name == "MAX"

    def test_boolean_literals(self):
        """Test parsing of boolean literals."""
        ast = parse_formula("TRUE")
        assert isinstance(ast, Constant)
        assert ast.value is True

        ast = parse_formula("FALSE")
        assert isinstance(ast, Constant)
        assert ast.value is False

    def test_comparison_operators(self):
        """Test parsing of comparison operators."""
        operators = ["=", "<>", "<", ">", "<=", ">="]

        for op in operators:
            ast = parse_formula(f"A1 {op} B1")
            assert isinstance(ast, BinaryOperation)
            assert ast.operator == op
            assert isinstance(ast.left, CellReference)
            assert isinstance(ast.right, CellReference)

    def test_name_references(self):
        """Test parsing of named ranges."""
        ast = parse_formula("MyRange")
        assert isinstance(ast, NameReference)
        assert ast.name == "MyRange"

    def test_error_handling(self):
        """Test error handling for invalid formulas."""

        def expect_error(formula, error_msg):
            with pytest.raises(ParseError, match=error_msg):
                parse_formula(formula)

        expect_error("SUM(", "Unexpected end of formula")
        expect_error("SUM)", "Expected LPAREN")
        expect_error("A1:", "Unexpected end of formula")
        expect_error("Sheet1!", "Expected cell reference")
        expect_error("1 + ", "Unexpected end of formula")
        expect_error("(1 + 2", "Expected closing parenthesis")

    def test_complex_formulas(self):
        """Test parsing of complex formula combinations."""
        # Complex arithmetic with functions and ranges
        ast = parse_formula("SUM(A1:B2) * (MAX(C1, C2) + 10)")
        assert isinstance(ast, BinaryOperation)
        assert ast.operator == "*"
        assert isinstance(ast.left, ExcelFunction)
        assert isinstance(ast.right, BinaryOperation)

        # Nested functions with comparisons
        ast = parse_formula("IF(A1 > 10, SUM(B1:B2), 0)")
        assert isinstance(ast, ExcelFunction)
        assert ast.name == "IF"
        assert len(ast.arguments) == 3
        assert isinstance(ast.arguments[0], BinaryOperation)


if __name__ == "__main__":
    pytest.main([__file__])
