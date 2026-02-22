import pytest
from excel_interpreter.errors import TokenizerError
from excel_interpreter.tokenizer import ExcelTokenizer, Token, TokenType


def tokenize(formula: str) -> list[Token]:
    """Helper function to tokenize a formula."""
    tokenizer = ExcelTokenizer(formula)
    return tokenizer.tokenize()


def assert_tokens(formula: str, expected: list[tuple[TokenType, str]]):
    """Helper function to assert tokens match expected types and values."""
    tokens = tokenize(formula)
    assert len(tokens) == len(expected), (
        f"Expected {len(expected)} tokens, got {len(tokens)}\n"
        f"Expected: {expected}\n"
        f"Got: {[(t.type, t.value) for t in tokens]}"
    )
    for token, (exp_type, exp_value) in zip(tokens, expected):
        assert token.type == exp_type, f"Expected {exp_type}, got {token.type}"
        assert token.value == exp_value, f"Expected {exp_value}, got {token.value}"


class TestExcelTokenizer:
    def test_simple_arithmetic(self):
        """Test basic arithmetic operators and numbers."""
        assert_tokens(
            "1 + 2",
            [
                (TokenType.NUMBER, "1"),
                (TokenType.OPERATOR, "+"),
                (TokenType.NUMBER, "2"),
            ],
        )

        assert_tokens(
            "2 * 3",
            [
                (TokenType.NUMBER, "2"),
                (TokenType.OPERATOR, "*"),
                (TokenType.NUMBER, "3"),
            ],
        )

    def test_decimal_numbers(self):
        """Test decimal number handling."""
        assert_tokens(
            "1.5 + 2.75",
            [
                (TokenType.NUMBER, "1.5"),
                (TokenType.OPERATOR, "+"),
                (TokenType.NUMBER, "2.75"),
            ],
        )

        # Test invalid decimal formats
        invalid_decimals = [
            ("1.2.3", "multiple decimal points"),
            (".", "lone decimal point"),
            ("1.", "trailing decimal point"),
        ]

        for invalid_num, error_msg in invalid_decimals:
            with pytest.raises(
                TokenizerError, match=f"Invalid number format.*{error_msg}"
            ):
                tokenize(invalid_num)

    def test_cell_references(self):
        """Test various forms of cell references."""
        # Simple cell reference
        assert_tokens(
            "A1",
            [(TokenType.IDENTIFIER, "A1")],
        )

        # Absolute references
        assert_tokens(
            "$A$1",
            [(TokenType.IDENTIFIER, "$A$1")],
        )

        # Mixed absolute/relative
        assert_tokens(
            "$A1 + A$1",
            [
                (TokenType.IDENTIFIER, "$A1"),
                (TokenType.OPERATOR, "+"),
                (TokenType.IDENTIFIER, "A$1"),
            ],
        )

    def test_cell_ranges(self):
        """Test cell range expressions."""
        assert_tokens(
            "A1:B2",
            [
                (TokenType.IDENTIFIER, "A1"),
                (TokenType.COLON, ":"),
                (TokenType.IDENTIFIER, "B2"),
            ],
        )

        # With sheet reference
        assert_tokens(
            "Sheet1!A1:B2",
            [
                (TokenType.IDENTIFIER, "Sheet1"),
                (TokenType.BANG, "!"),
                (TokenType.IDENTIFIER, "A1"),
                (TokenType.COLON, ":"),
                (TokenType.IDENTIFIER, "B2"),
            ],
        )

    def test_sheet_references(self):
        """Test sheet name handling."""
        # Simple sheet reference
        assert_tokens(
            "Sheet1!A1",
            [
                (TokenType.IDENTIFIER, "Sheet1"),
                (TokenType.BANG, "!"),
                (TokenType.IDENTIFIER, "A1"),
            ],
        )

        # Quoted sheet name
        assert_tokens(
            "'Sheet 1'!A1",
            [
                (TokenType.QUOTED_STRING, "Sheet 1"),
                (TokenType.BANG, "!"),
                (TokenType.IDENTIFIER, "A1"),
            ],
        )

    def test_functions(self):
        """Test function calls and arguments."""
        assert_tokens(
            "SUM(A1:B2)",
            [
                (TokenType.IDENTIFIER, "SUM"),
                (TokenType.LPAREN, "("),
                (TokenType.IDENTIFIER, "A1"),
                (TokenType.COLON, ":"),
                (TokenType.IDENTIFIER, "B2"),
                (TokenType.RPAREN, ")"),
            ],
        )

        # Multiple arguments
        assert_tokens(
            "SUM(A1, B2, C3)",
            [
                (TokenType.IDENTIFIER, "SUM"),
                (TokenType.LPAREN, "("),
                (TokenType.IDENTIFIER, "A1"),
                (TokenType.COMMA, ","),
                (TokenType.IDENTIFIER, "B2"),
                (TokenType.COMMA, ","),
                (TokenType.IDENTIFIER, "C3"),
                (TokenType.RPAREN, ")"),
            ],
        )

    def test_boolean_literals(self):
        """Test boolean literal handling."""
        assert_tokens(
            "TRUE",
            [(TokenType.BOOLEAN, "TRUE")],
        )

        assert_tokens(
            "FALSE",
            [(TokenType.BOOLEAN, "FALSE")],
        )

        # Case insensitive
        assert_tokens(
            "True",
            [(TokenType.BOOLEAN, "TRUE")],
        )

    def test_operators(self):
        """Test operators."""

        def _test(op):
            assert_tokens(
                f"A1 {op} B1",
                [
                    (TokenType.IDENTIFIER, "A1"),
                    (TokenType.OPERATOR, op),
                    (TokenType.IDENTIFIER, "B1"),
                ],
            )

        _test("+")
        _test("-")
        _test("*")
        _test("/")
        _test("^")
        _test("%")
        _test("&")
        _test("=")
        _test("<")
        _test(">")
        _test("<=")
        _test(">=")
        _test("<>")

    def test_complex_formulas(self):
        """Test more complex formula combinations."""
        # Nested functions
        assert_tokens(
            "SUM(A1:A3, MAX(B1:B3))",
            [
                (TokenType.IDENTIFIER, "SUM"),
                (TokenType.LPAREN, "("),
                (TokenType.IDENTIFIER, "A1"),
                (TokenType.COLON, ":"),
                (TokenType.IDENTIFIER, "A3"),
                (TokenType.COMMA, ","),
                (TokenType.IDENTIFIER, "MAX"),
                (TokenType.LPAREN, "("),
                (TokenType.IDENTIFIER, "B1"),
                (TokenType.COLON, ":"),
                (TokenType.IDENTIFIER, "B3"),
                (TokenType.RPAREN, ")"),
                (TokenType.RPAREN, ")"),
            ],
        )

        # Complex arithmetic with parentheses
        assert_tokens(
            "(A1 + B1) * (C1 + D1)",
            [
                (TokenType.LPAREN, "("),
                (TokenType.IDENTIFIER, "A1"),
                (TokenType.OPERATOR, "+"),
                (TokenType.IDENTIFIER, "B1"),
                (TokenType.RPAREN, ")"),
                (TokenType.OPERATOR, "*"),
                (TokenType.LPAREN, "("),
                (TokenType.IDENTIFIER, "C1"),
                (TokenType.OPERATOR, "+"),
                (TokenType.IDENTIFIER, "D1"),
                (TokenType.RPAREN, ")"),
            ],
        )

    def test_error_handling(self):
        with pytest.raises(TokenizerError, match="Unexpected character"):
            tokenize("A1 @ B1")

        with pytest.raises(TokenizerError, match="Unterminated quoted identifier"):
            tokenize("'Sheet1")

    def test_empty_formula(self):
        assert tokenize("") == []
        assert tokenize("   ") == []


class TestStringLiterals:
    def test_basic_strings(self):
        """Test basic string literal handling."""
        assert_tokens(
            '="hello"', [(TokenType.OPERATOR, "="), (TokenType.STRING, "hello")]
        )
        assert_tokens('="123"', [(TokenType.OPERATOR, "="), (TokenType.STRING, "123")])
        assert_tokens(
            '=""', [(TokenType.OPERATOR, "="), (TokenType.STRING, "")]
        )  # Empty string

    def test_escaped_quotes(self):
        """Test strings with escaped quotes."""
        assert_tokens(
            '="he""llo"', [(TokenType.OPERATOR, "="), (TokenType.STRING, 'he"llo')]
        )
        assert_tokens(
            '="multiple""quotes""here"',
            [(TokenType.OPERATOR, "="), (TokenType.STRING, 'multiple"quotes"here')],
        )
        assert_tokens(
            '=""""', [(TokenType.OPERATOR, "="), (TokenType.STRING, '"')]
        )  # Just a quote

    def test_special_characters(self):
        """Test strings with special characters."""
        assert_tokens(
            '="hello\nworld"',
            [(TokenType.OPERATOR, "="), (TokenType.STRING, "hello\nworld")],
        )
        assert_tokens(
            '="!@#$%^&*()"',
            [(TokenType.OPERATOR, "="), (TokenType.STRING, "!@#$%^&*()")],
        )
        assert_tokens(
            '="tab\there"', [(TokenType.OPERATOR, "="), (TokenType.STRING, "tab\there")]
        )

    def test_string_errors(self):
        """Test string literal error cases."""
        error_cases = [
            '"unterminated',
            '="unterminated',
            '="missing quote',
            '="unmatched"quote"',
        ]
        for formula in error_cases:
            with pytest.raises(TokenizerError):
                tokenize(formula)


class TestQuotedIdentifiers:
    def test_basic_identifiers(self):
        """Test basic quoted identifier handling."""
        cases = [
            (
                "='Sheet 1'",
                [(TokenType.OPERATOR, "="), (TokenType.QUOTED_STRING, "Sheet 1")],
            ),
            (
                "='My-Sheet'",
                [(TokenType.OPERATOR, "="), (TokenType.QUOTED_STRING, "My-Sheet")],
            ),
            (
                "=''",
                [(TokenType.OPERATOR, "="), (TokenType.QUOTED_STRING, "")],
            ),  # Empty identifier
        ]
        for formula, expected in cases:
            assert_tokens(formula, expected)

    def test_escaped_quotes(self):
        """Test identifiers with escaped quotes."""
        cases = [
            (
                "='Bob''s Sheet'",
                [(TokenType.OPERATOR, "="), (TokenType.QUOTED_STRING, "Bob's Sheet")],
            ),
            (
                "='Multiple''Quotes''Here'",
                [
                    (TokenType.OPERATOR, "="),
                    (TokenType.QUOTED_STRING, "Multiple'Quotes'Here"),
                ],
            ),
            (
                "=''''",
                [(TokenType.OPERATOR, "="), (TokenType.QUOTED_STRING, "'")],
            ),  # Just a quote
        ]
        for formula, expected in cases:
            assert_tokens(formula, expected)

    def test_sheet_references(self):
        """Test quoted sheet references in formulas."""
        cases = [
            (
                "='Sheet 1'!A1",
                [
                    (TokenType.OPERATOR, "="),
                    (TokenType.QUOTED_STRING, "Sheet 1"),
                    (TokenType.BANG, "!"),
                    (TokenType.IDENTIFIER, "A1"),
                ],
            ),
            (
                "='My Sheet'!B2+'Other Sheet'!C3",
                [
                    (TokenType.OPERATOR, "="),
                    (TokenType.QUOTED_STRING, "My Sheet"),
                    (TokenType.BANG, "!"),
                    (TokenType.IDENTIFIER, "B2"),
                    (TokenType.OPERATOR, "+"),
                    (TokenType.QUOTED_STRING, "Other Sheet"),
                    (TokenType.BANG, "!"),
                    (TokenType.IDENTIFIER, "C3"),
                ],
            ),
        ]
        for formula, expected in cases:
            assert_tokens(formula, expected)

    def test_identifier_errors(self):
        """Test quoted identifier error cases."""
        error_cases = [
            "'unterminated",
            "='unterminated",
            "='missing quote",
            "='unmatched'quote'",
        ]
        for formula in error_cases:
            with pytest.raises(TokenizerError):
                tokenize(formula)


class TestMixedQuoting:
    def test_mixed_quotes(self):
        """Test formulas with both types of quotes."""
        cases = [
            (
                "='Sheet 1'!A1&\"hello\"",
                [
                    (TokenType.OPERATOR, "="),
                    (TokenType.QUOTED_STRING, "Sheet 1"),
                    (TokenType.BANG, "!"),
                    (TokenType.IDENTIFIER, "A1"),
                    (TokenType.OPERATOR, "&"),
                    (TokenType.STRING, "hello"),
                ],
            ),
            (
                "=\"Value: \"&'My Sheet'!B2",
                [
                    (TokenType.OPERATOR, "="),
                    (TokenType.STRING, "Value: "),
                    (TokenType.OPERATOR, "&"),
                    (TokenType.QUOTED_STRING, "My Sheet"),
                    (TokenType.BANG, "!"),
                    (TokenType.IDENTIFIER, "B2"),
                ],
            ),
        ]
        for formula, expected in cases:
            assert_tokens(formula, expected)

    def test_nested_quotes(self):
        """Test complex quoting scenarios."""
        cases = [
            (
                "=\"'Quoted' text\"",
                [(TokenType.OPERATOR, "="), (TokenType.STRING, "'Quoted' text")],
            ),
            (
                "='\"Sheet\" Name'!A1",
                [
                    (TokenType.OPERATOR, "="),
                    (TokenType.QUOTED_STRING, '"Sheet" Name'),
                    (TokenType.BANG, "!"),
                    (TokenType.IDENTIFIER, "A1"),
                ],
            ),
        ]
        for formula, expected in cases:
            assert_tokens(formula, expected)


class TestScientificNotation:
    def test_basic_scientific_notation(self):
        """Test basic scientific notation formats."""
        assert_tokens(
            "4.07e-05",
            [(TokenType.NUMBER, "4.07e-05")],
        )
        assert_tokens(
            "1.23E+10",
            [(TokenType.NUMBER, "1.23E+10")],
        )
        assert_tokens(
            "1e5",
            [(TokenType.NUMBER, "1e5")],
        )

    def test_scientific_notation_in_expressions(self):
        """Test scientific notation in expressions."""
        assert_tokens(
            "1.5e3 + 2.4E-2",
            [
                (TokenType.NUMBER, "1.5e3"),
                (TokenType.OPERATOR, "+"),
                (TokenType.NUMBER, "2.4E-2"),
            ],
        )

    def test_invalid_scientific_notation(self):
        """Test invalid scientific notation formats."""
        with pytest.raises(TokenizerError):
            tokenize("1.2e")

        with pytest.raises(TokenizerError):
            tokenize("1.2e+")

        with pytest.raises(TokenizerError):
            tokenize("1.2e-")

        with pytest.raises(TokenizerError):
            tokenize("1.2e1.1")

        with pytest.raises(TokenizerError):
            tokenize("1.2ee5")


if __name__ == "__main__":
    pytest.main([__file__])
