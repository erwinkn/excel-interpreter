from typing import List, Optional, Union, Callable

from excel_interpreter.errors import ParseError
from excel_interpreter.types import parse_number
from .tokenizer import Token, TokenType, ExcelTokenizer
from .ast import (
    ASTNode,
    ArrayLiteral,
    BinaryOperation,
    CellRange,
    CellReference,
    Constant,
    ExcelFunction,
    NameReference,
    UnaryOperation,
)
from .utils import extract_cell_reference


# Helper function to parse a formula string into an AST.
def parse_formula(formula: str) -> ASTNode:
    """Helper function to parse a formula string into an AST."""
    tokens = ExcelTokenizer(formula).tokenize()
    return ExcelParser(tokens).parse()


class ExcelParser:
    def __init__(self, tokens: List[Token]):
        self.tokens = tokens
        self.current = 0

    def parse(self) -> ASTNode:
        """Parse tokens into an AST."""
        # Skip leading equals sign if present
        self.current = 0
        if (
            len(self.tokens) > 0
            and self.tokens[0].type == TokenType.OPERATOR
            and self.tokens[0].value == "="
        ):
            self.current = 1

        return self.parse_expression()

    # We call .peek(), .read(), .read_optional() very often, so we duplicate
    # code to avoid unnecessary function calls.
    def peek(self) -> Optional[Token]:
        """Look at the current token without consuming it."""
        if self.current >= len(self.tokens):
            return None
        return self.tokens[self.current]

    def read(self) -> Token:
        """Consume and return the current token."""
        if self.current >= len(self.tokens):
            raise ParseError("Unexpected end of formula")
        tok = self.tokens[self.current]
        self.current += 1
        return tok

    def read_optional(self) -> Token | None:
        """Consume and return the current token, if there's one."""
        if self.current >= len(self.tokens):
            return None
        tok = self.tokens[self.current]
        self.current += 1
        return tok

    def read_if_match(self, *types: TokenType) -> Optional[Token]:
        """Consume and return current token if it matches any of the given types."""
        token = self.peek()
        if token is not None and token.type in types:
            self.current += 1
            return token
        return None

    def _parse_binary_operation(
        self, parse_operand: Callable[[], ASTNode], valid_operators: set[str]
    ) -> ASTNode:
        """Parse a binary operation with the given operand parser and operators."""
        left = parse_operand()

        while True:
            next_tok = self.peek()
            if (
                not next_tok
                or next_tok.type != TokenType.OPERATOR
                or next_tok.value not in valid_operators
            ):
                break

            self.read()  # consume operator
            right = parse_operand()
            left = BinaryOperation(left=left, operator=next_tok.value, right=right)

        return left

    def parse_expression(self) -> ASTNode:
        """Parse an expression (lowest precedence: comparisons)."""
        return self._parse_binary_operation(
            self.parse_concat, {"=", "<>", "<", ">", "<=", ">="}
        )

    def parse_concat(self) -> ASTNode:
        """Parse string concatenation (&)."""
        return self._parse_binary_operation(self.parse_additive, {"&"})

    def parse_additive(self) -> ASTNode:
        """Parse addition/subtraction (+, -)."""
        return self._parse_binary_operation(self.parse_term, {"+", "-"})

    def parse_term(self) -> ASTNode:
        """Parse multiplication/division (*, /)."""
        return self._parse_binary_operation(self.parse_power, {"*", "/"})

    def parse_power(self) -> ASTNode:
        """Parse exponentiation (^)."""
        return self._parse_binary_operation(self.parse_factor, {"^"})

    def parse_factor(self) -> ASTNode:
        """Parse a factor (highest precedence: literals, references, functions)."""
        token = self.peek()
        if token is None:
            raise ParseError("Unexpected end of formula")

        if token.type == TokenType.NUMBER:
            self.read()
            return Constant(parse_number(token.value))

        if token.type == TokenType.STRING:
            self.read()
            return Constant(token.value)

        elif token.type == TokenType.BOOLEAN:
            self.read()
            return Constant(token.value == "TRUE")

        elif token.type == TokenType.ERROR:
            self.read()
            return Constant(token.value)

        elif token.type == TokenType.LBRACE:
            return self.parse_array()

        elif token.type in [TokenType.QUOTED_STRING, TokenType.IDENTIFIER]:
            # Both quoted strings and identifiers could be sheet names or function names
            return self.parse_identifier()

        elif token.type == TokenType.OPERATOR and token.value in ["+", "-"]:
            operator = self.read().value
            operand = self.parse_factor()
            return UnaryOperation(operator=operator, operand=operand)

        elif token.type == TokenType.LPAREN:
            self.read()  # consume '('
            expr = self.parse_expression()
            if not self.read_if_match(TokenType.RPAREN):
                raise ParseError("Expected closing parenthesis ')'")
            return expr

        raise ParseError(
            f"Unexpected token: {token.type.name} at position {token.position}"
        )

    def parse_array(self) -> ASTNode:
        """Parse an array literal like {1,2,3} or {1;2;3} for vertical arrays."""
        assert self.read().type == TokenType.LBRACE  # consume '{'
        elements = []
        separator = None

        while (
            next_token := self.peek()
        ) is not None and next_token.type != TokenType.RBRACE:
            elements.append(self.parse_expression())

            next_tok = self.peek()
            if not next_tok:
                raise ParseError("Unexpected end of array literal")

            if next_tok.type == TokenType.RBRACE:
                self.read()  # consume '}'
                break

            if next_tok.type in (TokenType.COMMA, TokenType.SEMICOLON):
                if separator is None:
                    separator = next_tok.type
                elif separator != next_tok.type:
                    raise ParseError(
                        "Cannot mix comma and semicolon separators in array literal"
                    )
                self.read()  # consume separator
                continue

            raise ParseError(
                f"Expected ',', ';' or '}}' in array literal, got {next_tok.type.name}"
            )

        return ArrayLiteral(
            elements=tuple(elements), vertical=separator == TokenType.SEMICOLON
        )

    def parse_identifier(self) -> ASTNode:
        """Parse an identifier (function call, cell reference, or name)."""
        token = self.read()

        # Look ahead
        next_token = self.peek()

        if next_token and next_token.type == TokenType.LPAREN:
            self.read()  # consume '('
            return self.parse_function_call(token.value)
        if next_token and next_token.type == TokenType.RPAREN:
            # Only raise LPAREN error if we're not inside a function call
            if self.current == 1:  # We're at the start (after reading first token)
                raise ParseError("Expected LPAREN")

        # Look ahead for '!' to check for sheet reference
        if next_token and next_token.type == TokenType.BANG:
            self.read()  # consume '!'
            sheet = token.value
            return self.parse_cell_reference(sheet)

        # Look ahead for ':' to check for range
        if next_token and next_token.type == TokenType.COLON:
            self.read()  # consume ':'
            start_ref = extract_cell_reference(token.value)
            if not start_ref:
                raise ParseError(f"Invalid cell reference: {token.value}")

            end_token = self.read()
            if end_token.type != TokenType.IDENTIFIER:
                raise ParseError(
                    f"Expected cell reference for end of range, got {end_token.type.name}"
                )

            end_ref = extract_cell_reference(end_token.value)
            if not end_ref:
                raise ParseError(f"Invalid cell reference: {end_token.value}")

            return CellRange(start=start_ref, end=end_ref)

        if cell_ref := extract_cell_reference(token.value):
            return cell_ref

        return NameReference(name=token.value)

    def parse_function_call(self, name: str) -> ExcelFunction:
        """Parse a function call with its arguments."""
        args = []

        # Handle empty argument list
        if self.read_if_match(TokenType.RPAREN):
            return ExcelFunction(name=name, arguments=())

        # Parse arguments
        while True:
            args.append(self.parse_expression())

            next_tok = self.peek()
            if not next_tok:
                raise ParseError("Unexpected end of formula in function call")

            if next_tok.type == TokenType.RPAREN:
                self.read()  # consume ')'
                break

            if next_tok.type == TokenType.COMMA:
                self.read()  # consume ','
                continue

            raise ParseError(
                f"Expected ',' or ')' in function call, got {next_tok.type.name}"
            )

        return ExcelFunction(name=name, arguments=tuple(args))

    def parse_cell_reference(
        self, sheet: Optional[str] = None
    ) -> Union[CellReference, CellRange]:
        """Parse a cell reference or range with optional sheet reference."""
        token = self.read_if_match(TokenType.IDENTIFIER)
        if not token:
            curr = self.peek()
            raise ParseError(
                f"Expected cell reference, got "
                f"{curr.type.name if curr else 'end of formula'}"
            )

        ref = extract_cell_reference(token.value, sheet)
        if not ref:
            raise ParseError(f"Invalid cell reference: {token.value}")

        # Check for range
        if (next_token := self.peek()) and next_token.type == TokenType.COLON:
            self.read()  # consume colon

            # Check if the next token is a sheet reference
            end_sheet = None
            next_token = self.peek()
            if next_token and next_token.type in [
                TokenType.QUOTED_STRING,
                TokenType.IDENTIFIER,
            ]:
                potential_sheet_token = self.read()
                next_token = self.peek()

                if next_token and next_token.type == TokenType.BANG:
                    # This is a sheet reference
                    self.read()  # consume '!'
                    end_sheet = potential_sheet_token.value

                    # If sheets are different, raise an error
                    if end_sheet != sheet:
                        raise ParseError(
                            f"Range must be on the same sheet: {sheet} vs {end_sheet}"
                        )
                else:
                    # Not a sheet reference, put the token back
                    self.current -= 1

            end_token = self.read_if_match(TokenType.IDENTIFIER)
            if not end_token:
                curr = self.peek()
                raise ParseError(
                    f"Expected cell reference after ':', got "
                    f"{curr.type.name if curr else 'end of formula'}"
                )
            end_ref = extract_cell_reference(end_token.value, end_sheet or sheet)
            if not end_ref:
                raise ParseError(f"Invalid cell reference: {end_token.value}")
            return CellRange(start=ref, end=end_ref)

        return ref

    def expect(self, *types: TokenType) -> Token:
        """Read and return the current token if it matches expected types, otherwise error."""
        token = self.read_if_match(*types)
        if token is None:
            curr = self.peek()
            type_names = " or ".join(t.name for t in types)
            raise ParseError(
                f"Expected {type_names}, got "
                f"{curr.type.name if curr else 'end of formula'}"
                f" at position {curr.position if curr else len(self.tokens)}"
            )
        return token

    def peek_prev(self) -> Optional[Token]:
        """Look at the previous token without changing position."""
        if self.current > 0:
            return self.tokens[self.current - 1]
        return None
