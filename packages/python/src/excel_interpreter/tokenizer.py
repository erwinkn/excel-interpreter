from enum import Enum, auto
from typing import List, NamedTuple

from excel_interpreter.errors import TokenizerError


class TokenType(Enum):
    IDENTIFIER = auto()
    NUMBER = auto()
    OPERATOR = auto()
    LPAREN = auto()
    RPAREN = auto()
    COMMA = auto()
    SEMICOLON = auto()
    COLON = auto()
    BANG = auto()
    QUOTED_STRING = auto()
    BOOLEAN = auto()
    STRING = auto()
    LBRACE = auto()
    RBRACE = auto()
    ERROR = auto()  # New token type for Excel errors


class Token(NamedTuple):
    type: TokenType
    value: str
    position: int


class ExcelTokenizer:
    TWO_CHAR_OPERATORS = {"<": {"=", ">"}, ">": {"="}}
    EXCEL_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}

    def __init__(self, formula: str):
        self.formula = formula.strip()
        self.pos = 0
        self.length = len(self.formula)

    def tokenize(self) -> List[Token]:
        """Tokenize the formula and return list of tokens."""
        tokens = []
        while self.pos < self.length:
            char = self.formula[self.pos]

            if char.isspace():
                self.pos += 1
                continue
            elif char == '"':
                tokens.append(self._tokenize_string())
            elif char == "'":
                tokens.append(self._tokenize_quoted_identifier())
            elif char == "#":
                tokens.append(self._tokenize_error())
            elif char.isdigit() or char == ".":
                tokens.append(self._tokenize_number())
            elif char.isalpha() or char == "_" or char == "$":
                tokens.append(self._tokenize_identifier())
            elif char in "+-*/=<>&%^":
                tokens.append(self._tokenize_operator())
            elif char == "(":
                tokens.append(Token(TokenType.LPAREN, char, self.pos))
                self.pos += 1
            elif char == ")":
                tokens.append(Token(TokenType.RPAREN, char, self.pos))
                self.pos += 1
            elif char == ",":
                tokens.append(Token(TokenType.COMMA, char, self.pos))
                self.pos += 1
            elif char == ";":
                tokens.append(Token(TokenType.SEMICOLON, char, self.pos))
                self.pos += 1
            elif char == ":":
                tokens.append(Token(TokenType.COLON, char, self.pos))
                self.pos += 1
            elif char == "!":
                tokens.append(Token(TokenType.BANG, char, self.pos))
                self.pos += 1
            elif char == "{":
                tokens.append(Token(TokenType.LBRACE, char, self.pos))
                self.pos += 1
            elif char == "}":
                tokens.append(Token(TokenType.RBRACE, char, self.pos))
                self.pos += 1
            else:
                raise TokenizerError(
                    f"Unexpected character: {char} at position {self.pos}"
                )

        return tokens

    def _tokenize_error(self) -> Token:
        """Tokenize an Excel error value."""
        start = self.pos
        # Look ahead to find the complete error token
        for error in self.EXCEL_ERRORS:
            if self.formula[start:].startswith(error):
                self.pos += len(error)
                return Token(TokenType.ERROR, error, start)
        
        # If we get here, it's not a valid error token
        raise TokenizerError(f"Invalid Excel error value at position {start}")

    def _tokenize_identifier(self) -> Token:
        """Tokenize an identifier (function name, cell reference, etc)."""
        start = self.pos
        while self.pos < self.length and (
            self.formula[self.pos].isalnum() or self.formula[self.pos] in "_$"
        ):
            self.pos += 1

        value = self.formula[start : self.pos]
        if value.upper() in ["TRUE", "FALSE"]:
            return Token(TokenType.BOOLEAN, value.upper(), start)
        else:
            return Token(TokenType.IDENTIFIER, value, start)

    def _tokenize_number(self) -> Token:
        """Tokenize a number (integer or decimal)."""
        start = self.pos
        seen_decimal = False
        has_digits = False
        seen_exponent = False

        while self.pos < self.length:
            char = self.formula[self.pos]

            if char.isdigit():
                has_digits = True
                self.pos += 1
            elif char == "." and not seen_decimal:
                # If we see a decimal point, mark it but continue
                seen_decimal = True
                self.pos += 1
            elif char == "." and seen_decimal:
                # Second decimal point - invalid number
                raise TokenizerError(
                    f"Invalid number format at position {start}: multiple decimal points"
                )
            elif (char == "e" or char == "E") and not seen_exponent and has_digits:
                # Handle scientific notation
                seen_exponent = True
                self.pos += 1
                # Check for optional sign after e/E
                if self.pos < self.length and (self.formula[self.pos] in "+-"):
                    self.pos += 1
                # Must have at least one digit after e/E
                if self.pos >= self.length or not self.formula[self.pos].isdigit():
                    raise TokenizerError(
                        f"Invalid scientific notation at position {start}: missing exponent"
                    )
            else:
                break

        value = self.formula[start : self.pos]

        # Validate the number format
        if value == ".":
            raise TokenizerError(
                f"Invalid number format at position {start}: lone decimal point"
            )
        elif not has_digits:
            raise TokenizerError(
                f"Invalid number format at position {start}: no digits"
            )
        elif value.endswith("."):
            raise TokenizerError(
                f"Invalid number format at position {start}: trailing decimal point"
            )
        elif (
            value.endswith("e")
            or value.endswith("E")
            or value.endswith("+")
            or value.endswith("-")
        ):
            raise TokenizerError(
                f"Invalid scientific notation at position {start}: incomplete exponent"
            )

        return Token(TokenType.NUMBER, value, start)

    def _tokenize_operator(self) -> Token:
        """Tokenize an operator (+, -, *, /, =, <, >, <=, >=, <>, &, %, ^)."""
        start = self.pos
        current_char = self.formula[self.pos]
        next_char = self.formula[self.pos + 1] if self.pos + 1 < self.length else None

        # Simpler two-character operator check
        if (
            next_char
            and current_char in self.TWO_CHAR_OPERATORS
            and next_char in self.TWO_CHAR_OPERATORS[current_char]
        ):
            self.pos += 2
            value = self.formula[start : self.pos]
        else:
            self.pos += 1
            value = current_char

        return Token(TokenType.OPERATOR, value, start)

    def _tokenize_string(self) -> Token:
        """Tokenize an Excel string literal.
        Rules:
        1. Strings start and end with double quotes
        2. Double quotes inside strings are escaped by doubling them
        3. Strings can contain any character including newlines
        """
        start = self.pos
        self.pos += 1  # Skip opening quote
        value = []
        while self.pos < self.length:
            char = self.formula[self.pos]
            if char == '"':
                self.pos += 1
                if self.pos < self.length and self.formula[self.pos] == '"':
                    # Double quote escaped by doubling
                    value.append('"')
                    self.pos += 1
                else:
                    # End of string
                    break
            else:
                value.append(char)
                self.pos += 1
        else:
            raise TokenizerError(f"Unterminated string literal '{''.join(value)}'")

        return Token(TokenType.STRING, "".join(value), start)

    def _tokenize_quoted_identifier(self) -> Token:
        """Tokenize a quoted identifier (e.g., 'Sheet Name'!A1).
        Rules:
        1. Single quotes are used to escape special characters in identifiers
        2. Two single quotes represent a literal single quote
        3. The quoted section ends at the next single quote
        """
        start = self.pos
        self.pos += 1  # Skip opening quote
        value = []
        while self.pos < self.length:
            char = self.formula[self.pos]
            if char == "'":
                self.pos += 1
                if self.pos < self.length and self.formula[self.pos] == "'":
                    # Double single quote escaped by doubling
                    value.append("'")
                    self.pos += 1
                else:
                    # End of quoted identifier
                    break
            else:
                value.append(char)
                self.pos += 1
        else:
            raise TokenizerError("Unterminated quoted identifier")

        return Token(TokenType.QUOTED_STRING, "".join(value), start)
