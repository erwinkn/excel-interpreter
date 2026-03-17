const std = @import("std");

pub const TokenKind = enum {
    number,
};

pub const Token = struct {
    kind: TokenKind,
    lexeme: []const u8,
};

pub fn tokenizeSingleNumericLiteral(source: []const u8) !Token {
    const trimmed = std.mem.trim(u8, source, " \t\r\n");
    if (trimmed.len == 0) {
        return error.EmptyFormula;
    }

    _ = try std.fmt.parseFloat(f64, trimmed);
    return Token{
        .kind = .number,
        .lexeme = trimmed,
    };
}
