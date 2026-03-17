const std = @import("std");
const Expr = @import("../ir/expr.zig").Expr;
const tokenizer = @import("tokenizer.zig");

pub const ParsedFormula = struct {
    expr: Expr,
};

pub fn parseFormula(source: []const u8) !ParsedFormula {
    const trimmed = std.mem.trim(u8, source, " \t\r\n");
    if (trimmed.len == 0) {
        return error.EmptyFormula;
    }

    const body = if (trimmed[0] == '=') trimmed[1..] else trimmed;
    const token = try tokenizer.tokenizeSingleNumericLiteral(body);
    const value = try std.fmt.parseFloat(f64, token.lexeme);

    return ParsedFormula{
        .expr = Expr{ .number_literal = value },
    };
}
