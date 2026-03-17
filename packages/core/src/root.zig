pub const ir = struct {
    pub const Expr = @import("ir/expr.zig").Expr;
};

pub const runtime = struct {
    pub const Value = @import("runtime/value.zig").Value;
    pub const Evaluator = @import("runtime/evaluator.zig").Evaluator;
};

pub const syntax = struct {
    pub const ParsedFormula = @import("syntax/parser.zig").ParsedFormula;
    pub const parseFormula = @import("syntax/parser.zig").parseFormula;
};

pub const version = struct {
    pub const major: u32 = 0;
    pub const minor: u32 = 1;
    pub const patch: u32 = 0;

    pub fn string() []const u8 {
        return "0.1.0";
    }
};

test "parse and evaluate a numeric literal formula" {
    const evaluator = runtime.Evaluator{};
    const parsed = try syntax.parseFormula("=42.5");
    const value = try evaluator.evaluate(parsed.expr);

    try std.testing.expectEqual(@as(f64, 42.5), value.number);
}

test "parse and evaluate a raw numeric literal" {
    const evaluator = runtime.Evaluator{};
    const parsed = try syntax.parseFormula("7");
    const value = try evaluator.evaluate(parsed.expr);

    try std.testing.expectEqual(@as(f64, 7), value.number);
}

const std = @import("std");
