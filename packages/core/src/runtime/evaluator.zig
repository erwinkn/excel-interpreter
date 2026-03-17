const Expr = @import("../ir/expr.zig").Expr;
const Value = @import("value.zig").Value;

pub const Evaluator = struct {
    pub fn evaluate(_: Evaluator, expr: Expr) !Value {
        return switch (expr) {
            .number_literal => |value| Value{ .number = value },
        };
    }
};
