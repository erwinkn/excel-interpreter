const std = @import("std");
const excel_interpreter = @import("excel_interpreter");

pub const EiStatus = enum(c_int) {
    ok = 0,
    invalid_argument = 1,
    unsupported_formula = 2,
};

pub const EiValueTag = enum(c_int) {
    blank = 0,
    number = 1,
    boolean = 2,
};

pub const EiValue = extern struct {
    tag: c_int,
    number: f64,
    boolean: u8,
};

pub export fn ei_version_major() u32 {
    return excel_interpreter.version.major;
}

pub export fn ei_version_minor() u32 {
    return excel_interpreter.version.minor;
}

pub export fn ei_version_patch() u32 {
    return excel_interpreter.version.patch;
}

pub export fn ei_version_string() [*:0]const u8 {
    return "0.1.0";
}

pub export fn ei_add_f64(lhs: f64, rhs: f64) f64 {
    return lhs + rhs;
}

pub export fn ei_demo_greeting() [*:0]const u8 {
    return "hello from zig";
}

pub export fn ei_eval_formula_utf8(
    formula_ptr: [*]const u8,
    formula_len: usize,
    out_value: *EiValue,
) EiStatus {
    if (formula_len == 0) {
        return .invalid_argument;
    }

    const formula = formula_ptr[0..formula_len];
    const parsed = excel_interpreter.syntax.parseFormula(formula) catch return .unsupported_formula;
    const evaluator = excel_interpreter.runtime.Evaluator{};
    const value = evaluator.evaluate(parsed.expr) catch return .unsupported_formula;

    out_value.* = switch (value) {
        .blank => EiValue{
            .tag = @intFromEnum(EiValueTag.blank),
            .number = 0,
            .boolean = 0,
        },
        .number => |number| EiValue{
            .tag = @intFromEnum(EiValueTag.number),
            .number = number,
            .boolean = 0,
        },
        .boolean => |boolean| EiValue{
            .tag = @intFromEnum(EiValueTag.boolean),
            .number = 0,
            .boolean = @intFromBool(boolean),
        },
    };

    return .ok;
}

test "c api evaluates numeric formulas" {
    var value = EiValue{
        .tag = @intFromEnum(EiValueTag.blank),
        .number = 0,
        .boolean = 0,
    };

    const status = ei_eval_formula_utf8("=3.5".ptr, "=3.5".len, &value);

    try std.testing.expectEqual(EiStatus.ok, status);
    try std.testing.expectEqual(@intFromEnum(EiValueTag.number), value.tag);
    try std.testing.expectEqual(@as(f64, 3.5), value.number);
}

test "c api exposes demo helpers" {
    try std.testing.expectEqual(@as(f64, 5.0), ei_add_f64(2, 3));
    try std.testing.expectEqualStrings("hello from zig", std.mem.span(ei_demo_greeting()));
}
