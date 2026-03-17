const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});

    const core_module = b.addModule("excel_interpreter", .{
        .root_source_file = b.path("packages/core/src/root.zig"),
        .target = target,
        .optimize = optimize,
    });

    const zig_lib = b.addLibrary(.{
        .linkage = .static,
        .name = "excel_interpreter_core",
        .root_module = core_module,
    });
    b.installArtifact(zig_lib);

    const c_api_lib = b.addLibrary(.{
        .linkage = .dynamic,
        .name = "excel_interpreter",
        .root_module = b.createModule(.{
            .root_source_file = b.path("packages/core/src/ffi/c_api.zig"),
            .target = target,
            .optimize = optimize,
        }),
    });
    c_api_lib.root_module.addImport("excel_interpreter", core_module);
    c_api_lib.linkLibC();
    b.installArtifact(c_api_lib);

    const install_header = b.addInstallFileWithDir(
        b.path("packages/core/include/excel_interpreter.h"),
        .header,
        "excel_interpreter.h",
    );
    b.getInstallStep().dependOn(&install_header.step);

    const core_tests = b.addTest(.{
        .root_module = core_module,
    });

    const run_core_tests = b.addRunArtifact(core_tests);

    const test_step = b.step("test", "Run packages/core Zig tests");
    test_step.dependOn(&run_core_tests.step);

    const fmt_step = b.step("fmt", "Format Zig sources");
    const fmt = b.addFmt(.{
        .paths = &.{
            "build.zig",
            "packages/core/src",
        },
    });
    fmt_step.dependOn(&fmt.step);
}
