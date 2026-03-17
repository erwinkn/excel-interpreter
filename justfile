set shell := ["powershell.exe", "-NoLogo", "-NoProfile", "-Command"]

# Show all available repo tasks.
default:
    @just --list

# Build the Zig core and install artifacts under zig-out.
build-core:
    zig build

# Run the Zig unit tests.
test-core:
    zig build test

# Rebuild the Zig core as a quick smoke check.
smoke-core:
    zig build

# Build the Python package.
build-python:
    Set-Location packages/python; uv build

# Run the Python test suite.
test-python:
    uv run --project packages/python --with pytest pytest packages/python/tests

# Call the native Zig library from Python.
smoke-python-native:
    uv run --project packages/python python -c "import excel_interpreter; print(excel_interpreter.get_native_core_status()); print(excel_interpreter.native_add(2, 3)); print(excel_interpreter.native_greeting())"

# Build the TypeScript package with Bun.
build-ts:
    Set-Location packages/ts; bun x tsdown

# Run the TypeScript tests with Bun.
test-ts:
    Set-Location packages/ts; bun test

# Call the native Zig library from TypeScript via Bun FFI.
smoke-ts-native:
    bun packages/ts/scripts/native-smoke.ts

# Run both Python and TypeScript native wrapper smoke tests.
smoke-native:
    just smoke-python-native
    just smoke-ts-native

# Run the core and Python test suites.
test:
    just test-core
    just test-python

# Point to the manual Windows Excel oracle fuzzing flow.
fuzz-excel:
    Write-Host "See packages/core/fuzz/windows/README.md for the manual Windows Excel oracle harness."
