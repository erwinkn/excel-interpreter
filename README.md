# excel-interpreter

Monorepo for Excel formula tooling with a native Zig core and language-specific wrappers.

Packages:

- `packages/core`: Zig core for parsing, expression IR, evaluation, and the shared C ABI.
- `packages/python`: Python library and future native wrapper around the Zig core.
- `packages/ts`: TypeScript package and future Node binding around the Zig core.
- `tests/corpus`: shared deterministic regression fixtures.

## Architecture

The current Python implementation remains the semantic reference while the Zig core is built out.

Target split:

- Zig owns parsing, IR, evaluation, simplification, and the C ABI.
- Python owns `openpyxl` integration and Python-native ergonomics.
- TypeScript owns the JS/TS-facing API and Node binding.
- Windows-only fuzzing compares desktop Excel against the native core and promotes failures into explicit regression fixtures.

## Build

Native core:

```powershell
zig build
zig build test
```

Python package:

```powershell
cd packages/python
uv sync
uv run --with pytest pytest tests
```

TypeScript package:

```powershell
cd packages/ts
corepack pnpm install
corepack pnpm build
```

## Orchestration

A root `justfile` is included for common tasks:

- `just --list`
- `just build-core`
- `just test-core`
- `just test-python`
- `just smoke-python-native`
- `just smoke-ts-native`
- `just smoke-native`
- `just build-ts`
- `just fuzz-excel`

The `just` command runner is not required to work on the repo, but it is the intended top-level task interface.
