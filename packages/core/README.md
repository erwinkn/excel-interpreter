# excel-interpreter-core

Native Zig core for the Excel interpreter monorepo.

Planned responsibilities:

- Tokenization and parsing for Excel formulas.
- Pure expression IR and rendering.
- Workbook snapshot evaluation runtime.
- Stable C ABI for Python and Node wrappers.
- Shared regression corpus and Windows-only Excel oracle fuzzing.

Current scaffold:

- `src/syntax`: parsing entry points and tokenizer/parser placeholders.
- `src/ir`: pure expression tree types.
- `src/runtime`: value model and evaluator.
- `src/ffi`: C ABI layer used by language wrappers.
- `include`: public C headers.
- `fuzz/windows`: manual Excel oracle harness design notes.

Build from the repository root:

```powershell
zig build
zig build test
```
