# excel-interpreter-py

Python library for:

- Evaluating Excel formulas against an `openpyxl` workbook (`ExcelInterpreter`).
- Building Excel expression trees in Python with operator overloads (`Expression`, `Cell`, `Range`, `Excel` helpers).
- Rendering expression trees back to Excel formulas with simplification support.

Core entry points:

- `excel_interpreter.ExcelInterpreter`
- `excel_interpreter.ExcelReader`
- `excel_interpreter.ExcelWriter`
- `excel_interpreter.parse_formula`
- `excel_interpreter.Excel` / expression nodes
- `excel_interpreter.get_native_core_status`
- `excel_interpreter.native_add`
- `excel_interpreter.native_greeting`

## Native Core Direction

The Python package remains the semantic reference implementation while the Zig core is built out under `packages/core`.

The planned split is:

- Python keeps workbook integration, `openpyxl` interoperability, and high-level ergonomics.
- Zig will eventually own parsing, evaluation, simplification, and the stable C ABI.
- The Python wrapper will call the Zig core through a compiled extension module once the ABI stabilizes.
