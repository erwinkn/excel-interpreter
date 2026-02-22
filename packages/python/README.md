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
