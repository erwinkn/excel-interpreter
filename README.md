# excel-interpreter

Monorepo for Excel formula tooling in Python and TypeScript:

- `python/excel_interpreter_py`: Python library for workbook-aware formula evaluation and expression building.
- `ts/excel-interpreter-ts`: TypeScript library for formula evaluation and expression building.

## Goal

Provide shared, cross-language capabilities for:

- Parsing and evaluating Excel formulas in workbook context.
- Programmatically building valid Excel expressions.
- Extensible function/runtime implementations for custom usage.

## Quick start

Python:

```bash
cd python/excel_interpreter_py
uv sync
```

TypeScript:

```bash
cd ts/excel-interpreter-ts
npm install
npm run build
```
