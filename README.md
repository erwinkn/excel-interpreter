# excel-interpreter

Monorepo for Excel formula tooling in Python and TypeScript:

- `packages/python`: Python library for workbook-aware formula evaluation and expression building.
- `packages/ts`: TypeScript library for formula evaluation and expression building.

## Goal

Provide shared, cross-language capabilities for:

- Parsing and evaluating Excel formulas in workbook context.
- Programmatically building valid Excel expressions.
- Extensible function/runtime implementations for custom usage.

## Quick start

Python:

```bash
cd packages/python
uv sync
```

TypeScript:

```bash
cd packages/ts
corepack pnpm install
corepack pnpm build
```
