# Shared Regression Corpus

This directory is reserved for cross-language, deterministic test cases.

Intended usage:

- The current Python test suite defines the semantic baseline.
- Native-core tests should gradually absorb stable cases from Python into machine-readable fixtures here.
- Python and TS wrappers should both run the same regression corpus once they call into the Zig core.
- Fuzzing failures from the Windows Excel oracle should be reduced into explicit cases under `tests/corpus/regressions`.

Suggested fixture shape:

```json
{
  "name": "sum-basic-range",
  "formula": "=SUM(A1:A3)",
  "sheet": "Sheet1",
  "cells": {
    "Sheet1!A1": 1,
    "Sheet1!A2": 2,
    "Sheet1!A3": 3
  },
  "expected": 6
}
```
