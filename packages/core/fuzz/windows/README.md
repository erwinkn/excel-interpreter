# Windows Excel Oracle Fuzzing

This harness is intentionally manual and Windows-only.

Scope:

- Compare the native interpreter against desktop Excel.
- Generate failures locally on a developer workstation.
- Minimize failures into explicit regression cases.
- Never run Excel automation in CI, containers, or cloud VMs.

Planned layout:

- `generator`: produces formulas and workbook snapshots within the supported subset.
- `oracle`: drives Excel through COM automation and captures recalculated outputs.
- `reducer`: shrinks mismatches to minimal repro cases.
- `promoter`: writes regression fixtures into `tests/corpus/regressions`.

Expected workflow:

1. Build the native core from the repository root with `zig build`.
2. Generate candidate formulas and workbook states.
3. Evaluate them in desktop Excel and in the native interpreter.
4. Minimize any mismatch.
5. Commit the minimized case as a deterministic regression fixture.
