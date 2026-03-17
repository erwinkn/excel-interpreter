# @erwinkn/excel-interpreter-ts

TypeScript primitives for Excel formula interpretation.

Current package surface:

- `ExpressionBuilder`: build Excel expressions programmatically.
- `evaluateFormula`: entry point for workbook-aware evaluation (placeholder).
- `getNativeBindingStatus`: reports the current native binding scaffold status.
- `nativeAdd`: demo native math call into the Zig DLL.
- `nativeGreeting`: demo native string call into the Zig DLL.

## Native Core Direction

The native implementation is being built in `packages/core` with Zig. The TS package will eventually load a Node binding backed by the shared C ABI from that package.

## Build

```bash
pnpm build
```
