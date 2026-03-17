declare module "bun:ffi" {
  export const suffix: string;

  export const FFIType: {
    readonly f64: "f64";
    readonly cstring: "cstring";
  };

  export function dlopen(
    path: string,
    symbols: Record<string, { args: unknown[]; returns: unknown }>,
  ): {
    symbols: Record<string, (...args: unknown[]) => unknown>;
    close(): void;
  };
}
