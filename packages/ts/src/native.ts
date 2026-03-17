import { existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

export interface NativeBindingStatus {
  available: boolean;
  reason: string;
  artifactHint?: string;
}

type NativeSymbols = {
  ei_add_f64(lhs: number, rhs: number): number;
  ei_demo_greeting(): string;
};

type NativeBinding = {
  path: string;
  symbols: NativeSymbols;
};

let bindingPromise: Promise<NativeBinding> | null = null;

function moduleDirname(): string {
  return dirname(fileURLToPath(import.meta.url));
}

function libraryCandidates(): string[] {
  const root = resolve(moduleDirname(), "../../../");
  if (process.platform === "win32") {
    return [resolve(root, "zig-out", "bin", "excel_interpreter.dll")];
  }
  if (process.platform === "darwin") {
    return [
      resolve(root, "zig-out", "lib", "libexcel_interpreter.dylib"),
      resolve(root, "zig-out", "lib", "excel_interpreter.dylib"),
    ];
  }
  return [
    resolve(root, "zig-out", "lib", "libexcel_interpreter.so"),
    resolve(root, "zig-out", "lib", "excel_interpreter.so"),
  ];
}

function libraryPath(): string {
  for (const candidate of libraryCandidates()) {
    if (existsSync(candidate)) {
      return candidate;
    }
  }
  throw new Error("Native Zig library not found. Build the repo root with `zig build` first.");
}

async function loadBinding(): Promise<NativeBinding> {
  if (bindingPromise) {
    return bindingPromise;
  }

  bindingPromise = (async () => {
    if (typeof Bun === "undefined") {
      throw new Error("This native TS demo currently requires the Bun runtime.");
    }

    const path = libraryPath();
    const { dlopen, FFIType } = await import("bun:ffi");
    const handle = dlopen(path, {
      ei_add_f64: {
        args: [FFIType.f64, FFIType.f64],
        returns: FFIType.f64,
      },
      ei_demo_greeting: {
        args: [],
        returns: FFIType.cstring,
      },
    });

    return {
      path,
      symbols: handle.symbols as unknown as NativeSymbols,
    };
  })();

  return bindingPromise;
}

export function getNativeBindingStatus(): NativeBindingStatus {
  try {
    const path = libraryPath();
    return {
      available: true,
      reason: typeof Bun === "undefined"
        ? "Native Zig library is present, but runtime loading requires Bun for this demo."
        : "Native Zig library is available.",
      artifactHint: path,
    };
  } catch (error) {
    return {
      available: false,
      reason: error instanceof Error ? error.message : String(error),
      artifactHint: "Build the repo root with `zig build`; the shared library is expected under `zig-out`.",
    };
  }
}

export async function nativeAdd(lhs: number, rhs: number): Promise<number> {
  const binding = await loadBinding();
  return binding.symbols.ei_add_f64(lhs, rhs);
}

export async function nativeGreeting(): Promise<string> {
  const binding = await loadBinding();
  return binding.symbols.ei_demo_greeting();
}
