from __future__ import annotations

import ctypes
import sys
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path


@dataclass(frozen=True)
class NativeCoreStatus:
    available: bool
    reason: str
    artifact_hint: str | None = None


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[4]


def _library_candidates() -> list[Path]:
    root = _repo_root()
    if sys.platform == "win32":
        return [root / "zig-out" / "bin" / "excel_interpreter.dll"]
    if sys.platform == "darwin":
        return [
            root / "zig-out" / "lib" / "libexcel_interpreter.dylib",
            root / "zig-out" / "lib" / "excel_interpreter.dylib",
        ]
    return [
        root / "zig-out" / "lib" / "libexcel_interpreter.so",
        root / "zig-out" / "lib" / "excel_interpreter.so",
    ]


def _library_path() -> Path:
    for candidate in _library_candidates():
        if candidate.exists():
            return candidate
    raise FileNotFoundError(
        "Native Zig library not found. Build the repo root with `zig build` first."
    )


@lru_cache(maxsize=1)
def _load_library() -> ctypes.CDLL:
    library = ctypes.CDLL(str(_library_path()))
    library.ei_add_f64.argtypes = [ctypes.c_double, ctypes.c_double]
    library.ei_add_f64.restype = ctypes.c_double
    library.ei_demo_greeting.argtypes = []
    library.ei_demo_greeting.restype = ctypes.c_char_p
    return library


def native_add(lhs: float, rhs: float) -> float:
    library = _load_library()
    return float(library.ei_add_f64(lhs, rhs))


def native_greeting() -> str:
    library = _load_library()
    greeting = library.ei_demo_greeting()
    assert greeting is not None
    return greeting.decode("utf-8")


def get_native_core_status() -> NativeCoreStatus:
    try:
        path = _library_path()
        _load_library()
    except Exception as exc:
        return NativeCoreStatus(
            available=False,
            reason=str(exc),
            artifact_hint="Build the repo root with `zig build`; the shared library is expected under `zig-out`.",
        )

    return NativeCoreStatus(
        available=True,
        reason="Native Zig library is available.",
        artifact_hint=str(path),
    )
