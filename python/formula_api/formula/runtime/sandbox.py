from __future__ import annotations

import builtins
import io
import os
from typing import Any, Dict

# Capture original references at import time so apply_sandbox() can be called
# repeatedly (Pyodide worker) and can both tighten and loosen restrictions.
_ORIGINAL_OPEN = builtins.open
_ORIGINAL_IMPORT = builtins.__import__

try:  # pragma: no cover - implementation dependent
    import _io as _io_builtin  # type: ignore
except Exception:  # pragma: no cover - native only
    _io_builtin = None

_ORIGINAL_IO_OPEN = getattr(io, "open", None)
_ORIGINAL__IO_OPEN = getattr(_io_builtin, "open", None) if _io_builtin else None

_OS_FS_FUNCS = (
    # File descriptor based access.
    "open",
    "fdopen",
    # Directory enumeration.
    "listdir",
    "scandir",
    # Destructive operations / mutations.
    "remove",
    "unlink",
    "rmdir",
    "mkdir",
    "makedirs",
    "rename",
    "replace",
    "link",
    "symlink",
    "readlink",
    # Platform-specific helpers.
    "startfile",
)

_OS_PROCESS_FUNCS = (
    "system",
    "popen",
    "spawnl",
    "spawnle",
    "spawnlp",
    "spawnlpe",
    "spawnv",
    "spawnve",
    "spawnvp",
    "spawnvpe",
    "execl",
    "execle",
    "execlp",
    "execlpe",
    "execv",
    "execve",
    "execvp",
    "execvpe",
    "startfile",
)

_ORIGINAL_OS = {name: getattr(os, name) for name in set(_OS_FS_FUNCS + _OS_PROCESS_FUNCS) if hasattr(os, name)}


def _normalize_permission(value: Any, allowed: set[str], default: str) -> str:
    if not isinstance(value, str):
        return default
    lowered = value.lower()
    return lowered if lowered in allowed else default


def _filesystem_permission(permissions: Dict[str, Any]) -> str:
    # Support the docs/08-macro-compatibility.md key casing as well.
    raw = permissions.get("filesystem", permissions.get("fileSystem", "none"))
    return _normalize_permission(raw, {"none", "read", "readwrite"}, "none")


def _network_permission(permissions: Dict[str, Any]) -> str:
    raw = permissions.get("network", "none")
    return _normalize_permission(raw, {"none", "allowlist", "full"}, "none")


def _is_write_mode(mode: str) -> bool:
    return any(ch in mode for ch in ("w", "a", "x", "+"))


def _extract_mode(args: tuple[Any, ...], kwargs: dict[str, Any]) -> str:
    if "mode" in kwargs:
        mode = kwargs.get("mode")
        return mode if isinstance(mode, str) else "r"
    if len(args) >= 2 and isinstance(args[1], str):
        return args[1]
    return "r"


def apply_memory_limit(max_memory_bytes: int | None) -> None:
    """Best-effort address space limit (native Python only)."""

    if not max_memory_bytes:
        return

    try:
        import resource  # Unix only

        # RLIMIT_AS limits the total available address space.
        resource.setrlimit(resource.RLIMIT_AS, (max_memory_bytes, max_memory_bytes))
    except Exception:
        # Best-effort: not all platforms support resource or RLIMIT_AS.
        pass


def apply_sandbox(permissions: Dict[str, Any]) -> None:
    """
    Apply lightweight sandbox restrictions for Python scripts.

    Notes:
    - This is not a hardened security boundary.
    - In native Python, stdin is used for RPC, so input() is always disabled.
    - Process execution is treated as an escape hatch and is only permitted when
      both filesystem + network are fully enabled.
    """

    filesystem = _filesystem_permission(permissions)
    network = _network_permission(permissions)

    # Process execution is a common escape hatch for both filesystem and network
    # restrictions (e.g. spawning `curl` or using shell redirection). Only allow
    # it when both network + filesystem are explicitly fully enabled.
    block_process_execution = not (filesystem == "readwrite" and network == "full")

    def blocked_input(*_args: Any, **_kwargs: Any) -> Any:
        raise PermissionError("Interactive input is not permitted")

    builtins.input = blocked_input  # type: ignore[assignment]

    # Reset import hook first so we can reconfigure it based on current permissions.
    builtins.__import__ = _ORIGINAL_IMPORT  # type: ignore[assignment]

    def blocked_fs(*_args: Any, **_kwargs: Any) -> Any:
        raise PermissionError("Filesystem access is not permitted")

    def blocked_fs_write(*_args: Any, **_kwargs: Any) -> Any:
        raise PermissionError("Filesystem write access is not permitted")

    def blocked_process(*_args: Any, **_kwargs: Any) -> Any:
        raise PermissionError("Process execution is not permitted")

    # ---- open() policy -----------------------------------------------------
    if filesystem == "readwrite":
        builtins.open = _ORIGINAL_OPEN  # type: ignore[assignment]
        if _ORIGINAL_IO_OPEN is not None:
            io.open = _ORIGINAL_IO_OPEN  # type: ignore[assignment]
        if _io_builtin is not None and _ORIGINAL__IO_OPEN is not None:
            _io_builtin.open = _ORIGINAL__IO_OPEN  # type: ignore[attr-defined]
    elif filesystem == "read":

        def guarded_open(*args: Any, **kwargs: Any) -> Any:
            mode = _extract_mode(args, kwargs)
            if _is_write_mode(mode):
                raise PermissionError("Filesystem write access is not permitted")
            return _ORIGINAL_OPEN(*args, **kwargs)

        builtins.open = guarded_open  # type: ignore[assignment]
        io.open = guarded_open  # type: ignore[assignment]
        if _io_builtin is not None:
            _io_builtin.open = guarded_open  # type: ignore[attr-defined]
    else:
        builtins.open = blocked_fs  # type: ignore[assignment]
        io.open = blocked_fs  # type: ignore[assignment]
        if _io_builtin is not None:
            _io_builtin.open = blocked_fs  # type: ignore[attr-defined]

    # ---- os.* filesystem policy -------------------------------------------
    # Restore patched functions first so repeated apply_sandbox() calls can
    # loosen restrictions.
    for name, fn in _ORIGINAL_OS.items():
        try:
            setattr(os, name, fn)
        except Exception:
            pass

    if filesystem == "none":
        for name in _OS_FS_FUNCS:
            if hasattr(os, name):
                try:
                    setattr(os, name, blocked_fs)
                except Exception:
                    pass
    elif filesystem == "read":
        # Directory enumeration is permitted in read-only mode, but destructive
        # operations are not.
        for name in (
            "remove",
            "unlink",
            "rmdir",
            "mkdir",
            "makedirs",
            "rename",
            "replace",
            "link",
            "symlink",
        ):
            if hasattr(os, name):
                try:
                    setattr(os, name, blocked_fs_write)
                except Exception:
                    pass

        if "open" in _ORIGINAL_OS:

            def guarded_os_open(path: Any, flags: Any, *args: Any, **kwargs: Any) -> Any:
                try:
                    flags_i = int(flags)
                except Exception:
                    raise PermissionError("Filesystem write access is not permitted")

                write_bits = (
                    os.O_WRONLY
                    | os.O_RDWR
                    | getattr(os, "O_APPEND", 0)
                    | getattr(os, "O_CREAT", 0)
                    | getattr(os, "O_TRUNC", 0)
                    | getattr(os, "O_EXCL", 0)
                )
                if flags_i & write_bits:
                    raise PermissionError("Filesystem write access is not permitted")
                return _ORIGINAL_OS["open"](path, flags, *args, **kwargs)

            os.open = guarded_os_open  # type: ignore[assignment]

        if "fdopen" in _ORIGINAL_OS:

            def guarded_fdopen(fd: Any, *args: Any, **kwargs: Any) -> Any:
                mode = _extract_mode(args, kwargs)
                if _is_write_mode(mode):
                    raise PermissionError("Filesystem write access is not permitted")
                return _ORIGINAL_OS["fdopen"](fd, *args, **kwargs)

            os.fdopen = guarded_fdopen  # type: ignore[assignment]

    # ---- os.* process execution policy ------------------------------------
    if block_process_execution:
        for name in _OS_PROCESS_FUNCS:
            if hasattr(os, name):
                try:
                    setattr(os, name, blocked_process)
                except Exception:
                    pass

    blocked_roots = set()
    if block_process_execution:
        blocked_roots.add("subprocess")
    if network == "none":
        blocked_roots.update({"socket", "ssl", "http", "urllib", "requests"})

    if blocked_roots:

        def guarded_import(name: str, globals=None, locals=None, fromlist=(), level: int = 0):
            root = name.split(".", 1)[0]
            if root in blocked_roots:
                raise PermissionError(f"Import of {root!r} is not permitted")
            return _ORIGINAL_IMPORT(name, globals, locals, fromlist, level)

        builtins.__import__ = guarded_import  # type: ignore[assignment]
