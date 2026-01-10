from __future__ import annotations

import builtins
from typing import Any, Dict


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
    - This is *not* a hardened security boundary.
    - In Pyodide, filesystem access is limited to the in-memory FS, but we still
      block `open()` by default to keep behavior consistent across runtimes.
    """

    filesystem = permissions.get("filesystem", "none")
    network = permissions.get("network", "none")

    block_process_execution = filesystem == "none" or network == "none"

    if filesystem == "none":

        def blocked_open(*_args: Any, **_kwargs: Any) -> Any:
            raise PermissionError("Filesystem access is not permitted")

        builtins.open = blocked_open  # type: ignore[assignment]

        try:  # pragma: no cover - platform dependent
            import os

            def blocked_fs(*_args: Any, **_kwargs: Any) -> Any:
                raise PermissionError("Filesystem access is not permitted")

            # Common filesystem-related helpers that allow reads/writes without
            # going through builtins.open.
            for attr in (
                "open",
                "fdopen",
                "startfile",
                "listdir",
                "scandir",
            ):
                if hasattr(os, attr):
                    setattr(os, attr, blocked_fs)
        except Exception:
            pass

    if block_process_execution:
        try:  # pragma: no cover - platform dependent
            import os

            def blocked_process(*_args: Any, **_kwargs: Any) -> Any:
                raise PermissionError("Process execution is not permitted")

            for attr in (
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
            ):
                if hasattr(os, attr):
                    setattr(os, attr, blocked_process)
        except Exception:
            pass

    # Always block interactive input (native runtime stdin is used for RPC, and
    # Pyodide has no meaningful stdin).
    def blocked_input(*_args: Any, **_kwargs: Any) -> Any:
        raise PermissionError("Interactive input is not permitted")

    builtins.input = blocked_input  # type: ignore[assignment]

    blocked_roots = set()
    # `subprocess` is an easy escape hatch for both filesystem and network access.
    if filesystem == "none" or network == "none":
        blocked_roots.add("subprocess")
    if network == "none":
        blocked_roots.update({"socket", "ssl", "http", "urllib", "requests"})

    if blocked_roots:
        original_import = builtins.__import__

        def guarded_import(name: str, globals=None, locals=None, fromlist=(), level: int = 0):
            root = name.split(".", 1)[0]
            if root in blocked_roots:
                raise PermissionError(f"Import of {root!r} is not permitted")
            return original_import(name, globals, locals, fromlist, level)

        builtins.__import__ = guarded_import  # type: ignore[assignment]
