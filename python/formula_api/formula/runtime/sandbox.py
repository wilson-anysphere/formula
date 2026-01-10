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

    if filesystem == "none":

        def blocked_open(*_args: Any, **_kwargs: Any) -> Any:
            raise PermissionError("Filesystem access is not permitted")

        builtins.open = blocked_open  # type: ignore[assignment]

    # Always block interactive input (native runtime stdin is used for RPC, and
    # Pyodide has no meaningful stdin).
    def blocked_input(*_args: Any, **_kwargs: Any) -> Any:
        raise PermissionError("Interactive input is not permitted")

    builtins.input = blocked_input  # type: ignore[assignment]

    blocked_roots = set()
    if filesystem == "none":
        blocked_roots.update({"os", "pathlib", "shutil", "subprocess"})
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

