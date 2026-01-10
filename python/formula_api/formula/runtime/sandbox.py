from __future__ import annotations

import builtins
import io
import os
import threading
import sysconfig
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

_SOCKET_THREAD_LOCAL = threading.local()
_ORIGINAL_SOCKET_CREATE_CONNECTION = None
_ORIGINAL_SOCKET_CONNECT = None
_ORIGINAL_SOCKET_CONNECT_EX = None

try:  # pragma: no cover - environment dependent
    _IMPORT_ALLOWLIST_ROOTS = tuple(
        sorted(
            {
                os.path.abspath(path)
                for path in (
                    sysconfig.get_path("stdlib"),
                    sysconfig.get_path("platstdlib"),
                    sysconfig.get_path("purelib"),
                    sysconfig.get_path("platlib"),
                )
                if path
            }
        )
    )
except Exception:  # pragma: no cover - environment dependent
    _IMPORT_ALLOWLIST_ROOTS = ()

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


def _is_allowed_import_path(path: Any) -> bool:
    """Allow Python's import system to read stdlib/site-packages when filesystem='none'."""

    if not _IMPORT_ALLOWLIST_ROOTS:
        return False

    try:
        fs_path = os.fspath(path)
    except Exception:
        return False

    if isinstance(fs_path, bytes):
        try:
            fs_path = fs_path.decode("utf-8", errors="ignore")
        except Exception:
            return False

    if not isinstance(fs_path, str) or not fs_path:
        return False

    abs_path = os.path.abspath(fs_path)
    for root in _IMPORT_ALLOWLIST_ROOTS:
        if abs_path == root or abs_path.startswith(root + os.sep):
            return True
    return False


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


def apply_cpu_time_limit(*, max_cpu_seconds: int | None = None, timeout_ms: int | None = None) -> None:
    """Best-effort CPU time limit (native Python only).

    Notes:
    - `resource.RLIMIT_CPU` is Unix-specific and uses seconds granularity.
    - This is complementary to host-side wall clock timeouts (timeoutMs). CPU time
      limits protect against tight loops that burn CPU while still making some
      progress (e.g. printing / RPC).
    """

    if max_cpu_seconds is None and timeout_ms is None:
        return

    limit_seconds = None
    if isinstance(max_cpu_seconds, int) and max_cpu_seconds > 0:
        limit_seconds = max_cpu_seconds
    elif isinstance(timeout_ms, int) and timeout_ms > 0:
        # RLIMIT_CPU is integer seconds; clamp to 1s minimum.
        limit_seconds = max(1, (timeout_ms + 999) // 1000)

    if not limit_seconds:
        return

    try:
        import resource  # Unix only

        # Provide a small grace window between soft/hard limits so the runtime
        # can exit cleanly after SIGXCPU (default action is termination).
        resource.setrlimit(resource.RLIMIT_CPU, (limit_seconds, limit_seconds + 1))
    except Exception:
        # Best-effort: not all platforms support resource or RLIMIT_CPU.
        pass


def _normalize_hostname(host: str) -> str:
    host = host.strip().lower()
    if host.startswith("[") and host.endswith("]"):
        host = host[1:-1]
    return host.rstrip(".")


def _extract_hostname(address: Any) -> str | None:
    host: Any
    if isinstance(address, (tuple, list)):
        if not address:
            return None
        host = address[0]
    else:
        host = address

    if isinstance(host, bytes):
        try:
            host = host.decode("utf-8", errors="ignore")
        except Exception:
            return None

    if isinstance(host, str):
        normalized = _normalize_hostname(host)
        return normalized or None

    return None


def _parse_network_allowlist(raw: Any) -> tuple[set[str], tuple[str, ...]]:
    """Return (exact_matches, wildcard_suffixes)."""

    if not isinstance(raw, (list, tuple, set)):
        return set(), ()

    exact: set[str] = set()
    wildcard_suffixes: list[str] = []

    for item in raw:
        if not isinstance(item, str):
            continue
        entry = _normalize_hostname(item)
        if not entry:
            continue
        if entry.startswith("*.") and len(entry) > 2:
            suffix = entry[2:]
            if suffix:
                wildcard_suffixes.append(suffix)
        else:
            exact.add(entry)

    # Keep deterministic ordering for easier debugging.
    return exact, tuple(sorted(set(wildcard_suffixes)))


def _hostname_in_allowlist(hostname: str, exact: set[str], wildcard_suffixes: tuple[str, ...]) -> bool:
    if hostname in exact:
        return True
    for suffix in wildcard_suffixes:
        # "*.example.com" should match "api.example.com" (and deeper), but not
        # the bare "example.com".
        if hostname != suffix and hostname.endswith(f".{suffix}"):
            return True
    return False


def _apply_socket_network_policy(network: str, permissions: Dict[str, Any]) -> None:
    """Patch socket connection APIs for network allowlist enforcement."""

    global _ORIGINAL_SOCKET_CREATE_CONNECTION
    global _ORIGINAL_SOCKET_CONNECT
    global _ORIGINAL_SOCKET_CONNECT_EX

    try:
        import socket  # type: ignore
    except Exception:
        return

    if _ORIGINAL_SOCKET_CREATE_CONNECTION is None:
        _ORIGINAL_SOCKET_CREATE_CONNECTION = getattr(socket, "create_connection", None)
    if _ORIGINAL_SOCKET_CONNECT is None:
        _ORIGINAL_SOCKET_CONNECT = getattr(socket.socket, "connect", None)
    if _ORIGINAL_SOCKET_CONNECT_EX is None:
        _ORIGINAL_SOCKET_CONNECT_EX = getattr(socket.socket, "connect_ex", None)

    # Restore first so apply_sandbox() can both tighten and loosen restrictions.
    if _ORIGINAL_SOCKET_CREATE_CONNECTION is not None:
        try:
            socket.create_connection = _ORIGINAL_SOCKET_CREATE_CONNECTION  # type: ignore[assignment]
        except Exception:
            pass
    if _ORIGINAL_SOCKET_CONNECT is not None:
        try:
            socket.socket.connect = _ORIGINAL_SOCKET_CONNECT  # type: ignore[assignment]
        except Exception:
            pass
    if _ORIGINAL_SOCKET_CONNECT_EX is not None:
        try:
            socket.socket.connect_ex = _ORIGINAL_SOCKET_CONNECT_EX  # type: ignore[assignment]
        except Exception:
            pass

    if network == "full":
        return

    raw_allowlist = permissions.get("networkAllowlist", permissions.get("network_allowlist", []))
    exact, wildcard_suffixes = _parse_network_allowlist(raw_allowlist) if network == "allowlist" else (set(), ())

    def enforce_hostname(hostname: str | None) -> None:
        if hostname is None or not _hostname_in_allowlist(hostname, exact, wildcard_suffixes):
            raise PermissionError(f"Network access to {hostname!r} is not permitted")

    def guarded_connect(self, address):  # type: ignore[no-untyped-def]
        # When socket.create_connection() is used, it resolves DNS and calls
        # socket.connect() with IP literals. Preserve the original allowlist
        # decision made on the hostname by skipping checks in that call chain.
        if getattr(_SOCKET_THREAD_LOCAL, "bypass_connect_check", 0):
            return _ORIGINAL_SOCKET_CONNECT(self, address)  # type: ignore[misc]

        enforce_hostname(_extract_hostname(address))
        return _ORIGINAL_SOCKET_CONNECT(self, address)  # type: ignore[misc]

    def guarded_connect_ex(self, address):  # type: ignore[no-untyped-def]
        if getattr(_SOCKET_THREAD_LOCAL, "bypass_connect_check", 0):
            return _ORIGINAL_SOCKET_CONNECT_EX(self, address)  # type: ignore[misc]

        enforce_hostname(_extract_hostname(address))
        return _ORIGINAL_SOCKET_CONNECT_EX(self, address)  # type: ignore[misc]

    def guarded_create_connection(address, *args, **kwargs):  # type: ignore[no-untyped-def]
        enforce_hostname(_extract_hostname(address))

        _SOCKET_THREAD_LOCAL.bypass_connect_check = getattr(_SOCKET_THREAD_LOCAL, "bypass_connect_check", 0) + 1
        try:
            return _ORIGINAL_SOCKET_CREATE_CONNECTION(address, *args, **kwargs)  # type: ignore[misc]
        finally:
            _SOCKET_THREAD_LOCAL.bypass_connect_check -= 1

    if _ORIGINAL_SOCKET_CREATE_CONNECTION is not None:
        try:
            socket.create_connection = guarded_create_connection  # type: ignore[assignment]
        except Exception:
            pass
    if _ORIGINAL_SOCKET_CONNECT is not None:
        try:
            socket.socket.connect = guarded_connect  # type: ignore[assignment]
        except Exception:
            pass
    if _ORIGINAL_SOCKET_CONNECT_EX is not None:
        try:
            socket.socket.connect_ex = guarded_connect_ex  # type: ignore[assignment]
        except Exception:
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

    # Network allowlist enforcement may need to import socket modules. Apply it
    # after restoring the original import hook and before tightening filesystem
    # restrictions so the standard library can load.
    _apply_socket_network_policy(network, permissions)

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

        def guarded_open(*args: Any, **kwargs: Any) -> Any:
            mode = _extract_mode(args, kwargs)
            if _is_write_mode(mode):
                raise PermissionError("Filesystem access is not permitted")

            target = None
            if args:
                target = args[0]
            elif "file" in kwargs:
                target = kwargs.get("file")

            if not _is_allowed_import_path(target):
                raise PermissionError("Filesystem access is not permitted")

            return _ORIGINAL_OPEN(*args, **kwargs)

        builtins.open = guarded_open  # type: ignore[assignment]
        io.open = guarded_open  # type: ignore[assignment]
        if _io_builtin is not None:
            _io_builtin.open = guarded_open  # type: ignore[attr-defined]

    # ---- os.* filesystem policy -------------------------------------------
    # Restore patched functions first so repeated apply_sandbox() calls can
    # loosen restrictions.
    for name, fn in _ORIGINAL_OS.items():
        try:
            setattr(os, name, fn)
        except Exception:
            pass

    if filesystem == "none":
        if "listdir" in _ORIGINAL_OS:

            def guarded_listdir(path: Any = None) -> Any:
                target = "." if path is None else path
                if not _is_allowed_import_path(target):
                    raise PermissionError("Filesystem access is not permitted")
                return _ORIGINAL_OS["listdir"](target)

            os.listdir = guarded_listdir  # type: ignore[assignment]

        if "scandir" in _ORIGINAL_OS:

            def guarded_scandir(path: Any = None) -> Any:
                target = "." if path is None else path
                if not _is_allowed_import_path(target):
                    raise PermissionError("Filesystem access is not permitted")
                return _ORIGINAL_OS["scandir"](target)

            os.scandir = guarded_scandir  # type: ignore[assignment]

        for name in _OS_FS_FUNCS:
            if name in {"listdir", "scandir"}:
                continue
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
