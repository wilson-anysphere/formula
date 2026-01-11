from __future__ import annotations

import builtins
import io
import os
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

_ORIGINAL_SOCKET_CREATE_CONNECTION = None
_ORIGINAL_SOCKET_CONNECT = None
_ORIGINAL_SOCKET_CONNECT_EX = None
_ORIGINAL_SOCKET_SENDTO = None
_ORIGINAL_SOCKET_SENDMSG = None
_ORIGINAL_SOCKET_GETADDRINFO = None
_ORIGINAL_SOCKET_GLOBAL_DEFAULT_TIMEOUT = None
_ORIGINAL_SOCKET_SOCKET_CLASS = None
_ORIGINAL__SOCKET_SOCKET = None
_ORIGINAL__SOCKET_SOCKETTYPE = None

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
    "fork",
    "forkpty",
    "posix_spawn",
    "posix_spawnp",
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
    global _ORIGINAL_SOCKET_SENDTO
    global _ORIGINAL_SOCKET_SENDMSG
    global _ORIGINAL_SOCKET_GETADDRINFO
    global _ORIGINAL_SOCKET_GLOBAL_DEFAULT_TIMEOUT
    global _ORIGINAL_SOCKET_SOCKET_CLASS
    global _ORIGINAL__SOCKET_SOCKET
    global _ORIGINAL__SOCKET_SOCKETTYPE

    try:
        import sys

        socket = sys.modules.get("socket")
        if socket is None:
            # Avoid importing `socket` when `network="none"` so scripts cannot
            # bypass the import hook by grabbing a pre-imported module from
            # sys.modules (e.g. to call `socket.socket.sendto`).
            if network == "none":
                return
            import socket as socket  # type: ignore
    except Exception:
        return

    if _ORIGINAL_SOCKET_CREATE_CONNECTION is None:
        _ORIGINAL_SOCKET_CREATE_CONNECTION = getattr(socket, "create_connection", None)
    if _ORIGINAL_SOCKET_CONNECT is None:
        _ORIGINAL_SOCKET_CONNECT = getattr(socket.socket, "connect", None)
    if _ORIGINAL_SOCKET_CONNECT_EX is None:
        _ORIGINAL_SOCKET_CONNECT_EX = getattr(socket.socket, "connect_ex", None)
    if _ORIGINAL_SOCKET_SENDTO is None:
        _ORIGINAL_SOCKET_SENDTO = getattr(socket.socket, "sendto", None)
    if _ORIGINAL_SOCKET_SENDMSG is None:
        _ORIGINAL_SOCKET_SENDMSG = getattr(socket.socket, "sendmsg", None)
    if _ORIGINAL_SOCKET_GETADDRINFO is None:
        _ORIGINAL_SOCKET_GETADDRINFO = getattr(socket, "getaddrinfo", None)
    if _ORIGINAL_SOCKET_GLOBAL_DEFAULT_TIMEOUT is None:
        _ORIGINAL_SOCKET_GLOBAL_DEFAULT_TIMEOUT = getattr(socket, "_GLOBAL_DEFAULT_TIMEOUT", object())
    if _ORIGINAL_SOCKET_SOCKET_CLASS is None:
        _ORIGINAL_SOCKET_SOCKET_CLASS = getattr(socket, "socket", None)

    _socket_mod = getattr(socket, "_socket", None)
    if _socket_mod is not None:
        if _ORIGINAL__SOCKET_SOCKET is None:
            _ORIGINAL__SOCKET_SOCKET = getattr(_socket_mod, "socket", None)
        if _ORIGINAL__SOCKET_SOCKETTYPE is None:
            _ORIGINAL__SOCKET_SOCKETTYPE = getattr(_socket_mod, "SocketType", None)

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
    if _ORIGINAL_SOCKET_SENDTO is not None:
        try:
            socket.socket.sendto = _ORIGINAL_SOCKET_SENDTO  # type: ignore[assignment]
        except Exception:
            pass
    if _ORIGINAL_SOCKET_SENDMSG is not None:
        try:
            socket.socket.sendmsg = _ORIGINAL_SOCKET_SENDMSG  # type: ignore[assignment]
        except Exception:
            pass

    if _socket_mod is not None:
        if _ORIGINAL__SOCKET_SOCKET is not None:
            try:
                setattr(_socket_mod, "socket", _ORIGINAL__SOCKET_SOCKET)
            except Exception:
                pass
        if _ORIGINAL__SOCKET_SOCKETTYPE is not None:
            try:
                setattr(_socket_mod, "SocketType", _ORIGINAL__SOCKET_SOCKETTYPE)
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
        enforce_hostname(_extract_hostname(address))
        return _ORIGINAL_SOCKET_CONNECT(self, address)  # type: ignore[misc]

    def guarded_connect_ex(self, address):  # type: ignore[no-untyped-def]
        enforce_hostname(_extract_hostname(address))
        return _ORIGINAL_SOCKET_CONNECT_EX(self, address)  # type: ignore[misc]

    def guarded_sendto(self, data, *args, **kwargs):  # type: ignore[no-untyped-def]
        # Signature: sendto(bytes[, flags], address)
        address = None
        if "address" in kwargs:
            address = kwargs.get("address")
        elif len(args) == 1:
            address = args[0]
        elif len(args) >= 2:
            address = args[1]

        enforce_hostname(_extract_hostname(address))
        return _ORIGINAL_SOCKET_SENDTO(self, data, *args, **kwargs)  # type: ignore[misc]

    def guarded_sendmsg(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        # Signature: sendmsg(buffers[, ancdata[, flags[, address]]])
        address = None
        if "address" in kwargs:
            address = kwargs.get("address")
        elif len(args) >= 4:
            address = args[3]

        if address is not None:
            enforce_hostname(_extract_hostname(address))
        return _ORIGINAL_SOCKET_SENDMSG(self, *args, **kwargs)  # type: ignore[misc]

    def guarded_create_connection(
        address, timeout=_ORIGINAL_SOCKET_GLOBAL_DEFAULT_TIMEOUT, source_address=None, all_errors=False
    ):  # type: ignore[no-untyped-def]
        """
        Re-implementation of socket.create_connection with allowlist enforcement.

        We avoid delegating to the stdlib implementation directly because it can be
        influenced by user monkeypatching (e.g. overriding socket.getaddrinfo),
        which can otherwise be abused to bypass allowlist checks.
        """

        enforce_hostname(_extract_hostname(address))

        if (
            _ORIGINAL_SOCKET_GETADDRINFO is None
            or _ORIGINAL_SOCKET_SOCKET_CLASS is None
            or not isinstance(address, (tuple, list))
            or len(address) < 2
        ):
            # Fallback to the original implementation (best-effort).
            return _ORIGINAL_SOCKET_CREATE_CONNECTION(address, timeout=timeout, source_address=source_address)  # type: ignore[misc]

        host, port = address[0], address[1]
        errors: list[BaseException] = []

        try:
            addr_infos = _ORIGINAL_SOCKET_GETADDRINFO(host, port, 0, getattr(socket, "SOCK_STREAM", 1))  # type: ignore[misc]
        except Exception as err:
            raise err

        for res in addr_infos:
            af, socktype, proto, _canonname, sa = res
            sock_obj = None
            try:
                sock_obj = _ORIGINAL_SOCKET_SOCKET_CLASS(af, socktype, proto)
                if timeout is not _ORIGINAL_SOCKET_GLOBAL_DEFAULT_TIMEOUT:
                    sock_obj.settimeout(timeout)
                if source_address:
                    sock_obj.bind(source_address)
                _ORIGINAL_SOCKET_CONNECT(sock_obj, sa)  # type: ignore[misc]
                return sock_obj
            except BaseException as err:
                errors.append(err)
                if sock_obj is not None:
                    try:
                        sock_obj.close()
                    except Exception:
                        pass

        if all_errors:
            try:
                raise ExceptionGroup("create_connection failed", errors)  # type: ignore[name-defined]
            except NameError:
                # Python < 3.11; fall back to the last error.
                pass

        if errors:
            raise errors[-1]
        raise OSError("getaddrinfo returns an empty list")

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
    if _ORIGINAL_SOCKET_SENDTO is not None:
        try:
            socket.socket.sendto = guarded_sendto  # type: ignore[assignment]
        except Exception:
            pass
    if _ORIGINAL_SOCKET_SENDMSG is not None:
        try:
            socket.socket.sendmsg = guarded_sendmsg  # type: ignore[assignment]
        except Exception:
            pass

    # Prevent trivial allowlist bypasses via `import _socket; _socket.socket(...)`.
    #
    # We cannot monkeypatch methods on the builtin `_socket.socket` type itself
    # (it's immutable), but we can replace the module-level constructor aliases
    # with a guarded subclass that enforces the same allowlist policy.
    if _socket_mod is not None and isinstance(_ORIGINAL__SOCKET_SOCKET, type):
        BaseSocketType = _ORIGINAL__SOCKET_SOCKET

        class GuardedSocketType(BaseSocketType):  # type: ignore[misc,valid-type]
            __slots__ = ()

            def connect(self, address):  # type: ignore[no-untyped-def]
                enforce_hostname(_extract_hostname(address))
                return super().connect(address)

            def connect_ex(self, address):  # type: ignore[no-untyped-def]
                enforce_hostname(_extract_hostname(address))
                return super().connect_ex(address)

            def sendto(self, data, *args, **kwargs):  # type: ignore[no-untyped-def]
                address = None
                if "address" in kwargs:
                    address = kwargs.get("address")
                elif len(args) == 1:
                    address = args[0]
                elif len(args) >= 2:
                    address = args[1]

                enforce_hostname(_extract_hostname(address))
                return super().sendto(data, *args, **kwargs)

            def sendmsg(self, *args, **kwargs):  # type: ignore[no-untyped-def]
                address = None
                if "address" in kwargs:
                    address = kwargs.get("address")
                elif len(args) >= 4:
                    address = args[3]

                if address is not None:
                    enforce_hostname(_extract_hostname(address))
                return super().sendmsg(*args, **kwargs)

        try:
            setattr(_socket_mod, "socket", GuardedSocketType)
        except Exception:
            pass
        try:
            setattr(_socket_mod, "SocketType", GuardedSocketType)
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
        blocked_roots.update({"socket", "_socket", "ssl", "_ssl", "http", "urllib", "requests"})

    if blocked_roots:

        def guarded_import(name: str, globals=None, locals=None, fromlist=(), level: int = 0):
            root = name.split(".", 1)[0]
            if root in blocked_roots:
                raise PermissionError(f"Import of {root!r} is not permitted")
            return _ORIGINAL_IMPORT(name, globals, locals, fromlist, level)

        builtins.__import__ = guarded_import  # type: ignore[assignment]
