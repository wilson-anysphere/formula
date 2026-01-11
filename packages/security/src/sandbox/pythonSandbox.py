import builtins
import contextlib
import io
import json
import os
import sys
import traceback
import urllib.parse


class PermissionDenied(Exception):
    def __init__(self, request: dict, reason: str):
        super().__init__(reason)
        self.request = request
        self.reason = reason


class OutputLimitExceeded(Exception):
    def __init__(self, max_bytes: int):
        super().__init__(f"Sandbox output exceeded limit of {max_bytes} bytes")
        self.max_bytes = max_bytes


class LimitedStringIO(io.StringIO):
    def __init__(self, max_bytes: int):
        super().__init__()
        self.max_bytes = max_bytes
        self._bytes_written = 0

    def write(self, s):  # type: ignore[override]
        if not isinstance(s, str):
            s = str(s)

        encoded = s.encode("utf-8", errors="replace")
        remaining = self.max_bytes - self._bytes_written
        if remaining <= 0:
            raise OutputLimitExceeded(self.max_bytes)

        if len(encoded) > remaining:
            truncated = encoded[:remaining].decode("utf-8", errors="ignore")
            super().write(truncated)
            self._bytes_written = self.max_bytes
            raise OutputLimitExceeded(self.max_bytes)

        self._bytes_written += len(encoded)
        return super().write(s)


def _format_exc_without_source() -> str:
    """
    Format the current exception without reading files from disk.

    Python's default traceback formatting consults `linecache`, which performs
    filesystem reads (os.stat/open) to include source code context. Our sandbox
    intercepts filesystem calls, so traceback formatting must avoid reading
    arbitrary files (including this runner).
    """

    try:
        exc_type, exc_value, exc_tb = sys.exc_info()
        if exc_value is None:
            return ""
        tb = traceback.TracebackException(
            exc_type,
            exc_value,
            exc_tb,
            lookup_lines=False,
            capture_locals=False,
        )
        return "".join(tb.format())
    except Exception:
        return ""


def _is_within_scope(target_path: str, scope_path: str) -> bool:
    abs_target = os.path.abspath(target_path)
    abs_scope = os.path.abspath(scope_path)
    try:
        common = os.path.commonpath([abs_scope, abs_target])
    except ValueError:
        return False
    return common == abs_scope


def _network_target_allowed(host: str, port, allowlist_entries: list[str]) -> bool:
    host = host.lower()
    for entry in allowlist_entries:
        entry = str(entry).strip()
        if not entry:
            continue

        if "://" in entry:
            parsed = urllib.parse.urlparse(entry)
            if not parsed.hostname:
                continue
            entry_host = parsed.hostname.lower()
            entry_port = parsed.port
            if entry_port is None:
                if parsed.scheme == "https":
                    entry_port = 443
                elif parsed.scheme == "http":
                    entry_port = 80
            if host != entry_host:
                continue
            if port is None or entry_port is None or entry_port == port:
                return True
            continue

        if entry.startswith("*."):
            suffix = entry[2:].lower()
            if host == suffix or host.endswith("." + suffix):
                return True
            continue

        if host == entry.lower():
            return True

    return False


def _check_permission(permissions: dict, request: dict):
    kind = request.get("kind")

    if kind == "filesystem":
        fs_perm = permissions.get("filesystem", {}) if permissions else {}
        read_scopes = [os.path.abspath(p) for p in fs_perm.get("read", [])] + [
            os.path.abspath(p) for p in fs_perm.get("readwrite", [])
        ]
        write_scopes = [os.path.abspath(p) for p in fs_perm.get("readwrite", [])]

        abs_path = os.path.abspath(str(request.get("path", "")))
        access = "readwrite" if request.get("access") == "readwrite" else "read"

        if access == "readwrite":
            if any(_is_within_scope(abs_path, scope) for scope in write_scopes):
                return True, None
            return False, f"Filesystem write access denied for {abs_path}"

        if any(_is_within_scope(abs_path, scope) for scope in read_scopes):
            return True, None
        return False, f"Filesystem read access denied for {abs_path}"

    if kind == "network":
        net_perm = permissions.get("network", {}) if permissions else {}
        mode = net_perm.get("mode", "none")
        url = str(request.get("url", ""))

        if mode == "full":
            return True, None
        if mode == "none":
            return False, f"Network access denied for {url}"

        allowlist = net_perm.get("allowlist", [])
        parsed = urllib.parse.urlparse(url)
        host = parsed.hostname
        port = parsed.port
        if host is None:
            return False, f"Network access denied for {url}"
        if _network_target_allowed(host, port, allowlist):
            return True, None
        return False, f"Network access denied for {url}"

    if kind == "clipboard":
        return (permissions.get("clipboard") is True), "Clipboard permission denied"
    if kind == "notifications":
        return (permissions.get("notifications") is True), "Notifications permission denied"
    if kind == "automation":
        return (permissions.get("automation") is True), "Automation permission denied"

    return False, f"Unknown permission kind: {kind}"


def _main():
    payload = {}
    principal = {"type": "python", "id": "unknown"}
    audit_events = []

    def audit(event_type: str, success: bool, metadata=None):
        audit_events.append(
            {
                "eventType": event_type,
                "actor": principal,
                "success": bool(success),
                "metadata": metadata or {},
            }
        )

    def ensure(request: dict):
        allowed, reason = _check_permission(permissions, request)
        audit(
            "security.permission.checked",
            allowed,
            {"request": request, **({} if allowed else {"reason": reason})},
        )
        if allowed:
            return

        audit("security.permission.denied", False, {"request": request, "reason": reason})
        raise PermissionDenied(request=request, reason=str(reason))

    try:
        payload = json.loads(sys.stdin.read() or "{}")
    except Exception:
        payload = {}

    if isinstance(payload.get("principal"), dict):
        principal = {
            "type": str(payload.get("principal", {}).get("type", "python")),
            "id": str(payload.get("principal", {}).get("id", "unknown")),
        }

    permissions = payload.get("permissions", {}) if isinstance(payload.get("permissions"), dict) else {}
    timeout_ms = int(payload.get("timeoutMs", 5000))
    memory_mb = int(payload.get("memoryMb", 128))
    max_output_bytes = int(payload.get("maxOutputBytes", 128 * 1024))
    code = payload.get("code", "")

    # Best-effort resource limits (Linux/macOS). If unavailable, continue.
    try:
        import resource

        cpu_seconds = max(1, int((timeout_ms / 1000.0) + 1))
        resource.setrlimit(resource.RLIMIT_CPU, (cpu_seconds, cpu_seconds))

        bytes_limit = max(32, int(memory_mb)) * 1024 * 1024
        resource.setrlimit(resource.RLIMIT_AS, (bytes_limit, bytes_limit))
    except Exception:
        pass

    # Import guardrails: keep ctypes blocked always; allow subprocess-style modules only if automation is granted.
    blocked = {"ctypes"}
    if permissions.get("automation") is not True:
        blocked.update({"subprocess", "multiprocessing"})

    original_import = builtins.__import__

    def guarded_import(name, globals=None, locals=None, fromlist=(), level=0):
        root = name.split(".")[0]
        if root in blocked:
            raise ImportError(f"Module '{root}' is blocked in the Python sandbox")
        return original_import(name, globals, locals, fromlist, level)

    builtins.__import__ = guarded_import

    original_open = builtins.open
    original_os_open = os.open

    def guarded_open(file, mode="r", *args, **kwargs):
        abs_path = os.path.abspath(str(file))
        needs_write = any(flag in mode for flag in ["w", "a", "+", "x"])
        access = "readwrite" if needs_write else "read"
        ensure({"kind": "filesystem", "access": access, "path": abs_path})
        audit(
            "security.filesystem.write" if needs_write else "security.filesystem.read",
            True,
            {"path": abs_path},
        )
        return original_open(file, mode, *args, **kwargs)

    def guarded_os_open(path, flags, mode=0o777, *args, **kwargs):
        abs_path = os.path.abspath(str(path))
        needs_write = (flags & (os.O_WRONLY | os.O_RDWR | os.O_APPEND | os.O_CREAT | os.O_TRUNC)) != 0
        access = "readwrite" if needs_write else "read"
        ensure({"kind": "filesystem", "access": access, "path": abs_path})
        audit(
            "security.filesystem.write" if needs_write else "security.filesystem.read",
            True,
            {"path": abs_path},
        )
        return original_os_open(path, flags, mode, *args, **kwargs)

    builtins.open = guarded_open
    os.open = guarded_os_open

    def guard_read_path(fn, path_arg_index=0):
        original = fn

        def wrapped(*args, **kwargs):
            if len(args) <= path_arg_index:
                return original(*args, **kwargs)
            abs_path = os.path.abspath(str(args[path_arg_index]))
            ensure({"kind": "filesystem", "access": "read", "path": abs_path})
            audit("security.filesystem.read", True, {"path": abs_path, "op": original.__name__})
            return original(*args, **kwargs)

        return wrapped

    def guard_write_path(fn, path_arg_index=0):
        original = fn

        def wrapped(*args, **kwargs):
            if len(args) <= path_arg_index:
                return original(*args, **kwargs)
            abs_path = os.path.abspath(str(args[path_arg_index]))
            ensure({"kind": "filesystem", "access": "readwrite", "path": abs_path})
            audit("security.filesystem.write", True, {"path": abs_path, "op": original.__name__})
            return original(*args, **kwargs)

        return wrapped

    os.listdir = guard_read_path(os.listdir)
    os.scandir = guard_read_path(os.scandir)
    os.stat = guard_read_path(os.stat)
    os.lstat = guard_read_path(os.lstat)
    os.unlink = guard_write_path(os.unlink)
    os.remove = guard_write_path(os.remove)
    os.rmdir = guard_write_path(os.rmdir)
    os.mkdir = guard_write_path(os.mkdir)
    os.makedirs = guard_write_path(os.makedirs)

    original_rename = os.rename
    original_replace = os.replace

    def guarded_rename(src, dst, *args, **kwargs):
        abs_src = os.path.abspath(str(src))
        abs_dst = os.path.abspath(str(dst))
        ensure({"kind": "filesystem", "access": "readwrite", "path": abs_src})
        ensure({"kind": "filesystem", "access": "readwrite", "path": abs_dst})
        audit("security.filesystem.write", True, {"path": abs_src, "op": "rename", "dst": abs_dst})
        return original_rename(src, dst, *args, **kwargs)

    def guarded_replace(src, dst, *args, **kwargs):
        abs_src = os.path.abspath(str(src))
        abs_dst = os.path.abspath(str(dst))
        ensure({"kind": "filesystem", "access": "readwrite", "path": abs_src})
        ensure({"kind": "filesystem", "access": "readwrite", "path": abs_dst})
        audit("security.filesystem.write", True, {"path": abs_src, "op": "replace", "dst": abs_dst})
        return original_replace(src, dst, *args, **kwargs)

    os.rename = guarded_rename
    os.replace = guarded_replace

    # Automation/subprocess escape routes.
    original_system = os.system
    original_popen = os.popen

    def guarded_system(*args, **kwargs):
        ensure({"kind": "automation"})
        audit("security.automation.run", True, {"action": "os.system"})
        return original_system(*args, **kwargs)

    def guarded_popen(*args, **kwargs):
        ensure({"kind": "automation"})
        audit("security.automation.run", True, {"action": "os.popen"})
        return original_popen(*args, **kwargs)

    os.system = guarded_system
    os.popen = guarded_popen

    import socket as _socket

    original_create_connection = _socket.create_connection
    original_connect = _socket.socket.connect
    original_connect_ex = getattr(_socket.socket, "connect_ex", None)

    def _ensure_network_allowed(address):
        if not isinstance(address, tuple) or len(address) < 2:
            ensure({"kind": "network", "url": f"tcp://{address}"})
            return
        host = str(address[0])
        port = int(address[1])
        url = f"tcp://{host}:{port}"
        ensure({"kind": "network", "url": url})
        audit("security.network.request", True, {"url": url, "method": "CONNECT"})

    def guarded_create_connection(address, *args, **kwargs):
        _ensure_network_allowed(address)
        return original_create_connection(address, *args, **kwargs)

    def guarded_connect(self, address):
        _ensure_network_allowed(address)
        return original_connect(self, address)

    def guarded_connect_ex(self, address):
        _ensure_network_allowed(address)
        return original_connect_ex(self, address)  # type: ignore[misc]

    _socket.create_connection = guarded_create_connection
    _socket.socket.connect = guarded_connect
    if original_connect_ex is not None:
        _socket.socket.connect_ex = guarded_connect_ex  # type: ignore[assignment]

    user_stdout = LimitedStringIO(max_output_bytes)
    user_stderr = LimitedStringIO(max_output_bytes)

    ok = False
    result = None
    error_payload = None

    try:
        sandbox_globals = {"__name__": "__main__", "__package__": None}
        with contextlib.redirect_stdout(user_stdout), contextlib.redirect_stderr(user_stderr):
            exec(compile(code, "<formula-python-sandbox>", "exec"), sandbox_globals, sandbox_globals)
        ok = True
        result = sandbox_globals.get("__result__", None)
    except PermissionDenied as e:
        error_payload = {
            "name": "PermissionDeniedError",
            "message": str(e),
            "stack": _format_exc_without_source(),
            "code": "PERMISSION_DENIED",
            "principal": principal,
            "request": e.request,
            "reason": e.reason,
        }
    except OutputLimitExceeded as e:
        error_payload = {
            "name": "SandboxOutputLimitError",
            "message": str(e),
            "stack": _format_exc_without_source(),
            "code": "SANDBOX_OUTPUT_LIMIT",
            "maxBytes": e.max_bytes,
        }
    except MemoryError as e:
        error_payload = {
            "name": "SandboxMemoryLimitError",
            "message": str(e),
            "stack": _format_exc_without_source(),
            "code": "SANDBOX_MEMORY_LIMIT",
            "memoryMb": memory_mb,
        }
    except Exception as e:
        error_payload = {
            "name": e.__class__.__name__,
            "message": str(e),
            "stack": _format_exc_without_source(),
            "code": "PYTHON_SANDBOX_ERROR",
        }

    response = {
        "ok": ok,
        "result": result,
        "stdout": user_stdout.getvalue(),
        "stderr": user_stderr.getvalue(),
        "audit": audit_events,
    }
    if error_payload:
        response["ok"] = False
        response["error"] = error_payload

    sys.stdout.write(json.dumps(response))


if __name__ == "__main__":
    _main()
