import builtins
import json
import os
import sys
import traceback
import types
import urllib.parse


def _is_within_scope(target_path: str, scope_path: str) -> bool:
    abs_target = os.path.abspath(target_path)
    abs_scope = os.path.abspath(scope_path)
    try:
        common = os.path.commonpath([abs_scope, abs_target])
    except ValueError:
        return False
    return common == abs_scope


def _network_host_allowed(host: str, allowlist_entries: list[str]) -> bool:
    host = host.lower()
    for entry in allowlist_entries:
        entry = str(entry).strip()
        if not entry:
            continue

        # Origin style entry: "https://api.example.com" (ignore scheme for socket-level checks)
        if "://" in entry:
            parsed = urllib.parse.urlparse(entry)
            if not parsed.hostname:
                continue
            entry_host = parsed.hostname.lower()
            if host == entry_host:
                return True
            continue

        # Wildcard host entry: "*.example.com"
        if entry.startswith("*."):
            suffix = entry[2:].lower()
            if host == suffix:
                return True
            if host.endswith("." + suffix):
                return True
            continue

        # Exact host entry: "example.com"
        if host == entry.lower():
            return True

    return False


def _install_filesystem_guards(permissions: dict):
    fs_perm = permissions.get("filesystem", {}) if permissions else {}
    read_scopes = [os.path.abspath(p) for p in fs_perm.get("read", [])] + [
        os.path.abspath(p) for p in fs_perm.get("readwrite", [])
    ]
    write_scopes = [os.path.abspath(p) for p in fs_perm.get("readwrite", [])]

    original_open = builtins.open

    def guarded_open(file, mode="r", *args, **kwargs):
        # Determine access type based on mode
        needs_write = any(flag in mode for flag in ["w", "a", "+", "x"])
        abs_path = os.path.abspath(str(file))

        if needs_write:
            allowed = any(_is_within_scope(abs_path, scope) for scope in write_scopes)
        else:
            allowed = any(_is_within_scope(abs_path, scope) for scope in read_scopes)

        if not allowed:
            raise PermissionError(f"Filesystem access denied for {abs_path} (mode={mode})")

        return original_open(file, mode, *args, **kwargs)

    builtins.open = guarded_open

    # Basic protection against shell escape routes.
    import os as _os

    def _blocked(*_args, **_kwargs):
        raise PermissionError("Automation/subprocess access is blocked in the Python sandbox")

    _os.system = _blocked
    _os.popen = _blocked


def _install_network_guards(permissions: dict):
    net_perm = permissions.get("network", {}) if permissions else {}
    mode = net_perm.get("mode", "none")
    allowlist = net_perm.get("allowlist", [])

    import socket as _socket

    original_create_connection = _socket.create_connection
    original_connect = _socket.socket.connect

    def _assert_allowed(host: str):
        if mode == "full":
            return
        if mode == "allowlist" and _network_host_allowed(host, allowlist):
            return
        raise PermissionError(f"Network access denied for host {host}")

    def guarded_create_connection(address, *args, **kwargs):
        host, _port = address
        _assert_allowed(str(host))
        return original_create_connection(address, *args, **kwargs)

    def guarded_connect(self, address):
        host, _port = address
        _assert_allowed(str(host))
        return original_connect(self, address)

    _socket.create_connection = guarded_create_connection
    _socket.socket.connect = guarded_connect


def _install_import_guards():
    blocked = {"subprocess", "ctypes", "multiprocessing"}
    original_import = builtins.__import__

    def guarded_import(name, globals=None, locals=None, fromlist=(), level=0):
        root = name.split(".")[0]
        if root in blocked:
            raise ImportError(f"Module '{root}' is blocked in the Python sandbox")
        return original_import(name, globals, locals, fromlist, level)

    builtins.__import__ = guarded_import


def _install_resource_limits(timeout_ms: int, memory_mb: int):
    # These limits are best-effort and primarily target Linux/macOS.
    try:
        import resource

        # CPU time in seconds. Add a small cushion for interpreter startup.
        cpu_seconds = max(1, int((timeout_ms / 1000.0) + 1))
        resource.setrlimit(resource.RLIMIT_CPU, (cpu_seconds, cpu_seconds))

        # Address space limit.
        bytes_limit = max(32, int(memory_mb)) * 1024 * 1024
        resource.setrlimit(resource.RLIMIT_AS, (bytes_limit, bytes_limit))
    except Exception:
        # If resource limits are unavailable, continue without hard enforcement.
        pass


def _main():
    try:
        payload = json.loads(sys.stdin.read() or "{}")
        permissions = payload.get("permissions", {})
        timeout_ms = int(payload.get("timeoutMs", 5000))
        memory_mb = int(payload.get("memoryMb", 128))
        code = payload.get("code", "")

        _install_resource_limits(timeout_ms=timeout_ms, memory_mb=memory_mb)
        _install_import_guards()
        _install_filesystem_guards(permissions)
        _install_network_guards(permissions)

        # Execute user code. We intentionally do not expose host objects.
        sandbox_globals = {"__name__": "__main__", "__package__": None}
        exec(compile(code, "<formula-python-sandbox>", "exec"), sandbox_globals, sandbox_globals)

        sys.stdout.write(json.dumps({"ok": True, "result": None}))
    except Exception as e:
        sys.stdout.write(
            json.dumps(
                {
                    "ok": False,
                    "error": {
                        "name": e.__class__.__name__,
                        "message": str(e),
                        "stack": traceback.format_exc(),
                        "code": "PYTHON_SANDBOX_ERROR",
                    },
                }
            )
        )


if __name__ == "__main__":
    _main()

