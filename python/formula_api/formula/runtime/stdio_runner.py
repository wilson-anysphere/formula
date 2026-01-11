from __future__ import annotations

import builtins
import importlib.machinery
import importlib._bootstrap
import json
import sys
import traceback
import types
from typing import Any, Dict

import formula
from formula._rpc_bridge import StdioRpcBridge
from formula.runtime.sandbox import apply_cpu_time_limit, apply_memory_limit, apply_sandbox

_ORIGINAL_IMPORT = builtins.__import__


def _normalize_permission(value: Any, allowed: set[str], default: str) -> str:
    if not isinstance(value, str):
        return default
    lowered = value.lower()
    return lowered if lowered in allowed else default


def _filesystem_permission(permissions: Dict[str, Any]) -> str:
    raw = permissions.get("filesystem", permissions.get("fileSystem", "none"))
    return _normalize_permission(raw, {"none", "read", "readwrite"}, "none")


def _network_permission(permissions: Dict[str, Any]) -> str:
    raw = permissions.get("network", "none")
    return _normalize_permission(raw, {"none", "allowlist", "full"}, "none")


def _normalize_hostname(host: str) -> str:
    host = host.strip().lower()
    if host.startswith("[") and host.endswith("]"):
        host = host[1:-1]
    return host.rstrip(".")


def _parse_network_allowlist(raw: Any) -> tuple[set[str], tuple[str, ...]]:
    if not isinstance(raw, (list, tuple, set)):
        return set(), ()

    exact: set[str] = set()
    wildcard_suffixes: set[str] = set()

    for item in raw:
        if not isinstance(item, str):
            continue
        entry = _normalize_hostname(item)
        if not entry:
            continue
        if entry.startswith("*.") and len(entry) > 2:
            wildcard_suffixes.add(entry[2:])
        else:
            exact.add(entry)

    return exact, tuple(sorted(wildcard_suffixes))


def _hostname_in_allowlist(hostname: str, exact: set[str], wildcard_suffixes: tuple[str, ...]) -> bool:
    if hostname in exact:
        return True
    for suffix in wildcard_suffixes:
        if hostname != suffix and hostname.endswith(f".{suffix}"):
            return True
    return False


def _extract_hostname(value: Any) -> str | None:
    host: Any
    if isinstance(value, (tuple, list)):
        if not value:
            return None
        host = value[0]
    else:
        host = value

    if isinstance(host, bytes):
        try:
            host = host.decode("utf-8", errors="ignore")
        except Exception:
            return None

    if isinstance(host, str):
        normalized = _normalize_hostname(host)
        return normalized or None

    return None


def _install_network_audit_hook(network: str, permissions: Dict[str, Any]) -> None:
    """Best-effort network enforcement via sys.addaudithook (native only)."""

    if network == "full" or not hasattr(sys, "addaudithook"):
        return

    raw_allowlist = permissions.get("networkAllowlist", permissions.get("network_allowlist", []))
    exact, wildcard_suffixes = _parse_network_allowlist(raw_allowlist)

    def audit_hook(event: str, args: Any) -> None:
        if not isinstance(event, str) or not event.startswith("socket."):
            return

        if network == "none":
            raise PermissionError("Network access is not permitted")

        # allowlist enforcement
        hostname: str | None = None
        if event in {"socket.connect", "socket.sendto", "socket.sendmsg"}:
            if isinstance(args, tuple) and len(args) >= 2:
                hostname = _extract_hostname(args[1])
        elif event == "socket.getaddrinfo":
            if isinstance(args, tuple) and args:
                hostname = _extract_hostname(args[0])
        elif event in {"socket.gethostbyname", "socket.gethostbyaddr"}:
            if isinstance(args, tuple) and args:
                hostname = _extract_hostname(args[0])
        elif event == "socket.getnameinfo":
            if isinstance(args, tuple) and args:
                hostname = _extract_hostname(args[0])
        else:
            return

        if hostname is None or not _hostname_in_allowlist(hostname, exact, wildcard_suffixes):
            raise PermissionError(f"Network access to {hostname!r} is not permitted")

    try:
        sys.addaudithook(audit_hook)
    except Exception:
        pass


def main() -> None:
    # Stdout is reserved for protocol messages. Redirect user output to stderr so
    # prints don't corrupt the JSON stream.
    protocol_out = sys.stdout
    sys.stdout = sys.stderr  # type: ignore[assignment]
    sys.__stdout__ = sys.stderr  # type: ignore[assignment]

    first_line = sys.stdin.readline()
    if not first_line:
        return

    cmd = json.loads(first_line)
    if cmd.get("type") != "execute":
        protocol_out.write(json.dumps({"type": "result", "success": False, "error": "Invalid command"}) + "\n")
        protocol_out.flush()
        return

    apply_memory_limit(cmd.get("max_memory_bytes"))
    apply_cpu_time_limit(
        max_cpu_seconds=cmd.get("max_cpu_seconds"),
        timeout_ms=cmd.get("timeout_ms", cmd.get("timeoutMs")),
    )

    permissions = cmd.get("permissions", {}) if isinstance(cmd.get("permissions", {}), dict) else {}
    filesystem = _filesystem_permission(permissions)
    network = _network_permission(permissions)

    _install_network_audit_hook(network, permissions)

    # Configure the bridge before applying sandbox restrictions so our own
    # imports aren't blocked.
    bridge = StdioRpcBridge(protocol_in=sys.stdin, protocol_out=protocol_out)
    formula.set_bridge(bridge)

    apply_sandbox(permissions)

    try:
        block_process_execution = not (filesystem == "readwrite" and network == "full")

        blocked_import_roots = set()
        if block_process_execution:
            blocked_import_roots.add("subprocess")
        if network == "none":
            blocked_import_roots.update({"socket", "_socket", "ssl", "_ssl", "http", "urllib", "requests"})

        # This runner is executed as `__main__` (via `python -m`). If user code imports
        # `__main__`, it should see its own module, not the stdio runner internals.
        #
        # Also, removing the runner + sandbox modules from `sys.modules` reduces the
        # chance that a script can bypass restrictions by restoring captured original
        # functions (best-effort; not a hardened security boundary).
        user_main = types.ModuleType("__main__")
        user_main.__dict__.update({"__name__": "__main__", "__file__": "<formula_script>", "__package__": None})
        sys.modules["__main__"] = user_main

        # Restore the builtin import function (avoid exposing the pre-sandbox import
        # callable via `builtins.__import__.__globals__`) and block imports via a
        # meta_path finder instead.
        builtins.__import__ = _ORIGINAL_IMPORT  # type: ignore[assignment]

        class _ImportBlocker:
            def find_spec(self, fullname: str, path=None, target=None):  # type: ignore[no-untyped-def]
                if fullname == "formula.runtime" or fullname.startswith("formula.runtime."):
                    raise PermissionError("Import of 'formula.runtime' is not permitted")

                root = fullname.split(".", 1)[0]
                if root in blocked_import_roots:
                    raise PermissionError(f"Import of {root!r} is not permitted")
                return None

        if blocked_import_roots:
            # Purge already-imported modules so import statements consult meta_path.
            for name in list(sys.modules.keys()):
                root = name.split(".", 1)[0]
                if root in blocked_import_roots:
                    sys.modules.pop(name, None)

            sys.meta_path.insert(0, _ImportBlocker())

            # Best-effort: disallow module reloads while sandboxing is active.
            # Reloading core modules like `socket` can restore unpatched connection
            # methods and bypass the allowlist policy.
            try:
                def blocked_reload(module):  # type: ignore[no-untyped-def]
                    name = getattr(module, "__name__", "<unknown>")
                    raise PermissionError(f"Reload of {name!r} is not permitted")

                importlib.reload = blocked_reload  # type: ignore[assignment]
            except Exception:
                pass

            # Best-effort: block common importlib escape hatches that can load
            # built-in modules directly without consulting sys.meta_path.
            try:
                original_builtin_from_name = importlib._bootstrap._builtin_from_name  # type: ignore[attr-defined]

                def guarded_builtin_from_name(name: str):  # type: ignore[no-untyped-def]
                    root = name.split(".", 1)[0]
                    if root in blocked_import_roots:
                        raise PermissionError(f"Import of {root!r} is not permitted")
                    return original_builtin_from_name(name)

                importlib._bootstrap._builtin_from_name = guarded_builtin_from_name  # type: ignore[attr-defined]
            except Exception:
                pass

            try:
                original_load_module = importlib.machinery.BuiltinImporter.load_module

                def guarded_load_module(name: str):  # type: ignore[no-untyped-def]
                    root = name.split(".", 1)[0]
                    if root in blocked_import_roots:
                        raise PermissionError(f"Import of {root!r} is not permitted")
                    return original_load_module(name)

                importlib.machinery.BuiltinImporter.load_module = guarded_load_module  # type: ignore[assignment]
            except Exception:
                pass

            try:
                original_find_spec = importlib.machinery.BuiltinImporter.find_spec

                def guarded_find_spec(fullname: str, path=None, target=None):  # type: ignore[no-untyped-def]
                    root = fullname.split(".", 1)[0]
                    if root in blocked_import_roots:
                        raise PermissionError(f"Import of {root!r} is not permitted")
                    return original_find_spec(fullname, path, target)

                importlib.machinery.BuiltinImporter.find_spec = guarded_find_spec  # type: ignore[assignment]
            except Exception:
                pass

            try:
                original_create_module = importlib.machinery.BuiltinImporter.create_module

                def guarded_create_module(spec):  # type: ignore[no-untyped-def]
                    name = getattr(spec, "name", "")
                    root = name.split(".", 1)[0] if isinstance(name, str) else ""
                    if root in blocked_import_roots:
                        raise PermissionError(f"Import of {root!r} is not permitted")
                    return original_create_module(spec)

                importlib.machinery.BuiltinImporter.create_module = guarded_create_module  # type: ignore[assignment]
            except Exception:
                pass

            try:
                original_exec_module = importlib.machinery.BuiltinImporter.exec_module

                def guarded_exec_module(module):  # type: ignore[no-untyped-def]
                    name = getattr(module, "__name__", "")
                    root = name.split(".", 1)[0] if isinstance(name, str) else ""
                    if root in blocked_import_roots:
                        raise PermissionError(f"Import of {root!r} is not permitted")
                    return original_exec_module(module)

                importlib.machinery.BuiltinImporter.exec_module = guarded_exec_module  # type: ignore[assignment]
            except Exception:
                pass

        # Drop references to the runner + sandbox modules so scripts can't fetch them
        # directly from sys.modules (e.g. to restore original import/open functions).
        for name in list(sys.modules.keys()):
            if name == "formula.runtime" or name.startswith("formula.runtime."):
                sys.modules.pop(name, None)

        formula_mod = sys.modules.get("formula")
        if formula_mod is not None and hasattr(formula_mod, "runtime"):
            try:
                delattr(formula_mod, "runtime")
            except Exception:
                pass

        globals_dict: Dict[str, Any] = user_main.__dict__
        exec(cmd.get("code", ""), globals_dict, globals_dict)
        protocol_out.write(json.dumps({"type": "result", "success": True}) + "\n")
        protocol_out.flush()
    except Exception as err:
        protocol_out.write(
            json.dumps(
                {
                    "type": "result",
                    "success": False,
                    "error": str(err),
                    "traceback": traceback.format_exc(),
                }
            )
            + "\n"
        )
        protocol_out.flush()


if __name__ == "__main__":
    main()
