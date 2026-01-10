from __future__ import annotations

import builtins
import json
import sys
import traceback
from typing import Any, Dict

import formula
from formula._rpc_bridge import StdioRpcBridge


def _apply_memory_limit(max_memory_bytes: int | None) -> None:
    if not max_memory_bytes:
        return

    try:
        import resource  # Unix only

        # RLIMIT_AS limits the total available address space.
        resource.setrlimit(resource.RLIMIT_AS, (max_memory_bytes, max_memory_bytes))
    except Exception:
        # Best-effort: not all platforms support resource or RLIMIT_AS.
        pass


def _apply_sandbox(permissions: Dict[str, Any]) -> None:
    filesystem = permissions.get("filesystem", "none")
    network = permissions.get("network", "none")

    if filesystem == "none":

        def blocked_open(*_args: Any, **_kwargs: Any) -> Any:
            raise PermissionError("Filesystem access is not permitted")

        builtins.open = blocked_open  # type: ignore[assignment]

    # Always block interactive input (it competes for stdin with the RPC protocol).
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


def main() -> None:
    # Stdout is reserved for protocol messages. Redirect user output to stderr so
    # prints don't corrupt the JSON stream.
    protocol_out = sys.stdout
    sys.stdout = sys.stderr  # type: ignore[assignment]

    first_line = sys.stdin.readline()
    if not first_line:
        return

    cmd = json.loads(first_line)
    if cmd.get("type") != "execute":
        protocol_out.write(json.dumps({"type": "result", "success": False, "error": "Invalid command"}) + "\n")
        protocol_out.flush()
        return

    _apply_memory_limit(cmd.get("max_memory_bytes"))

    # Configure the bridge before applying sandbox restrictions so our own
    # imports aren't blocked.
    bridge = StdioRpcBridge(protocol_in=sys.stdin, protocol_out=protocol_out)
    formula.set_bridge(bridge)

    _apply_sandbox(cmd.get("permissions", {}))

    try:
        globals_dict: Dict[str, Any] = {"__name__": "__main__"}
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
