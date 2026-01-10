from __future__ import annotations

import json
import sys
import traceback
from typing import Any, Dict

import formula
from formula._rpc_bridge import StdioRpcBridge
from formula.runtime.sandbox import apply_cpu_time_limit, apply_memory_limit, apply_sandbox


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

    # Configure the bridge before applying sandbox restrictions so our own
    # imports aren't blocked.
    bridge = StdioRpcBridge(protocol_in=sys.stdin, protocol_out=protocol_out)
    formula.set_bridge(bridge)

    apply_sandbox(cmd.get("permissions", {}))

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
