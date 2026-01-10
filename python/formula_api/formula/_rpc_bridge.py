from __future__ import annotations

import json
from typing import Any, Dict, List, Optional, TextIO


class StdioRpcBridge:
    """
    JSON-RPC-ish bridge over stdio.

    The host (Formula app) is expected to read request messages from stdout and
    write response messages to stdin.

    Stdout is reserved for protocol messages; user script stdout should be
    redirected to stderr by the runner to avoid corrupting the stream.
    """

    def __init__(self, protocol_in: TextIO, protocol_out: TextIO):
        self._in = protocol_in
        self._out = protocol_out
        self._next_id = 1

    def _request(self, method: str, params: Any) -> Any:
        msg_id = self._next_id
        self._next_id += 1

        self._out.write(json.dumps({"type": "rpc", "id": msg_id, "method": method, "params": params}) + "\n")
        self._out.flush()

        while True:
            line = self._in.readline()
            if not line:
                raise RuntimeError(f"RPC connection closed while waiting for {method!r}")
            msg = json.loads(line)
            if msg.get("type") != "rpc_response":
                continue
            if msg.get("id") != msg_id:
                continue
            if msg.get("error"):
                raise RuntimeError(msg["error"])
            return msg.get("result")

    # Workbook/sheet operations
    def get_active_sheet_id(self) -> str:
        return str(self._request("get_active_sheet_id", None))

    def get_sheet_id(self, name: str) -> Optional[str]:
        return self._request("get_sheet_id", {"name": name})

    def create_sheet(self, name: str) -> str:
        return str(self._request("create_sheet", {"name": name}))

    def get_sheet_name(self, sheet_id: str) -> str:
        return str(self._request("get_sheet_name", {"sheet_id": sheet_id}))

    def rename_sheet(self, sheet_id: str, name: str) -> None:
        self._request("rename_sheet", {"sheet_id": sheet_id, "name": name})

    # Range/cell operations
    def get_range_values(self, range_ref: Dict[str, Any]) -> List[List[Any]]:
        return self._request("get_range_values", {"range": range_ref})

    def set_range_values(self, range_ref: Dict[str, Any], values: Any) -> None:
        self._request("set_range_values", {"range": range_ref, "values": values})

    def set_cell_value(self, range_ref: Dict[str, Any], value: Any) -> None:
        self._request("set_cell_value", {"range": range_ref, "value": value})

    def get_cell_formula(self, range_ref: Dict[str, Any]) -> Optional[str]:
        return self._request("get_cell_formula", {"range": range_ref})

    def set_cell_formula(self, range_ref: Dict[str, Any], formula: str) -> None:
        self._request("set_cell_formula", {"range": range_ref, "formula": formula})

    def clear_range(self, range_ref: Dict[str, Any]) -> None:
        self._request("clear_range", {"range": range_ref})

