from __future__ import annotations

from typing import Any, Dict, List, Optional

# When running under Pyodide, the JS runtime registers a `formula_bridge` module
# (via `pyodide.registerJsModule`). Importing it here provides the host callback
# surface for spreadsheet operations.
import formula_bridge  # type: ignore

try:  # pragma: no cover - only available in Pyodide
    from pyodide.ffi import to_py  # type: ignore
except Exception:  # pragma: no cover - native Python

    def to_py(value: Any) -> Any:
        return value


class JsBridge:
    """Bridge implementation backed by a JS module (Pyodide/WebView)."""

    # Workbook/sheet operations
    def get_active_sheet_id(self) -> str:
        return str(formula_bridge.get_active_sheet_id())

    def get_sheet_id(self, name: str) -> Optional[str]:
        return to_py(formula_bridge.get_sheet_id(name))

    def create_sheet(self, name: str) -> str:
        return str(formula_bridge.create_sheet(name))

    def get_sheet_name(self, sheet_id: str) -> str:
        return str(formula_bridge.get_sheet_name(sheet_id))

    def rename_sheet(self, sheet_id: str, name: str) -> None:
        formula_bridge.rename_sheet(sheet_id, name)

    # Selection operations
    def get_selection(self) -> Dict[str, Any]:
        return to_py(formula_bridge.get_selection())

    def set_selection(self, selection: Dict[str, Any]) -> None:
        formula_bridge.set_selection(selection)

    # Range/cell operations
    def get_range_values(self, range_ref: Dict[str, Any]) -> List[List[Any]]:
        return to_py(formula_bridge.get_range_values(range_ref))

    def set_range_values(self, range_ref: Dict[str, Any], values: Any) -> None:
        formula_bridge.set_range_values(range_ref, values)

    def set_cell_value(self, range_ref: Dict[str, Any], value: Any) -> None:
        formula_bridge.set_cell_value(range_ref, value)

    def get_cell_formula(self, range_ref: Dict[str, Any]) -> Optional[str]:
        return to_py(formula_bridge.get_cell_formula(range_ref))

    def set_cell_formula(self, range_ref: Dict[str, Any], formula: str) -> None:
        formula_bridge.set_cell_formula(range_ref, formula)

    def clear_range(self, range_ref: Dict[str, Any]) -> None:
        formula_bridge.clear_range(range_ref)

    # Formatting operations
    def set_range_format(self, range_ref: Dict[str, Any], format_obj: Any) -> None:
        formula_bridge.set_range_format(range_ref, format_obj)

    def get_range_format(self, range_ref: Dict[str, Any]) -> Any:
        return to_py(formula_bridge.get_range_format(range_ref))
