"""
`formula` – Python scripting API for Formula spreadsheets.

This module is designed to run in two environments:

1. Pyodide (browser/webview): spreadsheet operations are bridged via a JS module
   injected into the Pyodide runtime.
2. Native Python (desktop): spreadsheet operations are bridged via a JSON-RPC
   transport (stdio/IPC).

The public API intentionally mirrors the conceptual model from
docs/08-macro-compatibility.md:

- `active_sheet` (dynamic attribute)
- `get_sheet(name)`
- `create_sheet(name)`
- Sheet and Range objects with convenient cell/range accessors
- Optional pandas helpers (`to_dataframe`, `from_dataframe`)
- `@custom_function` decorator for UDF registration
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple, Union

from ._a1 import parse_a1
from ._bridge import Bridge

_bridge: Optional[Bridge] = None


def set_bridge(bridge: Bridge) -> None:
    """Configure the runtime bridge used to talk to the host spreadsheet."""

    global _bridge
    _bridge = bridge


def _require_bridge() -> Bridge:
    if _bridge is None:
        raise RuntimeError(
            "formula bridge is not configured. "
            "If you're running natively, use formula.runtime.stdio_runner. "
            "If you're running in Pyodide, inject a `formula_bridge` module and "
            "call formula.set_bridge(...) from your runtime."
        )
    return _bridge


def __getattr__(name: str) -> Any:
    # PEP 562: module-level dynamic attributes.
    if name == "active_sheet":
        bridge = _require_bridge()
        sheet_id = bridge.get_active_sheet_id()
        return Sheet(sheet_id=sheet_id, bridge=bridge)
    raise AttributeError(name)


def get_sheet(name: str) -> "Sheet":
    bridge = _require_bridge()
    sheet_id = bridge.get_sheet_id(name)
    if sheet_id is None:
        raise KeyError(f"Sheet not found: {name!r}")
    return Sheet(sheet_id=sheet_id, bridge=bridge)


def create_sheet(name: str) -> "Sheet":
    bridge = _require_bridge()
    sheet_id = bridge.create_sheet(name)
    return Sheet(sheet_id=sheet_id, bridge=bridge)


def get_selection() -> Dict[str, Any]:
    """
    Get the current selection as a range reference dict:
    {sheet_id, start_row, start_col, end_row, end_col}.
    """

    bridge = _require_bridge()
    return bridge.get_selection()


def set_selection(selection: Dict[str, Any]) -> None:
    """Set the current selection using a range reference dict."""

    bridge = _require_bridge()
    bridge.set_selection(selection)


def _pandas() -> Any:
    try:
        import pandas as pd  # type: ignore

        return pd
    except Exception:  # pragma: no cover - environment dependent
        return None


@dataclass(frozen=True)
class RangeRef:
    sheet_id: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int

    @property
    def is_single_cell(self) -> bool:
        return self.start_row == self.end_row and self.start_col == self.end_col


class Sheet:
    """Represents a worksheet in the spreadsheet."""

    def __init__(self, sheet_id: str, bridge: Bridge):
        self._id = sheet_id
        self._bridge = bridge

    @property
    def id(self) -> str:
        return self._id

    @property
    def name(self) -> str:
        return self._bridge.get_sheet_name(self._id)

    @name.setter
    def name(self, value: str) -> None:
        self._bridge.rename_sheet(self._id, value)

    def __getitem__(self, key: str) -> "Range":
        # Access cells via sheet["A1"] or sheet["A1:B10"].
        start_row, start_col, end_row, end_col = parse_a1(key)
        return Range(
            ref=RangeRef(
                sheet_id=self._id,
                start_row=start_row,
                start_col=start_col,
                end_row=end_row,
                end_col=end_col,
            ),
            bridge=self._bridge,
        )

    def __setitem__(self, key: str, value: Any) -> None:
        self[key].value = value


class Range:
    """Represents a range of cells."""

    def __init__(self, ref: RangeRef, bridge: Bridge):
        self._ref = ref
        self._bridge = bridge

    @property
    def ref(self) -> RangeRef:
        return self._ref

    @property
    def value(self) -> Any:
        values = self._bridge.get_range_values(self._ref.__dict__)
        if self._ref.is_single_cell:
            return values[0][0] if values and values[0] else None
        return values

    @value.setter
    def value(self, val: Any) -> None:
        pd = _pandas()
        if pd is not None and isinstance(val, pd.DataFrame):
            self.from_dataframe(val)
            return

        if self._ref.is_single_cell and isinstance(val, str):
            # Match DocumentController's string input semantics:
            # - Leading apostrophe escapes literal text.
            # - Leading whitespace is ignored when detecting formulas ("=..."), but
            #   a bare "=" is treated as a literal value.
            if val.startswith("'"):
                self._bridge.set_cell_value(self._ref.__dict__, val[1:])
                return

            trimmed = val.lstrip()
            if trimmed.startswith("=") and len(trimmed) > 1:
                self.formula = trimmed
                return

        if self._ref.is_single_cell and not isinstance(val, (list, tuple)):
            self._bridge.set_cell_value(self._ref.__dict__, val)
            return

        # Convenience: allow 1D sequences for single-row or single-column ranges.
        if (
            not self._ref.is_single_cell
            and isinstance(val, (list, tuple))
            and (len(val) == 0 or not isinstance(val[0], (list, tuple)))
        ):
            row_count = self._ref.end_row - self._ref.start_row + 1
            col_count = self._ref.end_col - self._ref.start_col + 1
            if row_count == 1:
                if len(val) != col_count:
                    raise ValueError(
                        f"Range expects {col_count} values for a 1x{col_count} row range, got {len(val)}"
                    )
                val = [list(val)]
            elif col_count == 1:
                if len(val) != row_count:
                    raise ValueError(
                        f"Range expects {row_count} values for a {row_count}x1 column range, got {len(val)}"
                    )
                val = [[v] for v in val]
            else:
                raise TypeError(
                    "Range.value expects a 2D list for multi-cell ranges "
                    "(a 1D list is only allowed for single-row or single-column ranges)"
                )

        self._bridge.set_range_values(self._ref.__dict__, val)

    @property
    def formula(self) -> Optional[str]:
        if not self._ref.is_single_cell:
            raise ValueError("Range.formula is only available for a single cell range")
        return self._bridge.get_cell_formula(self._ref.__dict__)

    @formula.setter
    def formula(self, val: Optional[str]) -> None:
        if not self._ref.is_single_cell:
            raise ValueError("Range.formula is only available for a single cell range")
        if val is None:
            # Clear both value and formula (preserving formatting), matching
            # DocumentController.setCellFormula(..., null).
            self._bridge.set_cell_value(self._ref.__dict__, None)
            return

        trimmed = str(val).lstrip()
        if trimmed == "":
            self._bridge.set_cell_value(self._ref.__dict__, None)
            return

        normalized = trimmed if trimmed.startswith("=") else f"={trimmed}"
        self._bridge.set_cell_formula(self._ref.__dict__, normalized)

    @property
    def format(self) -> Any:
        return self._bridge.get_range_format(self._ref.__dict__)

    @format.setter
    def format(self, val: Any) -> None:
        self._bridge.set_range_format(self._ref.__dict__, val)

    def set_format(self, val: Any) -> None:
        self.format = val

    def clear(self) -> None:
        self._bridge.clear_range(self._ref.__dict__)

    def to_dataframe(self, header: bool = True) -> Any:
        pd = _pandas()
        if pd is None:  # pragma: no cover - environment dependent
            raise ImportError("pandas is not available in this Python runtime")

        values = self.value
        if not isinstance(values, list):
            return pd.DataFrame([[values]])

        if header and values:
            return pd.DataFrame(values[1:], columns=values[0])
        return pd.DataFrame(values)

    def from_dataframe(self, df: Any, include_header: bool = True) -> None:
        pd = _pandas()
        if pd is None:  # pragma: no cover - environment dependent
            raise ImportError("pandas is not available in this Python runtime")
        if not isinstance(df, pd.DataFrame):  # type: ignore[attr-defined]
            raise TypeError("from_dataframe expects a pandas.DataFrame")

        values: List[List[Any]] = []
        if include_header:
            values.append([str(col) for col in df.columns.tolist()])
        values.extend(df.values.tolist())

        self._bridge.set_range_values(self._ref.__dict__, values)


_function_registry: Dict[str, Callable[..., Any]] = {}


def custom_function(func: Optional[Callable[..., Any]] = None, *, name: Optional[str] = None) -> Any:
    """Decorator to register a Python function as a spreadsheet custom function."""

    def register(fn: Callable[..., Any]) -> Callable[..., Any]:
        _function_registry[name or fn.__name__] = fn
        return fn

    return register(func) if func is not None else register


def list_custom_functions() -> List[str]:
    return sorted(_function_registry.keys())
