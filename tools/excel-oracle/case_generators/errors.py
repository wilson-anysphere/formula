from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
) -> None:
    # ------------------------------------------------------------------
    # Explicit error values
    # ------------------------------------------------------------------
    add_case(cases, prefix="err_div0", tags=["error"], formula="=1/0", inputs=[], output_cell="A1")
    add_case(cases, prefix="err_na", tags=["error"], formula="=NA()", inputs=[], output_cell="A1")
    add_case(cases, prefix="err_name", tags=["error"], formula="=NO_SUCH_FUNCTION(1)", inputs=[], output_cell="A1")
