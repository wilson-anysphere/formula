from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
) -> None:
    # ------------------------------------------------------------------
    # Arithmetic operators
    # ------------------------------------------------------------------
    arith_vals = [0, 1, -1, 2, -2, 0.5, 10, -10]
    arith_ops = [
        ("+", "add"),
        ("-", "sub"),
        ("*", "mul"),
        ("/", "div"),
        ("^", "pow"),
    ]
    for sym, name in arith_ops:
        for a in arith_vals:
            for b in arith_vals:
                add_case(
                    cases,
                    prefix=f"arith_{name}",
                    tags=["arith", name],
                    formula=f"=A1{sym}B1",
                    output_cell="C1",
                    inputs=[CellInput("A1", a), CellInput("B1", b)],
                )

    # ------------------------------------------------------------------
    # Comparison operators
    # ------------------------------------------------------------------
    cmp_vals = [0, 1, -1, 2]
    cmp_ops = [
        ("=", "eq"),
        ("<>", "ne"),
        ("<", "lt"),
        ("<=", "le"),
        (">", "gt"),
        (">=", "ge"),
    ]
    for sym, name in cmp_ops:
        for a in cmp_vals:
            for b in cmp_vals:
                add_case(
                    cases,
                    prefix=f"cmp_{name}",
                    tags=["cmp", name],
                    formula=f"=A1{sym}B1",
                    output_cell="C1",
                    inputs=[CellInput("A1", a), CellInput("B1", b)],
                )

