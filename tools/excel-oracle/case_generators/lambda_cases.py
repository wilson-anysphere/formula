from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
) -> None:
    # ------------------------------------------------------------------
    # LET / LAMBDA + higher-order array functions (Excel 365+)
    # ------------------------------------------------------------------
    # These are intentionally excluded from `compat_gate.py`'s default include-tag list until
    # we have a pinned Excel dataset for them. Keep tags stable so pinned datasets can be
    # regenerated deterministically on Windows + Excel.
    add_case(
        cases,
        prefix="lambda_let",
        tags=["lambda", "LET"],
        formula="=LET(x,1,x+1)",
        description="LET basic binding",
    )
    add_case(
        cases,
        prefix="lambda_let",
        tags=["lambda", "LET"],
        formula="=LET(x,1,y,2,x+y)",
        description="LET multiple bindings",
    )
    add_case(
        cases,
        prefix="lambda_let",
        tags=["lambda", "LET"],
        formula="=LET(x,1,x,2,x)",
        description="LET shadowing (last binding wins)",
    )

    add_case(
        cases,
        prefix="lambda_call",
        tags=["lambda", "LAMBDA"],
        formula="=LAMBDA(x,x+1)(5)",
        description="Postfix lambda invocation (call expression)",
    )
    add_case(
        cases,
        prefix="lambda_call",
        tags=["lambda", "LAMBDA"],
        formula="=LET(inc,LAMBDA(x,x+1),(inc)(5))",
        description="Call a lambda via a LET-bound name (call expression)",
    )

    add_case(
        cases,
        prefix="lambda_map",
        tags=["lambda", "MAP", "spill"],
        formula="=MAP({1,2,3},LAMBDA(x,x*2))",
        description="MAP over a 1x3 array literal",
    )
    add_case(
        cases,
        prefix="lambda_reduce",
        tags=["lambda", "REDUCE"],
        formula="=REDUCE(0,{1,2,3},LAMBDA(acc,x,acc+x))",
        description="REDUCE with explicit initial value",
    )
    add_case(
        cases,
        prefix="lambda_scan",
        tags=["lambda", "SCAN", "spill"],
        formula="=SCAN(0,{1,2,3},LAMBDA(acc,x,acc+x))",
        description="SCAN running sum",
    )
    add_case(
        cases,
        prefix="lambda_byrow",
        tags=["lambda", "BYROW", "spill"],
        formula="=BYROW({1,2;3,4},LAMBDA(r,SUM(r)))",
        description="BYROW with row-wise SUM",
    )
    add_case(
        cases,
        prefix="lambda_bycol",
        tags=["lambda", "BYCOL", "spill"],
        formula="=BYCOL({1,2;3,4},LAMBDA(c,SUM(c)))",
        description="BYCOL with column-wise SUM",
    )
    add_case(
        cases,
        prefix="lambda_makearray",
        tags=["lambda", "MAKEARRAY", "spill"],
        formula="=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))",
        description="MAKEARRAY with index-based values",
    )
