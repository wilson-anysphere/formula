from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
) -> None:
    # ------------------------------------------------------------------
    # Logical functions
    # ------------------------------------------------------------------
    bool_inputs = [True, False]
    for a in bool_inputs:
        for b in bool_inputs:
            add_case(
                cases,
                prefix="and",
                tags=["logical", "AND"],
                formula="=AND(A1,B1)",
                inputs=[CellInput("A1", a), CellInput("B1", b)],
            )
            add_case(
                cases,
                prefix="or",
                tags=["logical", "OR"],
                formula="=OR(A1,B1)",
                inputs=[CellInput("A1", a), CellInput("B1", b)],
            )

    # Excel supports `TRUE`/`FALSE` as both logical constants and zero-arg functions.
    add_case(cases, prefix="true", tags=["logical", "TRUE"], formula="=TRUE()")
    add_case(cases, prefix="false", tags=["logical", "FALSE"], formula="=FALSE()")

    for a in [0, 1, -1, "", "0", "1"]:
        add_case(
            cases,
            prefix="not",
            tags=["logical", "NOT"],
            formula="=NOT(A1)",
            inputs=[CellInput("A1", a)],
        )

    if_values = [
        (True, 1, 2),
        (False, 1, 2),
        (0, "yes", "no"),
        (1, "yes", "no"),
        ("", 10, 20),
        ("x", 10, 20),
    ]
    for cond, tv, fv in if_values:
        add_case(
            cases,
            prefix="if",
            tags=["logical", "IF"],
            formula="=IF(A1,B1,C1)",
            inputs=[CellInput("A1", cond), CellInput("B1", tv), CellInput("C1", fv)],
            output_cell="D1",
        )

    add_case(
        cases,
        prefix="iferror",
        tags=["logical", "IFERROR"],
        formula="=IFERROR(A1,42)",
        inputs=[CellInput("A1", formula="=1/0")],
    )
    add_case(
        cases,
        prefix="iferror",
        tags=["logical", "IFERROR"],
        formula="=IFERROR(A1,42)",
        inputs=[CellInput("A1", 1)],
    )
    add_case(
        cases,
        prefix="ifna",
        tags=["logical", "IFNA"],
        formula="=IFNA(A1,42)",
        inputs=[CellInput("A1", formula="=NA()")],
    )
    add_case(
        cases,
        prefix="ifna",
        tags=["logical", "IFNA", "error"],
        formula="=IFNA(A1,42)",
        inputs=[CellInput("A1", formula="=1/0")],
        description="IFNA only catches #N/A (other errors propagate)",
    )
    add_case(
        cases,
        prefix="ifna",
        tags=["logical", "IFNA"],
        formula="=IFNA(A1,42)",
        inputs=[CellInput("A1", 1)],
    )

    add_case(cases, prefix="ifs", tags=["logical", "IFS"], formula="=IFS(FALSE,1,TRUE,2)")
    add_case(
        cases,
        prefix="switch",
        tags=["logical", "SWITCH"],
        formula='=SWITCH(2,1,"one",2,"two","other")',
    )
    add_case(cases, prefix="xor", tags=["logical", "XOR"], formula="=XOR(TRUE,FALSE,TRUE)")
