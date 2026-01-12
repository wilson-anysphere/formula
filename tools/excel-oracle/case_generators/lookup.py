from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
) -> None:
    # ------------------------------------------------------------------
    # Lookup basics: VLOOKUP + INDEX/MATCH
    # ------------------------------------------------------------------
    # Small table in A1:B5, lookup key in D1, result in E1.
    table = [
        (1, "one"),
        (2, "two"),
        (3, "three"),
        (10, "ten"),
        (20, "twenty"),
    ]
    table_inputs = [CellInput("A1", table[0][0]), CellInput("B1", table[0][1])]
    for idx, (k, v) in enumerate(table[1:], start=2):
        table_inputs.append(CellInput(f"A{idx}", k))
        table_inputs.append(CellInput(f"B{idx}", v))

    for key in [1, 2, 3, 10, 20, 4, 0, -1]:
        add_case(
            cases,
            prefix="vlookup",
            tags=["lookup", "VLOOKUP"],
            formula="=VLOOKUP(D1,A1:B5,2,FALSE)",
            inputs=[*table_inputs, CellInput("D1", key)],
            output_cell="E1",
        )
        add_case(
            cases,
            prefix="match",
            tags=["lookup", "MATCH"],
            formula="=MATCH(D1,A1:A5,0)",
            inputs=[*table_inputs, CellInput("D1", key)],
            output_cell="E1",
        )
        add_case(
            cases,
            prefix="index_match",
            tags=["lookup", "INDEX", "MATCH"],
            formula="=INDEX(B1:B5,MATCH(D1,A1:A5,0))",
            inputs=[*table_inputs, CellInput("D1", key)],
            output_cell="E1",
        )

    # HLOOKUP uses a horizontal table (keys in first row).
    h_keys = [1, 2, 3, 10, 20]
    h_vals = ["one", "two", "three", "ten", "twenty"]
    h_inputs = []
    for col, (k, v) in enumerate(zip(h_keys, h_vals), start=1):
        col_letter = chr(ord("A") + col - 1)
        h_inputs.append(CellInput(f"{col_letter}1", k))
        h_inputs.append(CellInput(f"{col_letter}2", v))
    for key in [1, 2, 3, 10, 20, 4]:
        add_case(
            cases,
            prefix="hlookup",
            tags=["lookup", "HLOOKUP"],
            formula="=HLOOKUP(F1,A1:E2,2,FALSE)",
            inputs=[*h_inputs, CellInput("F1", key)],
            output_cell="G1",
        )

    # XLOOKUP / XMATCH (Excel 365+)
    for key in [1, 2, 3, 10, 20, 4]:
        add_case(
            cases,
            prefix="xlookup",
            tags=["lookup", "XLOOKUP"],
            formula='=XLOOKUP(D1,A1:A5,B1:B5,"NF")',
            inputs=[*table_inputs, CellInput("D1", key)],
            output_cell="E1",
        )
        add_case(
            cases,
            prefix="xmatch",
            tags=["lookup", "XMATCH"],
            formula="=XMATCH(D1,A1:A5,0)",
            inputs=[*table_inputs, CellInput("D1", key)],
            output_cell="E1",
        )

    # GETPIVOTDATA requires a real pivot table; without one, Excel returns #REF!.
    add_case(
        cases,
        prefix="getpivotdata",
        tags=["lookup", "GETPIVOTDATA", "error"],
        formula='=GETPIVOTDATA("Sales",A1)',
        inputs=[CellInput("A1", 0)],
        output_cell="C1",
        description="No pivot table present; GETPIVOTDATA should return #REF!",
    )

    # Reference helpers
    add_case(cases, prefix="address", tags=["ref", "ADDRESS"], formula="=ADDRESS(2,3)")
    add_case(cases, prefix="row", tags=["ref", "ROW"], formula="=ROW(A10)")
    add_case(cases, prefix="column", tags=["ref", "COLUMN"], formula="=COLUMN(C5)")
    add_case(cases, prefix="rows", tags=["ref", "ROWS"], formula="=ROWS(A1:B3)")
    add_case(cases, prefix="columns", tags=["ref", "COLUMNS"], formula="=COLUMNS(A1:B3)")

