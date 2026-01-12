from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
) -> None:
    # ------------------------------------------------------------------
    # Database functions (legacy list/database)
    # ------------------------------------------------------------------
    db_inputs = [
        # Database: A1:D4 (header row + 3 records)
        CellInput("A1", "Name"),
        CellInput("B1", "Dept"),
        CellInput("C1", "Age"),
        CellInput("D1", "Salary"),
        CellInput("A2", "Alice"),
        CellInput("B2", "Sales"),
        CellInput("C2", 30),
        CellInput("D2", 1000),
        CellInput("A3", "Bob"),
        CellInput("B3", "Sales"),
        CellInput("C3", 35),
        CellInput("D3", 1500),
        CellInput("A4", "Carol"),
        CellInput("B4", "HR"),
        CellInput("C4", 28),
        CellInput("D4", 1200),
        # Criteria: F1:G3
        # (Dept="Sales" AND Age>30) OR (Dept="HR" AND Age<30)
        CellInput("F1", "Dept"),
        CellInput("G1", "Age"),
        CellInput("F2", "Sales"),
        CellInput("G2", ">30"),
        CellInput("F3", "HR"),
        CellInput("G3", "<30"),
    ]

    for func in [
        "DAVERAGE",
        "DCOUNT",
        "DCOUNTA",
        "DMAX",
        "DMIN",
        "DPRODUCT",
        "DSUM",
        "DSTDEV",
        "DSTDEVP",
        "DVAR",
        "DVARP",
    ]:
        add_case(
            cases,
            prefix=f"database_{func.lower()}",
            tags=["database", func],
            formula=f'={func}(A1:D4,"Salary",F1:G3)',
            inputs=db_inputs,
            output_cell="J1",
        )

    # DGET requires exactly one matching record, so use a single-row criteria range.
    add_case(
        cases,
        prefix="database_dget",
        tags=["database", "DGET"],
        formula='=DGET(A1:D4,"Salary",H1:H2)',
        inputs=[*db_inputs, CellInput("H1", "Name"), CellInput("H2", "Alice")],
        output_cell="J1",
    )

