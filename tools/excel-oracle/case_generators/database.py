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
    db_table_inputs = [
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
    ]
    db_standard_criteria_inputs = [
        # Criteria: F1:G3
        # (Dept="Sales" AND Age>30) OR (Dept="HR" AND Age<30)
        CellInput("F1", "Dept"),
        CellInput("G1", "Age"),
        CellInput("F2", "Sales"),
        CellInput("G2", ">30"),
        CellInput("F3", "HR"),
        CellInput("G3", "<30"),
    ]
    db_inputs = [*db_table_inputs, *db_standard_criteria_inputs]

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

    # ------------------------------------------------------------------
    # Database "computed criteria" (blank criteria header + formula criteria)
    # ------------------------------------------------------------------
    # Basic computed criteria: header is blank, criteria cell contains a formula referencing the
    # first record row of the database (row 2), and it is evaluated as-if filled down.
    add_case(
        cases,
        prefix="database_dsum_computed",
        tags=["database", "DSUM", "computed-criteria"],
        formula='=DSUM(A1:D4,"Salary",F1:F2)',
        inputs=[
            *db_table_inputs,
            CellInput("F2", formula="=C2>30"),
        ],
        output_cell="J1",
    )

    # Computed criteria header may be any label that does not match a database field name.
    add_case(
        cases,
        prefix="database_dsum_computed",
        tags=["database", "DSUM", "computed-criteria"],
        formula='=DSUM(A1:D4,"Salary",F1:F2)',
        inputs=[
            *db_table_inputs,
            CellInput("F1", "Criteria"),
            CellInput("F2", formula="=C2>30"),
        ],
        output_cell="J1",
    )

    # Error propagation: a formula error for any record row should propagate out of DSUM.
    add_case(
        cases,
        prefix="database_dsum_computed",
        tags=["database", "DSUM", "computed-criteria", "errors"],
        formula='=DSUM(A1:D4,"Salary",F1:F2)',
        inputs=[
            *db_table_inputs,
            CellInput("F2", formula="=1/(C2-35)>0"),
        ],
        output_cell="J1",
    )

    # Non-formula criteria cells under a blank header should be invalid (Excel returns #VALUE!).
    add_case(
        cases,
        prefix="database_dsum_computed",
        tags=["database", "DSUM", "computed-criteria", "invalid"],
        formula='=DSUM(A1:D4,"Salary",F1:F2)',
        inputs=[
            *db_table_inputs,
            CellInput("F2", ">30"),
        ],
        output_cell="J1",
    )

    # OR across clauses mixing computed + standard criteria:
    # (Dept="Sales" AND computed Age>32) OR (Dept="HR" AND Age<30)
    add_case(
        cases,
        prefix="database_dsum_computed",
        tags=["database", "DSUM", "computed-criteria", "or"],
        formula='=DSUM(A1:D4,"Salary",F1:H3)',
        inputs=[
            *db_table_inputs,
            CellInput("F1", "Dept"),
            CellInput("H1", "Age"),
            CellInput("F2", "Sales"),
            CellInput("G2", formula="=C2>32"),
            CellInput("F3", "HR"),
            CellInput("H3", "<30"),
        ],
        output_cell="J1",
    )
