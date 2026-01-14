from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
    include_volatile: bool = False,
) -> None:
    # ------------------------------------------------------------------
    # Information functions
    # ------------------------------------------------------------------
    add_case(cases, prefix="isblank", tags=["info", "ISBLANK"], formula="=ISBLANK(A1)")
    add_case(
        cases,
        prefix="isblank",
        tags=["info", "ISBLANK"],
        formula="=ISBLANK(A1)",
        inputs=[CellInput("A1", "")],
        description="Empty string is not considered blank",
    )

    add_case(cases, prefix="isnumber", tags=["info", "ISNUMBER"], formula="=ISNUMBER(A1)", inputs=[CellInput("A1", 1)])
    add_case(cases, prefix="isnumber", tags=["info", "ISNUMBER"], formula="=ISNUMBER(A1)", inputs=[CellInput("A1", "1")])
    add_case(cases, prefix="istext", tags=["info", "ISTEXT"], formula="=ISTEXT(A1)", inputs=[CellInput("A1", "x")])
    add_case(cases, prefix="istext", tags=["info", "ISTEXT"], formula="=ISTEXT(A1)", inputs=[CellInput("A1", 1)])
    add_case(cases, prefix="isref", tags=["info", "ISREF"], formula="=ISREF(A1)")
    add_case(cases, prefix="isnontext", tags=["info", "ISNONTEXT"], formula='=ISNONTEXT("x")')
    add_case(cases, prefix="islogical", tags=["info", "ISLOGICAL"], formula="=ISLOGICAL(A1)", inputs=[CellInput("A1", True)])
    add_case(cases, prefix="islogical", tags=["info", "ISLOGICAL"], formula="=ISLOGICAL(A1)", inputs=[CellInput("A1", 0)])

    add_case(cases, prefix="iserror", tags=["info", "ISERROR"], formula="=ISERROR(A1)", inputs=[CellInput("A1", formula="=1/0")])
    add_case(cases, prefix="iserror", tags=["info", "ISERROR"], formula="=ISERROR(A1)", inputs=[CellInput("A1", 1)])
    add_case(cases, prefix="iserr", tags=["info", "ISERR"], formula="=ISERR(A1)", inputs=[CellInput("A1", formula="=1/0")])
    add_case(cases, prefix="iserr", tags=["info", "ISERR"], formula="=ISERR(A1)", inputs=[CellInput("A1", formula="=NA()")])
    add_case(cases, prefix="isna", tags=["info", "ISNA"], formula="=ISNA(A1)", inputs=[CellInput("A1", formula="=NA()")])
    add_case(cases, prefix="isna", tags=["info", "ISNA"], formula="=ISNA(A1)", inputs=[CellInput("A1", formula="=1/0")])

    add_case(cases, prefix="errtype", tags=["info", "ERROR.TYPE"], formula="=ERROR.TYPE(1/0)")
    add_case(cases, prefix="errtype", tags=["info", "ERROR.TYPE"], formula="=ERROR.TYPE(NA())")
    add_case(cases, prefix="errtype", tags=["info", "ERROR.TYPE", "error"], formula="=ERROR.TYPE(1)")

    add_case(cases, prefix="type", tags=["info", "TYPE"], formula="=TYPE(1)")
    add_case(cases, prefix="type", tags=["info", "TYPE"], formula='=TYPE("x")')
    add_case(cases, prefix="type", tags=["info", "TYPE"], formula="=TYPE(TRUE)")
    add_case(cases, prefix="type", tags=["info", "TYPE"], formula="=TYPE(NA())")

    add_case(cases, prefix="n", tags=["info", "N"], formula="=N(TRUE)")
    add_case(cases, prefix="n", tags=["info", "N"], formula='=N("x")')
    add_case(cases, prefix="n", tags=["info", "N"], formula="=N(NA())")

    add_case(cases, prefix="t", tags=["info", "T"], formula='=T("x")')
    add_case(cases, prefix="t", tags=["info", "T"], formula="=T(1)")
    add_case(cases, prefix="t", tags=["info", "T"], formula="=T(NA())")

    # NOTE: INFO/CELL are volatile (they depend on workbook/environment state) and are intentionally
    # excluded from the oracle corpus so results can be pinned/stably compared.
    if include_volatile:
        # INFO / CELL (worksheet introspection)
        add_case(cases, prefix="cell", tags=["info", "CELL"], formula='=CELL("address",A1)')
        add_case(cases, prefix="cell", tags=["info", "CELL"], formula='=CELL("row",A10)')
        add_case(cases, prefix="cell", tags=["info", "CELL"], formula='=CELL("col",C1)')
        add_case(cases, prefix="cell", tags=["info", "CELL"], formula='=CELL("type",A1)')
        add_case(
            cases,
            prefix="cell",
            tags=["info", "CELL"],
            formula='=CELL("contents",A1)',
            inputs=[CellInput("A1", 5)],
            description='CELL("contents") returns the value for constant cells',
        )

        add_case(cases, prefix="info", tags=["info", "INFO"], formula='=INFO("recalc")')
        add_case(cases, prefix="info", tags=["info", "INFO", "error"], formula='=INFO("no_such_key")')

    # Workbook / worksheet metadata functions (auditing helpers in modern Excel).
    add_case(cases, prefix="sheet", tags=["info", "SHEET"], formula="=SHEET()")
    add_case(cases, prefix="sheet", tags=["info", "SHEET"], formula="=SHEET(A1)")
    # Avoid `SHEETS()` (no arg) because the default sheet count in new workbooks can vary by
    # Excel user preferences; `SHEETS(A1)` is deterministic.
    add_case(cases, prefix="sheets", tags=["info", "SHEETS"], formula="=SHEETS(A1)")
    add_case(
        cases,
        prefix="formulatext",
        tags=["info", "FORMULATEXT"],
        formula="=FORMULATEXT(A1)",
        inputs=[CellInput("A1", formula="=1+1")],
    )
    add_case(
        cases,
        prefix="formulatext",
        tags=["info", "FORMULATEXT", "error"],
        formula="=FORMULATEXT(A1)",
        inputs=[CellInput("A1", 5)],
        description="FORMULATEXT returns #N/A when the referenced cell has no formula",
    )
    add_case(
        cases,
        prefix="isformula",
        tags=["info", "ISFORMULA"],
        formula="=ISFORMULA(A1)",
        inputs=[CellInput("A1", formula="=1+1")],
    )
    add_case(
        cases,
        prefix="isformula",
        tags=["info", "ISFORMULA"],
        formula="=ISFORMULA(A1)",
        inputs=[CellInput("A1", 5)],
        description="ISFORMULA returns FALSE for constant values",
    )
