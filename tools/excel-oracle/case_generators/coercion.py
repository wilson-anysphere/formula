from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
) -> None:
    # ------------------------------------------------------------------
    # Value coercion / conversion
    # ------------------------------------------------------------------
    # These cases are explicitly chosen to validate the coercion rules we
    # implement (text -> number/date/time), so we can diff against real Excel
    # later. Keep the set small to avoid bloating the corpus.

    # Implicit coercion (text used in arithmetic/logical contexts).
    add_case(cases, prefix="coercion", tags=["coercion", "implicit", "add"], formula='=1+""')
    add_case(cases, prefix="coercion", tags=["coercion", "implicit"], formula='=--""')
    add_case(cases, prefix="coercion", tags=["coercion", "implicit", "NOT"], formula='=NOT("")')
    add_case(cases, prefix="coercion", tags=["coercion", "implicit", "IF"], formula='=IF("",10,20)')
    add_case(cases, prefix="coercion", tags=["coercion", "implicit", "add"], formula='="1234"+1')
    add_case(cases, prefix="coercion", tags=["coercion", "implicit", "add"], formula='="(1000)"+0')
    add_case(cases, prefix="coercion", tags=["coercion", "implicit", "mul"], formula='="10%"*100')
    add_case(cases, prefix="coercion", tags=["coercion", "implicit", "add"], formula='=" 1234 "+0')

    # Explicit conversion functions.
    add_case(cases, prefix="value", tags=["coercion", "VALUE"], formula='=VALUE("2020-01-01")')
    add_case(cases, prefix="value", tags=["coercion", "VALUE"], formula='=VALUE("2020-01-01 13:30")')
    add_case(cases, prefix="timevalue", tags=["coercion", "TIMEVALUE"], formula='=TIMEVALUE("13:00")')
