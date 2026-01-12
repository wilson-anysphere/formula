from __future__ import annotations

from typing import Any


def generate(
    cases: list[dict[str, Any]],
    *,
    add_case,
    CellInput,
) -> None:
    # ------------------------------------------------------------------
    # Date functions (compare on raw serial values; display is locale-dependent)
    # ------------------------------------------------------------------
    date_parts = [
        (1900, 1, 1),
        (1900, 2, 28),
        (1900, 3, 1),
        (1999, 12, 31),
        (2000, 1, 1),
        (2020, 2, 29),
        (2024, 1, 10),
    ]
    for y, m, d in date_parts:
        add_case(
            cases,
            prefix="date",
            tags=["date", "DATE"],
            formula="=DATE(A1,B1,C1)",
            inputs=[CellInput("A1", y), CellInput("B1", m), CellInput("C1", d)],
            output_cell="D1",
        )
        add_case(
            cases,
            prefix="year",
            tags=["date", "YEAR"],
            formula="=YEAR(DATE(A1,B1,C1))",
            inputs=[CellInput("A1", y), CellInput("B1", m), CellInput("C1", d)],
            output_cell="D1",
        )
        add_case(
            cases,
            prefix="month",
            tags=["date", "MONTH"],
            formula="=MONTH(DATE(A1,B1,C1))",
            inputs=[CellInput("A1", y), CellInput("B1", m), CellInput("C1", d)],
            output_cell="D1",
        )
        add_case(
            cases,
            prefix="day",
            tags=["date", "DAY"],
            formula="=DAY(DATE(A1,B1,C1))",
            inputs=[CellInput("A1", y), CellInput("B1", m), CellInput("C1", d)],
            output_cell="D1",
        )

    # Additional date/time functions (keep results numeric to avoid locale-dependent display text).
    add_case(cases, prefix="datevalue", tags=["date", "DATEVALUE", "coercion"], formula='=DATEVALUE("2020-01-01")')
    add_case(cases, prefix="datevalue", tags=["date", "DATEVALUE", "error"], formula='=DATEVALUE("nope")')

    add_case(cases, prefix="time", tags=["date", "TIME"], formula="=TIME(1,30,0)")
    add_case(cases, prefix="time", tags=["date", "TIME"], formula="=TIME(24,0,0)")
    add_case(cases, prefix="time", tags=["date", "TIME", "error"], formula="=TIME(-1,0,0)")

    add_case(cases, prefix="timevalue", tags=["date", "TIMEVALUE"], formula='=TIMEVALUE("1:30")')
    add_case(cases, prefix="timevalue", tags=["date", "TIMEVALUE", "coercion"], formula='=TIMEVALUE("13:30")')
    add_case(cases, prefix="timevalue", tags=["date", "TIMEVALUE", "error"], formula='=TIMEVALUE("nope")')

    add_case(cases, prefix="hour", tags=["date", "HOUR"], formula="=HOUR(TIME(1,2,3))")
    add_case(cases, prefix="minute", tags=["date", "MINUTE"], formula="=MINUTE(TIME(1,2,3))")
    add_case(cases, prefix="second", tags=["date", "SECOND"], formula="=SECOND(TIME(1,2,3))")

    add_case(cases, prefix="edate", tags=["date", "EDATE"], formula="=EDATE(DATE(2020,1,31),1)")
    add_case(cases, prefix="eomonth", tags=["date", "EOMONTH"], formula="=EOMONTH(DATE(2020,1,15),0)")
    add_case(cases, prefix="eomonth", tags=["date", "EOMONTH"], formula="=EOMONTH(DATE(2020,1,15),1)")

    add_case(cases, prefix="weekday", tags=["date", "WEEKDAY"], formula="=WEEKDAY(1)")
    add_case(cases, prefix="weekday", tags=["date", "WEEKDAY"], formula="=WEEKDAY(1,2)")
    add_case(cases, prefix="weekday", tags=["date", "WEEKDAY", "error"], formula="=WEEKDAY(1,0)")

    add_case(cases, prefix="weeknum", tags=["date", "WEEKNUM"], formula="=WEEKNUM(DATE(2020,1,1),1)")
    add_case(cases, prefix="weeknum", tags=["date", "WEEKNUM"], formula="=WEEKNUM(DATE(2020,1,5),2)")
    add_case(cases, prefix="weeknum", tags=["date", "WEEKNUM"], formula="=WEEKNUM(DATE(2021,1,1),21)")
    add_case(cases, prefix="weeknum", tags=["date", "WEEKNUM", "error"], formula="=WEEKNUM(1,9)")

    add_case(cases, prefix="workday", tags=["date", "WORKDAY"], formula="=WORKDAY(DATE(2020,1,1),1)")
    add_case(
        cases,
        prefix="workday",
        tags=["date", "WORKDAY"],
        formula="=WORKDAY(DATE(2020,1,1),1,DATE(2020,1,2))",
    )

    add_case(
        cases,
        prefix="networkdays",
        tags=["date", "NETWORKDAYS"],
        formula="=NETWORKDAYS(DATE(2020,1,1),DATE(2020,1,10))",
    )
    add_case(
        cases,
        prefix="networkdays",
        tags=["date", "NETWORKDAYS"],
        formula="=NETWORKDAYS(DATE(2020,1,1),DATE(2020,1,10),{DATE(2020,1,2),DATE(2020,1,3)})",
    )

    add_case(cases, prefix="workday_intl", tags=["date", "WORKDAY.INTL"], formula="=WORKDAY.INTL(DATE(2020,1,3),1,11)")
    add_case(
        cases,
        prefix="workday_intl",
        tags=["date", "WORKDAY.INTL", "error"],
        formula="=WORKDAY.INTL(DATE(2020,1,3),1,99)",
    )
    add_case(
        cases,
        prefix="workday_intl",
        tags=["date", "WORKDAY.INTL", "error"],
        formula='=WORKDAY.INTL(DATE(2020,1,3),1,"abc")',
    )

    add_case(
        cases,
        prefix="networkdays_intl",
        tags=["date", "NETWORKDAYS.INTL"],
        formula="=NETWORKDAYS.INTL(DATE(2020,1,1),DATE(2020,1,10),11)",
    )
    add_case(
        cases,
        prefix="networkdays_intl",
        tags=["date", "NETWORKDAYS.INTL", "error"],
        formula="=NETWORKDAYS.INTL(DATE(2020,1,1),DATE(2020,1,10),99)",
    )

    add_case(cases, prefix="days", tags=["date", "DAYS"], formula="=DAYS(DATE(2020,1,10),DATE(2020,1,1))")
    add_case(cases, prefix="days360", tags=["date", "DAYS360"], formula="=DAYS360(DATE(2020,1,1),DATE(2020,2,1))")
    add_case(cases, prefix="datedif", tags=["date", "DATEDIF"], formula='=DATEDIF(DATE(2020,1,1),DATE(2021,2,1),"y")')
    add_case(cases, prefix="yearfrac", tags=["date", "YEARFRAC"], formula="=YEARFRAC(DATE(2020,1,1),DATE(2021,1,1))")
    # Integer-coercion variants for the `basis` argument.
    add_case(
        cases,
        prefix="yearfrac",
        tags=["date", "YEARFRAC", "coercion"],
        formula="=YEARFRAC(DATE(2020,1,1),DATE(2021,1,1),0.9)",
        description="basis=0.9",
    )
    add_case(
        cases,
        prefix="yearfrac",
        tags=["date", "YEARFRAC", "coercion"],
        formula="=YEARFRAC(DATE(2020,1,1),DATE(2021,1,1),1.9)",
        description="basis=1.9",
    )
    add_case(
        cases,
        prefix="yearfrac",
        tags=["date", "YEARFRAC", "coercion"],
        formula="=YEARFRAC(DATE(2020,1,1),DATE(2021,1,1),-0.1)",
        description="basis=-0.1",
    )
    add_case(cases, prefix="iso_weeknum", tags=["date", "ISO.WEEKNUM"], formula="=ISO.WEEKNUM(DATE(2021,1,1))")
    add_case(cases, prefix="isoweeknum", tags=["date", "ISOWEEKNUM"], formula="=ISOWEEKNUM(DATE(2021,1,1))")
