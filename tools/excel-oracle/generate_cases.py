#!/usr/bin/env python3
"""
Deterministically generate a curated (~1k) Excel formula corpus.

This corpus is intentionally "small enough to run in real Excel in CI",
but broad enough to cover many P0/P1 function behaviors and edge cases.

The output is committed at:
  tests/compatibility/excel-oracle/cases.json
"""

from __future__ import annotations

import argparse
import dataclasses
import datetime as dt
import hashlib
import itertools
import json
from typing import Any, Iterable


@dataclasses.dataclass(frozen=True)
class CellInput:
    cell: str
    value: Any | None = None
    formula: str | None = None

    def to_json(self) -> dict[str, Any]:
        payload: dict[str, Any] = {"cell": self.cell}
        if self.formula is not None:
            payload["formula"] = self.formula
        else:
            payload["value"] = self.value
        return payload


def _stable_case_id(case: dict[str, Any], prefix: str) -> str:
    canonical = json.dumps(case, sort_keys=True, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    digest = hashlib.sha1(canonical).hexdigest()[:12]
    return f"{prefix}_{digest}"


def _add_case(
    cases: list[dict[str, Any]],
    *,
    prefix: str,
    tags: list[str],
    formula: str,
    inputs: Iterable[CellInput] = (),
    output_cell: str = "C1",
    description: str | None = None,
) -> None:
    payload: dict[str, Any] = {
        "formula": formula,
        "outputCell": output_cell,
        "inputs": [i.to_json() for i in inputs],
        "tags": tags,
    }
    if description:
        payload["description"] = description

    payload["id"] = _stable_case_id(payload, prefix=prefix)
    cases.append(payload)


def _excel_serial_1900(year: int, month: int, day: int) -> int:
    """
    Excel 1900 date system serial with Lotus leap-year bug enabled.

    - Serial 1 == 1900-01-01
    - Serial 60 == fictitious 1900-02-29
    - For dates >= 1900-03-01, Excel serials are offset by +1 vs real day count.
    """

    base = dt.date(1899, 12, 31)
    cur = dt.date(year, month, day)
    serial = (cur - base).days
    if cur >= dt.date(1900, 3, 1):
        serial += 1
    return serial


def generate_cases() -> dict[str, Any]:
    cases: list[dict[str, Any]] = []

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
                _add_case(
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
                _add_case(
                    cases,
                    prefix=f"cmp_{name}",
                    tags=["cmp", name],
                    formula=f"=A1{sym}B1",
                    output_cell="C1",
                    inputs=[CellInput("A1", a), CellInput("B1", b)],
                )

    # ------------------------------------------------------------------
    # Basic math functions
    # ------------------------------------------------------------------
    math_vals = [0, 1, -1, 2, -2, 0.5, -0.5, 10.25, -10.25, 1e-9, -1e-9, 1e9, -1e9]

    for v in math_vals:
        _add_case(cases, prefix="abs", tags=["math", "ABS"], formula="=ABS(A1)", inputs=[CellInput("A1", v)])
        _add_case(cases, prefix="sign", tags=["math", "SIGN"], formula="=SIGN(A1)", inputs=[CellInput("A1", v)])
        _add_case(cases, prefix="int", tags=["math", "INT"], formula="=INT(A1)", inputs=[CellInput("A1", v)])

    mod_divisors = [1, 2, 3, -2, -3, 10]
    for a in [0, 1, -1, 2, -2, 10, -10, 10.5, -10.5]:
        for b in mod_divisors:
            _add_case(
                cases,
                prefix="mod",
                tags=["math", "MOD"],
                formula="=MOD(A1,B1)",
                inputs=[CellInput("A1", a), CellInput("B1", b)],
            )

    # Keep ROUND* combinations capped so the corpus stays ~1k cases total.
    round_vals = [0, 0.1, 0.5, 1.5, 2.5, 10.25, -10.25, 1234.5678, -1234.5678, 1e-7]
    round_digits = [-2, -1, 0, 1, 2]
    for func in ["ROUND", "ROUNDUP", "ROUNDDOWN"]:
        for v in round_vals:
            for d in round_digits:
                _add_case(
                    cases,
                    prefix=func.lower(),
                    tags=["math", func],
                    formula=f"={func}(A1,B1)",
                    inputs=[CellInput("A1", v), CellInput("B1", d)],
                )

    # ------------------------------------------------------------------
    # Aggregates over ranges
    # ------------------------------------------------------------------
    agg_vals = [0, 1, -1, 2, -2, 0.5]
    # Keep range sizes small to make Excel automation fast, and cap the number of
    # permutations so the overall corpus stays ~1k.
    for a, b, c in itertools.islice(itertools.product(agg_vals, repeat=3), 30):
        inputs = [
            CellInput("A1", a),
            CellInput("A2", b),
            CellInput("A3", c),
        ]
        _add_case(cases, prefix="sum_r", tags=["agg", "SUM"], formula="=SUM(A1:A3)", inputs=inputs)
        _add_case(cases, prefix="avg_r", tags=["agg", "AVERAGE"], formula="=AVERAGE(A1:A3)", inputs=inputs)
        _add_case(cases, prefix="min_r", tags=["agg", "MIN"], formula="=MIN(A1:A3)", inputs=inputs)
        _add_case(cases, prefix="max_r", tags=["agg", "MAX"], formula="=MAX(A1:A3)", inputs=inputs)

    # Type coercion edges (SUM ignores text in ranges, but direct args coerce differently).
    _add_case(
        cases,
        prefix="sum_args",
        tags=["agg", "SUM", "coercion"],
        formula='=SUM("5",3)',
        inputs=[],
    )
    _add_case(
        cases,
        prefix="sum_args",
        tags=["agg", "SUM", "coercion"],
        formula='=SUM("abc",3)',
        inputs=[],
    )
    _add_case(
        cases,
        prefix="sum_rng_text",
        tags=["agg", "SUM", "coercion"],
        formula="=SUM(A1:A3)",
        inputs=[CellInput("A1", 1), CellInput("A2", "text"), CellInput("A3", 3)],
    )

    # ------------------------------------------------------------------
    # Criteria-string semantics (COUNTIF/SUMIF/AVERAGEIF and *IFS variants)
    # ------------------------------------------------------------------
    # These functions have a large surface area of Excel-compat behaviors:
    # operator parsing, wildcards/escapes, blank/error/boolean matching, and
    # numeric-vs-text coercion.

    # Shared text range for wildcard / escape / blank / error criteria probing.
    criteria_text_inputs = [
        CellInput("A1", "apple"),
        CellInput("A2", "apricot"),
        CellInput("A3", "apex"),
        CellInput("A4", "fax"),
        CellInput("A5", "xylophone"),
        CellInput("A6", "max"),
        CellInput("A7", "*"),
        CellInput("A8", "?"),
        CellInput("A9", "~"),
        CellInput("A10", "~a"),
        CellInput("A11", "a"),
        CellInput("A12", ""),
        CellInput("A13", None),
        CellInput("A14", formula="=1/0"),
        CellInput("A15", formula="=NA()"),
    ]

    for crit in ["ap*", "*x*", "?a?"]:
        _add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "wildcards"],
            formula=f'=COUNTIF(A1:A15,"{crit}")',
            inputs=criteria_text_inputs,
        )

    for crit in ["~*", "~?", "~~", "~a"]:
        _add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "escapes"],
            formula=f'=COUNTIF(A1:A15,"{crit}")',
            inputs=criteria_text_inputs,
        )

    for crit in ["", "=", "<>"]:
        _add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "blanks"],
            formula=f'=COUNTIF(A1:A15,"{crit}")',
            inputs=criteria_text_inputs,
        )

    for crit in ["#DIV/0!", "#N/A"]:
        _add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "errors"],
            formula=f'=COUNTIF(A1:A15,"{crit}")',
            inputs=criteria_text_inputs,
        )

    # Numeric-vs-text and operator parsing.
    criteria_num_inputs = [
        CellInput("B1", 5),
        CellInput("B2", "5"),
        CellInput("B3", 6),
        CellInput("B4", "6"),
        CellInput("B5", 0),
        CellInput("B6", "0"),
        CellInput("B7", 3),
        CellInput("B8", "3"),
        CellInput("B9", None),
        CellInput("B10", ""),
    ]

    for crit_expr in ['">5"', '"<=0"', '"<>3"']:
        _add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "operators"],
            formula=f"=COUNTIF(B1:B10,{crit_expr})",
            inputs=criteria_num_inputs,
        )

    for crit_expr in ['"5"', "5", '"5*"']:
        _add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "numeric-vs-text"],
            formula=f"=COUNTIF(B1:B10,{crit_expr})",
            inputs=criteria_num_inputs,
        )

    # Boolean criteria. Mix booleans, numbers, and text booleans.
    criteria_bool_inputs = [
        CellInput("C1", True),
        CellInput("C2", False),
        CellInput("C3", 1),
        CellInput("C4", 0),
        CellInput("C5", "TRUE"),
        CellInput("C6", "FALSE"),
        CellInput("C7", None),
    ]

    for crit_expr in ["TRUE", "FALSE", '"TRUE"', '"FALSE"']:
        _add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "booleans"],
            formula=f"=COUNTIF(C1:C7,{crit_expr})",
            inputs=criteria_bool_inputs,
        )

    # Date-string criteria against date serial numbers.
    date_serials = [
        _excel_serial_1900(2019, 12, 31),
        _excel_serial_1900(2020, 1, 1),
        _excel_serial_1900(2020, 2, 1),
        _excel_serial_1900(2021, 1, 1),
    ]
    criteria_date_inputs = [CellInput(f"D{i+1}", v) for i, v in enumerate(date_serials)]
    for crit in [">1/1/2020", "<=1/1/2020", "1/1/2020"]:
        _add_case(
            cases,
            prefix="criteria_countif",
            tags=["agg", "criteria", "COUNTIF", "dates"],
            formula=f'=COUNTIF(D1:D4,"{crit}")',
            inputs=criteria_date_inputs,
        )

    # Multi-criteria fixtures (shared by *IFS variants).
    ifs_inputs = [
        CellInput("E1", "apple"),
        CellInput("E2", "apex"),
        CellInput("E3", "banana"),
        CellInput("E4", "fax"),
        CellInput("E5", "max"),
        CellInput("E6", "*"),
        CellInput("E7", None),
        CellInput("F1", 1),
        CellInput("F2", 2),
        CellInput("F3", 3),
        CellInput("F4", 4),
        CellInput("F5", 5),
        CellInput("F6", 6),
        CellInput("F7", 7),
        CellInput("G1", 10),
        CellInput("G2", 20),
        CellInput("G3", 30),
        CellInput("G4", 40),
        CellInput("G5", 50),
        CellInput("G6", 60),
        CellInput("G7", 70),
    ]

    # COUNTIFS basics + wildcards/escapes + blanks.
    _add_case(
        cases,
        prefix="criteria_countifs",
        tags=["agg", "criteria", "COUNTIFS", "wildcards"],
        formula='=COUNTIFS(E1:E7,"*x*",F1:F7,">=4")',
        inputs=ifs_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_countifs",
        tags=["agg", "criteria", "COUNTIFS", "escapes"],
        formula='=COUNTIFS(E1:E7,"~*",F1:F7,">=0")',
        inputs=ifs_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_countifs",
        tags=["agg", "criteria", "COUNTIFS", "blanks"],
        formula='=COUNTIFS(E1:E7,"",F1:F7,">=0")',
        inputs=ifs_inputs,
    )

    # COUNTIFS invalid-arity and shape mismatch should error (#VALUE).
    _add_case(
        cases,
        prefix="criteria_countifs",
        tags=["agg", "criteria", "COUNTIFS", "arg-count"],
        formula='=COUNTIFS(E1:E7,"*x*",F1:F7)',
        inputs=ifs_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_countifs",
        tags=["agg", "criteria", "COUNTIFS", "shape-mismatch"],
        formula='=COUNTIFS(E1:E6,"*x*",F1:F7,">0")',
        inputs=ifs_inputs,
    )

    # SUMIF: wildcards/escapes and operator parsing.
    _add_case(
        cases,
        prefix="criteria_sumif",
        tags=["agg", "criteria", "SUMIF", "wildcards"],
        formula='=SUMIF(E1:E7,"*x*",G1:G7)',
        inputs=ifs_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_sumif",
        tags=["agg", "criteria", "SUMIF", "escapes"],
        formula='=SUMIF(E1:E7,"~*",G1:G7)',
        inputs=ifs_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_sumif",
        tags=["agg", "criteria", "SUMIF", "operators"],
        formula='=SUMIF(F1:F7,">5",G1:G7)',
        inputs=ifs_inputs,
    )

    # SUMIF: errors in the summed range only matter if the row is included.
    sumif_err_inputs = [
        CellInput("I1", 1),
        CellInput("I2", 2),
        CellInput("I3", 3),
        CellInput("J1", 10),
        CellInput("J2", formula="=1/0"),
        CellInput("J3", 30),
    ]
    _add_case(
        cases,
        prefix="criteria_sumif",
        tags=["agg", "criteria", "SUMIF", "sum-range-errors"],
        formula='=SUMIF(I1:I3,">2",J1:J3)',
        inputs=sumif_err_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_sumif",
        tags=["agg", "criteria", "SUMIF", "sum-range-errors"],
        formula='=SUMIF(I1:I3,">1",J1:J3)',
        inputs=sumif_err_inputs,
    )

    # SUMIFS: multi-criteria, invalid-arity, and shape mismatch should error (#VALUE).
    _add_case(
        cases,
        prefix="criteria_sumifs",
        tags=["agg", "criteria", "SUMIFS", "wildcards"],
        formula='=SUMIFS(G1:G7,E1:E7,"*x*",F1:F7,">=4")',
        inputs=ifs_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_sumifs",
        tags=["agg", "criteria", "SUMIFS", "arg-count"],
        formula='=SUMIFS(G1:G7,E1:E7,"*x*",F1:F7)',
        inputs=ifs_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_sumifs",
        tags=["agg", "criteria", "SUMIFS", "shape-mismatch"],
        formula='=SUMIFS(G1:G7,E1:E6,"*x*",F1:F7,">0")',
        inputs=ifs_inputs,
    )

    sumifs_err_inputs = [
        CellInput("K1", 1),
        CellInput("K2", 2),
        CellInput("K3", 3),
        CellInput("K4", 4),
        CellInput("L1", "x"),
        CellInput("L2", "x"),
        CellInput("L3", "y"),
        CellInput("L4", "y"),
        CellInput("M1", 10),
        CellInput("M2", formula="=1/0"),
        CellInput("M3", 30),
        CellInput("M4", 40),
    ]
    _add_case(
        cases,
        prefix="criteria_sumifs",
        tags=["agg", "criteria", "SUMIFS", "sum-range-errors"],
        formula='=SUMIFS(M1:M4,K1:K4,">=3",L1:L4,"y")',
        inputs=sumifs_err_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_sumifs",
        tags=["agg", "criteria", "SUMIFS", "sum-range-errors"],
        formula='=SUMIFS(M1:M4,K1:K4,">=2",L1:L4,"x")',
        inputs=sumifs_err_inputs,
    )

    # AVERAGEIF: #DIV/0! when no numeric values are included.
    _add_case(
        cases,
        prefix="criteria_averageif",
        tags=["agg", "criteria", "AVERAGEIF", "wildcards"],
        formula='=AVERAGEIF(E1:E7,"*x*",G1:G7)',
        inputs=ifs_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_averageif",
        tags=["agg", "criteria", "AVERAGEIF", "operators"],
        formula='=AVERAGEIF(F1:F7,">5",G1:G7)',
        inputs=ifs_inputs,
    )
    avgif_no_numeric_inputs = [
        CellInput("N1", 1),
        CellInput("N2", 2),
        CellInput("N3", 3),
        CellInput("O1", "a"),
        CellInput("O2", "b"),
        CellInput("O3", "c"),
    ]
    _add_case(
        cases,
        prefix="criteria_averageif",
        tags=["agg", "criteria", "AVERAGEIF", "no-numeric"],
        formula='=AVERAGEIF(N1:N3,">0",O1:O3)',
        inputs=avgif_no_numeric_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_averageif",
        tags=["agg", "criteria", "AVERAGEIF", "no-numeric"],
        formula='=AVERAGEIF(N1:N3,">10",N1:N3)',
        inputs=avgif_no_numeric_inputs,
    )

    # AVERAGEIFS: multi-criteria, invalid-arity/shape mismatch (#VALUE), and #DIV/0! on no matches.
    _add_case(
        cases,
        prefix="criteria_averageifs",
        tags=["agg", "criteria", "AVERAGEIFS", "wildcards"],
        formula='=AVERAGEIFS(G1:G7,E1:E7,"*x*",F1:F7,">=4")',
        inputs=ifs_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_averageifs",
        tags=["agg", "criteria", "AVERAGEIFS", "arg-count"],
        formula='=AVERAGEIFS(G1:G7,E1:E7,"*x*",F1:F7)',
        inputs=ifs_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_averageifs",
        tags=["agg", "criteria", "AVERAGEIFS", "shape-mismatch"],
        formula='=AVERAGEIFS(G1:G7,E1:E6,"*x*",F1:F7,">0")',
        inputs=ifs_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_averageifs",
        tags=["agg", "criteria", "AVERAGEIFS", "no-numeric"],
        formula='=AVERAGEIFS(G1:G7,E1:E7,"no_match",F1:F7,">0")',
        inputs=ifs_inputs,
    )

    # ------------------------------------------------------------------
    # Logical functions
    # ------------------------------------------------------------------
    bool_inputs = [True, False]
    for a in bool_inputs:
        for b in bool_inputs:
            _add_case(
                cases,
                prefix="and",
                tags=["logical", "AND"],
                formula="=AND(A1,B1)",
                inputs=[CellInput("A1", a), CellInput("B1", b)],
            )
            _add_case(
                cases,
                prefix="or",
                tags=["logical", "OR"],
                formula="=OR(A1,B1)",
                inputs=[CellInput("A1", a), CellInput("B1", b)],
            )

    for a in [0, 1, -1, "", "0", "1"]:
        _add_case(
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
        _add_case(
            cases,
            prefix="if",
            tags=["logical", "IF"],
            formula="=IF(A1,B1,C1)",
            inputs=[CellInput("A1", cond), CellInput("B1", tv), CellInput("C1", fv)],
            output_cell="D1",
        )

    _add_case(cases, prefix="iferror", tags=["logical", "IFERROR"], formula="=IFERROR(A1,42)", inputs=[CellInput("A1", formula="=1/0")])
    _add_case(cases, prefix="iferror", tags=["logical", "IFERROR"], formula="=IFERROR(A1,42)", inputs=[CellInput("A1", 1)])

    # ------------------------------------------------------------------
    # Text functions (avoid locale-sensitive formatting)
    # ------------------------------------------------------------------
    strings = ["", "a", "foo", "Hello", "12345", "a b c", "This is a test", "こんにちは"]
    num_chars = [0, 1, 2, 3, 5]
    for s in strings:
        _add_case(cases, prefix="len", tags=["text", "LEN"], formula="=LEN(A1)", inputs=[CellInput("A1", s)])
        for n in num_chars:
            _add_case(
                cases,
                prefix="left",
                tags=["text", "LEFT"],
                formula="=LEFT(A1,B1)",
                inputs=[CellInput("A1", s), CellInput("B1", n)],
            )
            _add_case(
                cases,
                prefix="right",
                tags=["text", "RIGHT"],
                formula="=RIGHT(A1,B1)",
                inputs=[CellInput("A1", s), CellInput("B1", n)],
            )

    mid_starts = [1, 2, 3]
    mid_lens = [0, 1, 2]
    for s in strings:
        for start in mid_starts:
            for ln in mid_lens:
                _add_case(
                    cases,
                    prefix="mid",
                    tags=["text", "MID"],
                    formula="=MID(A1,B1,C1)",
                    inputs=[CellInput("A1", s), CellInput("B1", start), CellInput("C1", ln)],
                    output_cell="D1",
                )

    # CONCATENATE is legacy but widely used; CONCAT exists in newer Excel.
    for a in strings:
        for b in ["", "X", "123"]:
            _add_case(
                cases,
                prefix="concat",
                tags=["text", "CONCATENATE"],
                formula="=CONCATENATE(A1,B1)",
                inputs=[CellInput("A1", a), CellInput("B1", b)],
            )

    # FIND / SEARCH differences: FIND is case-sensitive, SEARCH is not.
    find_haystacks = ["foobar", "FooBar", "abcabc"]
    find_needles = ["foo", "Foo", "bar", "z"]
    for needle in find_needles:
        for hay in find_haystacks:
            _add_case(
                cases,
                prefix="find",
                tags=["text", "FIND"],
                formula="=FIND(A1,B1)",
                inputs=[CellInput("A1", needle), CellInput("B1", hay)],
            )
            _add_case(
                cases,
                prefix="search",
                tags=["text", "SEARCH"],
                formula="=SEARCH(A1,B1)",
                inputs=[CellInput("A1", needle), CellInput("B1", hay)],
            )

    # SUBSTITUTE
    for s in ["foo bar foo", "aaaa", "123123", ""]:
        _add_case(
            cases,
            prefix="substitute",
            tags=["text", "SUBSTITUTE"],
            formula='=SUBSTITUTE(A1,"foo","x")',
            inputs=[CellInput("A1", s)],
        )

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
        _add_case(
            cases,
            prefix="date",
            tags=["date", "DATE"],
            formula="=DATE(A1,B1,C1)",
            inputs=[CellInput("A1", y), CellInput("B1", m), CellInput("C1", d)],
            output_cell="D1",
        )
        _add_case(
            cases,
            prefix="year",
            tags=["date", "YEAR"],
            formula="=YEAR(DATE(A1,B1,C1))",
            inputs=[CellInput("A1", y), CellInput("B1", m), CellInput("C1", d)],
            output_cell="D1",
        )
        _add_case(
            cases,
            prefix="month",
            tags=["date", "MONTH"],
            formula="=MONTH(DATE(A1,B1,C1))",
            inputs=[CellInput("A1", y), CellInput("B1", m), CellInput("C1", d)],
            output_cell="D1",
        )
        _add_case(
            cases,
            prefix="day",
            tags=["date", "DAY"],
            formula="=DAY(DATE(A1,B1,C1))",
            inputs=[CellInput("A1", y), CellInput("B1", m), CellInput("C1", d)],
            output_cell="D1",
        )

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
        _add_case(
            cases,
            prefix="vlookup",
            tags=["lookup", "VLOOKUP"],
            formula="=VLOOKUP(D1,A1:B5,2,FALSE)",
            inputs=[*table_inputs, CellInput("D1", key)],
            output_cell="E1",
        )
        _add_case(
            cases,
            prefix="match",
            tags=["lookup", "MATCH"],
            formula="=MATCH(D1,A1:A5,0)",
            inputs=[*table_inputs, CellInput("D1", key)],
            output_cell="E1",
        )
        _add_case(
            cases,
            prefix="index_match",
            tags=["lookup", "INDEX", "MATCH"],
            formula="=INDEX(B1:B5,MATCH(D1,A1:A5,0))",
            inputs=[*table_inputs, CellInput("D1", key)],
            output_cell="E1",
        )

    # ------------------------------------------------------------------
    # Financial functions (P1)
    # ------------------------------------------------------------------
    # Keep this section relatively small; it's mainly intended to cover
    # algorithmic edge cases and common parameter combinations.
    _add_case(cases, prefix="pv", tags=["financial", "PV"], formula="=PV(0, 3, -10)")
    _add_case(cases, prefix="pv", tags=["financial", "PV"], formula="=PV(0.05, 10, -100)")
    _add_case(cases, prefix="fv", tags=["financial", "FV"], formula="=FV(0, 5, -10)")
    _add_case(cases, prefix="fv", tags=["financial", "FV"], formula="=FV(0.05, 10, -100)")
    _add_case(cases, prefix="pmt", tags=["financial", "PMT"], formula="=PMT(0, 2, 10)")
    _add_case(cases, prefix="pmt", tags=["financial", "PMT"], formula="=PMT(0.05, 10, 1000)")
    _add_case(cases, prefix="nper", tags=["financial", "NPER"], formula="=NPER(0, -10, 100)")
    _add_case(cases, prefix="nper", tags=["financial", "NPER"], formula="=NPER(0.05, -100, 1000)")
    _add_case(cases, prefix="rate", tags=["financial", "RATE"], formula="=RATE(10, -100, 1000)")
    _add_case(cases, prefix="rate", tags=["financial", "RATE"], formula="=RATE(12, -50, 500)")
    _add_case(cases, prefix="sln", tags=["financial", "SLN"], formula="=SLN(30, 0, 3)")
    _add_case(cases, prefix="syd", tags=["financial", "SYD"], formula="=SYD(30, 0, 3, 1)")
    _add_case(cases, prefix="ddb", tags=["financial", "DDB"], formula="=DDB(1000, 100, 5, 1)")

    # Range-based cashflow functions.
    cashflows = [-100.0, 30.0, 40.0, 50.0]
    cf_inputs = [CellInput(f"A{i+1}", v) for i, v in enumerate(cashflows)]
    _add_case(
        cases,
        prefix="npv",
        tags=["financial", "NPV"],
        formula="=NPV(0.1,A1:A4)",
        inputs=cf_inputs,
        output_cell="C1",
    )
    _add_case(
        cases,
        prefix="irr",
        tags=["financial", "IRR"],
        formula="=IRR(A1:A4)",
        inputs=cf_inputs,
        output_cell="C1",
    )
    _add_case(
        cases,
        prefix="irr_num",
        tags=["financial", "IRR", "error"],
        formula="=IRR(A1:A3)",
        inputs=[CellInput("A1", 10), CellInput("A2", 20), CellInput("A3", 30)],
        output_cell="C1",
        description="IRR requires at least one positive and one negative cashflow",
    )

    # XNPV/XIRR with explicit date serials (Excel 1900 system with Lotus bug).
    x_values = [-10000.0, 2000.0, 3000.0, 4000.0, 5000.0]
    x_dates = [
        _excel_serial_1900(2020, 1, 1),
        _excel_serial_1900(2020, 7, 1),
        _excel_serial_1900(2021, 1, 1),
        _excel_serial_1900(2021, 7, 1),
        _excel_serial_1900(2022, 1, 1),
    ]
    x_inputs = []
    for i, (v, d) in enumerate(zip(x_values, x_dates), start=1):
        x_inputs.append(CellInput(f"A{i}", v))
        x_inputs.append(CellInput(f"B{i}", d))
    _add_case(
        cases,
        prefix="xnpv",
        tags=["financial", "XNPV"],
        formula="=XNPV(0.1,A1:A5,B1:B5)",
        inputs=x_inputs,
        output_cell="D1",
    )
    _add_case(
        cases,
        prefix="xirr",
        tags=["financial", "XIRR"],
        formula="=XIRR(A1:A5,B1:B5)",
        inputs=x_inputs,
        output_cell="D1",
    )

    # ------------------------------------------------------------------
    # Dynamic arrays / spilling
    # ------------------------------------------------------------------
    _add_case(
        cases,
        prefix="spill_range",
        tags=["spill", "range"],
        formula="=A1:A3",
        inputs=[CellInput("A1", 1), CellInput("A2", 2), CellInput("A3", 3)],
        output_cell="C1",
        description="Reference spill",
    )
    _add_case(
        cases,
        prefix="spill_transpose",
        tags=["spill", "TRANSPOSE"],
        formula="=TRANSPOSE(A1:C1)",
        inputs=[CellInput("A1", 1), CellInput("B1", 2), CellInput("C1", 3)],
        output_cell="E1",
        description="Function spill",
    )
    _add_case(
        cases,
        prefix="spill_sequence",
        tags=["spill", "SEQUENCE", "dynarr"],
        formula="=SEQUENCE(2,2,1,1)",
        inputs=[],
        output_cell="C1",
        description="Dynamic array function (Excel 365+)",
    )

    # ------------------------------------------------------------------
    # Explicit error values
    # ------------------------------------------------------------------
    _add_case(cases, prefix="err_div0", tags=["error"], formula="=1/0", inputs=[], output_cell="A1")
    _add_case(cases, prefix="err_na", tags=["error"], formula="=NA()", inputs=[], output_cell="A1")
    _add_case(cases, prefix="err_name", tags=["error"], formula="=NO_SUCH_FUNCTION(1)", inputs=[], output_cell="A1")

    return {
        "schemaVersion": 1,
        "caseSet": "p0-p1-curated-1k",
        "defaultSheet": "Sheet1",
        "cases": cases,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", required=True, help="Output path for cases.json")
    args = parser.parse_args()

    payload = generate_cases()

    # Stable JSON formatting for review diffs.
    out_path = args.out
    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, sort_keys=False)
        f.write("\n")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
