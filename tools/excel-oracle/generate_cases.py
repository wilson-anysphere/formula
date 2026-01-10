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
