#!/usr/bin/env python3
"""
Deterministically generate a curated (~2k) Excel formula corpus.

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
import re
from pathlib import Path
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
    # Guardrail: the `output_cell` contains the formula under test. If an input also writes to
    # that cell, we can accidentally overwrite the formula or create an unintended circular
    # reference (e.g. output cell participates in a counted range).
    #
    # This has bitten us before in COUNTIF criteria cases, so keep the generator strict.
    inputs = list(inputs)
    if any(i.cell == output_cell for i in inputs):
        raise ValueError(
            f"case output_cell {output_cell!r} collides with an input cell for formula {formula!r}"
        )

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


_FUNC_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")


def _extract_function_names(formula: str | None) -> list[str]:
    if not formula:
        return []
    raw = formula.strip()
    if raw.startswith("="):
        raw = raw[1:]

    out: list[str] = []
    for match in _FUNC_RE.finditer(raw):
        name = match.group(1).upper()
        if name.startswith("_XLFN."):
            name = name[len("_XLFN.") :]
        out.append(name)
    return out


def _validate_against_function_catalog(payload: dict[str, Any]) -> None:
    """
    Keep the oracle corpus aligned with `shared/functionCatalog.json`.

    The goal of the Excel-oracle corpus is to provide end-to-end coverage for
    all deterministic functions. We intentionally exclude volatile functions
    from the corpus because they cannot be pinned/stably compared.

    Coverage is computed from `case.formula` (the formula under test). Input-cell
    formulas are allowed (e.g. `=NA()` to seed an error value), but do not count
    toward function coverage.
    """

    repo_root = Path(__file__).resolve().parents[2]
    catalog_path = repo_root / "shared" / "functionCatalog.json"
    catalog = json.loads(catalog_path.read_text(encoding="utf-8"))

    catalog_nonvolatile: set[str] = set()
    catalog_volatile: set[str] = set()
    for entry in catalog.get("functions", []):
        if not isinstance(entry, dict):
            continue
        name = str(entry.get("name", "")).upper()
        if not name:
            continue
        vol = entry.get("volatility")
        if vol == "volatile":
            catalog_volatile.add(name)
        elif vol == "non_volatile":
            catalog_nonvolatile.add(name)
        else:
            raise SystemExit(f"Unknown volatility in functionCatalog.json for {name!r}: {vol!r}")

    used_case_formulas: set[str] = set()
    used_input_formulas: set[str] = set()
    for case in payload.get("cases", []):
        if not isinstance(case, dict):
            continue
        used_case_formulas.update(_extract_function_names(case.get("formula")))
        # Input cells can also contain formulas (e.g. error values like `=NA()`).
        inputs = case.get("inputs", [])
        if isinstance(inputs, list):
            for cell_input in inputs:
                if not isinstance(cell_input, dict):
                    continue
                used_input_formulas.update(_extract_function_names(cell_input.get("formula")))

    missing_nonvolatile = sorted(catalog_nonvolatile.difference(used_case_formulas))
    if missing_nonvolatile:
        preview = ", ".join(missing_nonvolatile[:25])
        suffix = "" if len(missing_nonvolatile) <= 25 else f" (+{len(missing_nonvolatile) - 25} more)"
        raise SystemExit(
            "Oracle corpus does not cover all deterministic functions in shared/functionCatalog.json "
            "(coverage is based on case.formula only). "
            f"Missing ({len(missing_nonvolatile)}): {preview}{suffix}"
        )

    used_any = used_case_formulas | used_input_formulas
    present_volatile = sorted(catalog_volatile.intersection(used_any))
    if present_volatile:
        raise SystemExit(
            "Oracle corpus must not include volatile functions (non-deterministic). "
            f"Found: {', '.join(present_volatile)}"
        )


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

    # Keep ROUND* combinations capped so the corpus stays <2k cases total.
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
    # Extended math functions (function-catalog backfill)
    # ------------------------------------------------------------------
    _add_case(cases, prefix="pi", tags=["math", "PI"], formula="=PI()")
    _add_case(cases, prefix="sin", tags=["math", "SIN", "PI"], formula="=SIN(PI())")
    _add_case(cases, prefix="cos", tags=["math", "COS"], formula="=COS(0)")
    _add_case(cases, prefix="tan", tags=["math", "TAN"], formula="=TAN(0)")
    _add_case(cases, prefix="acos", tags=["math", "ACOS"], formula="=ACOS(1)")
    _add_case(cases, prefix="asin", tags=["math", "ASIN"], formula="=ASIN(0)")
    _add_case(cases, prefix="atan", tags=["math", "ATAN"], formula="=ATAN(1)")
    _add_case(cases, prefix="atan2", tags=["math", "ATAN2"], formula="=ATAN2(1,1)")
    _add_case(cases, prefix="exp_ln", tags=["math", "EXP", "LN"], formula="=LN(EXP(1))")
    _add_case(cases, prefix="log", tags=["math", "LOG"], formula="=LOG(100,10)")
    _add_case(cases, prefix="log10", tags=["math", "LOG10"], formula="=LOG10(100)")
    _add_case(cases, prefix="power", tags=["math", "POWER"], formula="=POWER(2,3)")
    _add_case(cases, prefix="sqrt", tags=["math", "SQRT"], formula="=SQRT(4)")
    _add_case(cases, prefix="product", tags=["math", "PRODUCT"], formula="=PRODUCT(1,2,3)")
    _add_case(cases, prefix="sumsq", tags=["math", "SUMSQ"], formula="=SUMSQ(1,2,3)")
    _add_case(cases, prefix="trunc", tags=["math", "TRUNC"], formula="=TRUNC(3.14159,2)")
    _add_case(cases, prefix="ceiling", tags=["math", "CEILING"], formula="=CEILING(1.2,1)")
    _add_case(cases, prefix="ceiling_math", tags=["math", "CEILING.MATH"], formula="=CEILING.MATH(1.2,1)")
    _add_case(cases, prefix="ceiling_precise", tags=["math", "CEILING.PRECISE"], formula="=CEILING.PRECISE(1.2,1)")
    _add_case(cases, prefix="iso_ceiling", tags=["math", "ISO.CEILING"], formula="=ISO.CEILING(1.2,1)")
    _add_case(cases, prefix="floor", tags=["math", "FLOOR"], formula="=FLOOR(1.2,1)")
    _add_case(cases, prefix="floor_math", tags=["math", "FLOOR.MATH"], formula="=FLOOR.MATH(1.2,1)")
    _add_case(cases, prefix="floor_precise", tags=["math", "FLOOR.PRECISE"], formula="=FLOOR.PRECISE(1.2,1)")

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

    # COUNT / COUNTA / COUNTBLANK
    count_range_inputs = [
        CellInput("A1", 1),
        CellInput("A2", "x"),
        CellInput("A3", True),
        CellInput("A4", None),  # blank
        CellInput("A5", ""),  # empty string counts as blank for COUNTBLANK
        CellInput("A6", formula="=1/0"),  # errors are ignored by COUNT* in Excel
    ]
    _add_case(cases, prefix="count", tags=["agg", "COUNT"], formula="=COUNT(A1:A6)", inputs=count_range_inputs)
    _add_case(cases, prefix="counta", tags=["agg", "COUNTA"], formula="=COUNTA(A1:A6)", inputs=count_range_inputs)
    _add_case(cases, prefix="countblank", tags=["agg", "COUNTBLANK"], formula="=COUNTBLANK(A1:A6)", inputs=count_range_inputs)

    # COUNTIF (include numeric, text + wildcards, blanks)
    countif_num_inputs = [CellInput("A1", 1), CellInput("A2", 2), CellInput("A3", 3), CellInput("A4", 4)]
    _add_case(
        cases,
        prefix="countif",
        tags=["agg", "COUNTIF"],
        formula='=COUNTIF(A1:A4,">2")',
        inputs=countif_num_inputs,
    )
    countif_text_inputs = [
        CellInput("A1", "apple"),
        CellInput("A2", "banana"),
        CellInput("A3", "apricot"),
        CellInput("A4", None),
        CellInput("A5", ""),
    ]
    _add_case(
        cases,
        prefix="countif",
        tags=["agg", "COUNTIF"],
        formula='=COUNTIF(A1:A5,"ap*")',
        inputs=countif_text_inputs,
    )
    _add_case(
        cases,
        prefix="countif",
        tags=["agg", "COUNTIF"],
        formula='=COUNTIF(A1:A5,"")',
        inputs=countif_text_inputs,
        description="Blank criteria matches truly blank cells and empty-string cells",
    )

    # SUMPRODUCT
    sumproduct_inputs = [
        CellInput("A1", 1),
        CellInput("A2", 2),
        CellInput("A3", 3),
        CellInput("B1", 4),
        CellInput("B2", 5),
        CellInput("B3", 6),
    ]
    _add_case(
        cases,
        prefix="sumproduct",
        tags=["agg", "SUMPRODUCT"],
        formula="=SUMPRODUCT(A1:A3,B1:B3)",
        inputs=sumproduct_inputs,
    )
    _add_case(
        cases,
        prefix="sumproduct",
        tags=["agg", "SUMPRODUCT", "error"],
        formula="=SUMPRODUCT(A1:A2,B1:B2)",
        inputs=[
            CellInput("A1", 1),
            CellInput("A2", formula="=1/0"),
            CellInput("B1", 2),
            CellInput("B2", 3),
        ],
        description="SUMPRODUCT propagates errors from any element",
    )

    subtotal_inputs = [CellInput("A1", 1), CellInput("A2", 2), CellInput("A3", 3)]
    _add_case(
        cases,
        prefix="subtotal",
        tags=["agg", "SUBTOTAL"],
        formula="=SUBTOTAL(9,A1:A3)",
        inputs=subtotal_inputs,
    )
    _add_case(
        cases,
        prefix="aggregate",
        tags=["agg", "AGGREGATE"],
        formula="=AGGREGATE(9,4,A1:A3)",
        inputs=subtotal_inputs,
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
            output_cell="D1",
        )

    # Date criteria against date serial numbers.
    #
    # Avoid locale-dependent date parsing (e.g. "1/1/2020") by building criteria
    # from DATE(...) so the criterion is numeric and stable across locales.
    date_serials = [
        _excel_serial_1900(2019, 12, 31),
        _excel_serial_1900(2020, 1, 1),
        _excel_serial_1900(2020, 2, 1),
        _excel_serial_1900(2021, 1, 1),
    ]
    criteria_date_inputs = [CellInput(f"D{i+1}", v) for i, v in enumerate(date_serials)]
    _add_case(
        cases,
        prefix="criteria_countif",
        tags=["agg", "criteria", "COUNTIF", "dates"],
        formula='=COUNTIF(D1:D4,">"&DATE(2020,1,1))',
        inputs=criteria_date_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_countif",
        tags=["agg", "criteria", "COUNTIF", "dates"],
        formula='=COUNTIF(D1:D4,"<="&DATE(2020,1,1))',
        inputs=criteria_date_inputs,
    )
    _add_case(
        cases,
        prefix="criteria_countif",
        tags=["agg", "criteria", "COUNTIF", "dates"],
        formula="=COUNTIF(D1:D4,DATE(2020,1,1))",
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
    # Statistical / regression functions (function-catalog backfill)
    # ------------------------------------------------------------------
    # Prefer array literals to keep the corpus compact and deterministic.
    _add_case(cases, prefix="avedev", tags=["stat", "AVEDEV"], formula="=AVEDEV({1,2,3,4})")
    _add_case(cases, prefix="averagea", tags=["stat", "AVERAGEA"], formula='=AVERAGEA({1,"x",TRUE})')
    _add_case(cases, prefix="maxa", tags=["stat", "MAXA"], formula='=MAXA({1,"x",TRUE})')
    _add_case(cases, prefix="mina", tags=["stat", "MINA"], formula='=MINA({1,"x",FALSE})')
    _add_case(cases, prefix="median", tags=["stat", "MEDIAN"], formula="=MEDIAN({1,2,3,4})")
    _add_case(cases, prefix="mode", tags=["stat", "MODE"], formula="=MODE({1,1,2,3})")
    _add_case(cases, prefix="mode_sngl", tags=["stat", "MODE.SNGL"], formula="=MODE.SNGL({1,1,2,3})")
    _add_case(cases, prefix="mode_mult", tags=["stat", "MODE.MULT"], formula="=MODE.MULT({1,1,2,2,3})", output_cell="C1")

    _add_case(cases, prefix="devsq", tags=["stat", "DEVSQ"], formula="=DEVSQ({1,2,3})")
    _add_case(cases, prefix="geomean", tags=["stat", "GEOMEAN"], formula="=GEOMEAN({1,2,3,4})")
    _add_case(cases, prefix="harmean", tags=["stat", "HARMEAN"], formula="=HARMEAN({1,2,4})")

    _add_case(cases, prefix="large", tags=["stat", "LARGE"], formula="=LARGE({1,2,3,4},2)")
    _add_case(cases, prefix="small", tags=["stat", "SMALL"], formula="=SMALL({1,2,3,4},2)")

    _add_case(cases, prefix="percentile", tags=["stat", "PERCENTILE"], formula="=PERCENTILE({1,2,3,4},0.25)")
    _add_case(cases, prefix="percentile_inc", tags=["stat", "PERCENTILE.INC"], formula="=PERCENTILE.INC({1,2,3,4},0.25)")
    _add_case(cases, prefix="percentile_exc", tags=["stat", "PERCENTILE.EXC"], formula="=PERCENTILE.EXC({1,2,3,4},0.25)")

    # PERCENTRANK and variants: use a 1..9 set so both inclusive and exclusive variants
    # yield simple finite decimals without needing explicit `significance`.
    _add_case(cases, prefix="percentrank", tags=["stat", "PERCENTRANK"], formula="=PERCENTRANK({1,2,3,4,5,6,7,8,9},2)")
    _add_case(cases, prefix="percentrank_inc", tags=["stat", "PERCENTRANK.INC"], formula="=PERCENTRANK.INC({1,2,3,4,5,6,7,8,9},2)")
    _add_case(cases, prefix="percentrank_exc", tags=["stat", "PERCENTRANK.EXC"], formula="=PERCENTRANK.EXC({1,2,3,4,5,6,7,8,9},2)")

    _add_case(cases, prefix="quartile", tags=["stat", "QUARTILE"], formula="=QUARTILE({1,2,3,4},1)")
    _add_case(cases, prefix="quartile_inc", tags=["stat", "QUARTILE.INC"], formula="=QUARTILE.INC({1,2,3,4},1)")
    _add_case(cases, prefix="quartile_exc", tags=["stat", "QUARTILE.EXC"], formula="=QUARTILE.EXC({1,2,3,4},1)")

    _add_case(cases, prefix="rank", tags=["stat", "RANK"], formula="=RANK(2,{1,2,2,3})")
    _add_case(cases, prefix="rank_eq", tags=["stat", "RANK.EQ"], formula="=RANK.EQ(2,{1,2,2,3})")
    _add_case(cases, prefix="rank_avg", tags=["stat", "RANK.AVG"], formula="=RANK.AVG(2,{1,2,2,3})")

    _add_case(cases, prefix="stdev", tags=["stat", "STDEV"], formula="=STDEV({1,2,3,4})")
    _add_case(cases, prefix="stdev_s", tags=["stat", "STDEV.S"], formula="=STDEV.S({1,2,3,4})")
    _add_case(cases, prefix="stdev_p", tags=["stat", "STDEV.P"], formula="=STDEV.P({1,2,3,4})")
    _add_case(cases, prefix="stdeva", tags=["stat", "STDEVA"], formula="=STDEVA({1,2,3,TRUE})")
    _add_case(cases, prefix="stdevp", tags=["stat", "STDEVP"], formula="=STDEVP({1,2,3,TRUE})")
    _add_case(cases, prefix="stdevpa", tags=["stat", "STDEVPA"], formula="=STDEVPA({1,2,3,TRUE})")

    _add_case(cases, prefix="var", tags=["stat", "VAR"], formula="=VAR({1,2,3,4})")
    _add_case(cases, prefix="var_s", tags=["stat", "VAR.S"], formula="=VAR.S({1,2,3,4})")
    _add_case(cases, prefix="var_p", tags=["stat", "VAR.P"], formula="=VAR.P({1,2,3,4})")
    _add_case(cases, prefix="vara", tags=["stat", "VARA"], formula="=VARA({1,2,3,TRUE})")
    _add_case(cases, prefix="varp", tags=["stat", "VARP"], formula="=VARP({1,2,3,TRUE})")
    _add_case(cases, prefix="varpa", tags=["stat", "VARPA"], formula="=VARPA({1,2,3,TRUE})")

    _add_case(cases, prefix="trimmean", tags=["stat", "TRIMMEAN"], formula="=TRIMMEAN({1,2,3,100},0.5)")

    _add_case(cases, prefix="standardize", tags=["stat", "STANDARDIZE"], formula="=STANDARDIZE(1,3,2)")

    _add_case(cases, prefix="correl", tags=["stat", "CORREL"], formula="=CORREL({1,2,3},{1,5,7})")
    _add_case(cases, prefix="pearson", tags=["stat", "PEARSON"], formula="=PEARSON({1,2,3},{1,5,7})")
    _add_case(cases, prefix="covar", tags=["stat", "COVAR"], formula="=COVAR({1,2,3},{1,5,7})")
    _add_case(cases, prefix="cov_p", tags=["stat", "COVARIANCE.P"], formula="=COVARIANCE.P({1,2,3},{1,5,7})")
    _add_case(cases, prefix="cov_s", tags=["stat", "COVARIANCE.S"], formula="=COVARIANCE.S({1,2,3},{1,5,7})")

    _add_case(cases, prefix="rsq", tags=["stat", "RSQ"], formula="=RSQ({1,2,3},{1,2,3})")
    _add_case(cases, prefix="slope", tags=["stat", "SLOPE"], formula="=SLOPE({1,2,3},{1,2,3})")
    _add_case(cases, prefix="intercept", tags=["stat", "INTERCEPT"], formula="=INTERCEPT({1,2,3},{1,2,3})")
    _add_case(cases, prefix="forecast", tags=["stat", "FORECAST"], formula="=FORECAST(4,{1,2,3},{1,2,3})")
    _add_case(cases, prefix="forecast_linear", tags=["stat", "FORECAST.LINEAR"], formula="=FORECAST.LINEAR(4,{1,2,3},{1,2,3})")

    # STEYX should be 0 for a perfectly linear relationship (y = 2x + 1).
    _add_case(cases, prefix="steyx", tags=["stat", "STEYX"], formula="=STEYX({3,5,7,9,11},{1,2,3,4,5})")

    # MAXIFS / MINIFS (criteria-based aggregates)
    maxifs_inputs = [
        CellInput("A1", 10),
        CellInput("A2", 20),
        CellInput("A3", 30),
        CellInput("B1", "A"),
        CellInput("B2", "B"),
        CellInput("B3", "A"),
    ]
    _add_case(
        cases,
        prefix="maxifs",
        tags=["agg", "MAXIFS"],
        formula='=MAXIFS(A1:A3,B1:B3,"A")',
        inputs=maxifs_inputs,
    )
    _add_case(
        cases,
        prefix="minifs",
        tags=["agg", "MINIFS"],
        formula='=MINIFS(A1:A3,B1:B3,"A")',
        inputs=maxifs_inputs,
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

    # Excel supports `TRUE`/`FALSE` as both logical constants and zero-arg functions.
    _add_case(cases, prefix="true", tags=["logical", "TRUE"], formula="=TRUE()")
    _add_case(cases, prefix="false", tags=["logical", "FALSE"], formula="=FALSE()")

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
    _add_case(cases, prefix="ifna", tags=["logical", "IFNA"], formula="=IFNA(A1,42)", inputs=[CellInput("A1", formula="=NA()")])
    _add_case(
        cases,
        prefix="ifna",
        tags=["logical", "IFNA", "error"],
        formula="=IFNA(A1,42)",
        inputs=[CellInput("A1", formula="=1/0")],
        description="IFNA only catches #N/A (other errors propagate)",
    )
    _add_case(cases, prefix="ifna", tags=["logical", "IFNA"], formula="=IFNA(A1,42)", inputs=[CellInput("A1", 1)])

    _add_case(cases, prefix="ifs", tags=["logical", "IFS"], formula="=IFS(FALSE,1,TRUE,2)")
    _add_case(
        cases,
        prefix="switch",
        tags=["logical", "SWITCH"],
        formula='=SWITCH(2,1,"one",2,"two","other")',
    )
    _add_case(cases, prefix="xor", tags=["logical", "XOR"], formula="=XOR(TRUE,FALSE,TRUE)")

    # ------------------------------------------------------------------
    # Value coercion / conversion
    # ------------------------------------------------------------------
    # These cases are explicitly chosen to validate the coercion rules we
    # implement (text -> number/date/time), so we can diff against real Excel
    # later. Keep the set small to avoid bloating the corpus.

    # Implicit coercion (text used in arithmetic/logical contexts).
    _add_case(cases, prefix="coercion", tags=["coercion", "implicit", "add"], formula='=1+""')
    _add_case(cases, prefix="coercion", tags=["coercion", "implicit"], formula='=--""')
    _add_case(cases, prefix="coercion", tags=["coercion", "implicit", "NOT"], formula='=NOT("")')
    _add_case(cases, prefix="coercion", tags=["coercion", "implicit", "IF"], formula='=IF("",10,20)')
    _add_case(cases, prefix="coercion", tags=["coercion", "implicit", "add"], formula='="1234"+1')
    _add_case(cases, prefix="coercion", tags=["coercion", "implicit", "add"], formula='="(1000)"+0')
    _add_case(cases, prefix="coercion", tags=["coercion", "implicit", "mul"], formula='="10%"*100')
    _add_case(cases, prefix="coercion", tags=["coercion", "implicit", "add"], formula='=" 1234 "+0')

    # Explicit conversion functions.
    _add_case(cases, prefix="value", tags=["coercion", "VALUE"], formula='=VALUE("2020-01-01")')
    _add_case(cases, prefix="value", tags=["coercion", "VALUE"], formula='=VALUE("2020-01-01 13:30")')
    _add_case(cases, prefix="timevalue", tags=["coercion", "TIMEVALUE"], formula='=TIMEVALUE("13:00")')

    # ------------------------------------------------------------------
    # Text functions (keep cases deterministic; avoid locale-dependent parsing where possible)
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

    # Additional text functions
    _add_case(
        cases,
        prefix="clean",
        tags=["text", "CLEAN"],
        formula="=CLEAN(A1)",
        inputs=[CellInput("A1", "a\u0000\u0009b\u001Fc\u007Fd")],
        description="CLEAN strips non-printable ASCII control codes",
    )
    _add_case(
        cases,
        prefix="trim",
        tags=["text", "TRIM"],
        formula="=TRIM(A1)",
        inputs=[CellInput("A1", "  a   b  ")],
    )
    _add_case(
        cases,
        prefix="trim",
        tags=["text", "TRIM"],
        formula="=TRIM(A1)",
        inputs=[CellInput("A1", "\ta  b")],
        description="TRIM collapses spaces but preserves tabs",
    )
    _add_case(cases, prefix="upper", tags=["text", "UPPER"], formula="=UPPER(A1)", inputs=[CellInput("A1", "Abc")])
    _add_case(cases, prefix="lower", tags=["text", "LOWER"], formula="=LOWER(A1)", inputs=[CellInput("A1", "AbC")])
    _add_case(
        cases,
        prefix="proper",
        tags=["text", "PROPER"],
        formula="=PROPER(A1)",
        inputs=[CellInput("A1", "hELLO wORLD")],
    )
    _add_case(
        cases,
        prefix="exact",
        tags=["text", "EXACT"],
        formula="=EXACT(A1,B1)",
        inputs=[CellInput("A1", "Hello"), CellInput("B1", "hello")],
    )
    _add_case(
        cases,
        prefix="exact",
        tags=["text", "EXACT"],
        formula="=EXACT(A1,B1)",
        inputs=[CellInput("A1", "Hello"), CellInput("B1", "Hello")],
    )
    _add_case(cases, prefix="replace", tags=["text", "REPLACE"], formula='=REPLACE("abcdef",2,3,"X")')
    _add_case(cases, prefix="replace", tags=["text", "REPLACE"], formula='=REPLACE("abc",5,1,"X")')

    # CONCAT (unlike CONCATENATE, CONCAT flattens ranges)
    _add_case(
        cases,
        prefix="concat_new",
        tags=["text", "CONCAT"],
        formula='=CONCAT(A1:A2,"c")',
        inputs=[CellInput("A1", "a"), CellInput("A2", "b")],
    )

    # TEXTJOIN
    textjoin_inputs = [
        CellInput("A1", "a"),
        CellInput("A2", None),
        CellInput("A3", ""),
        CellInput("A4", 1),
    ]
    _add_case(
        cases,
        prefix="textjoin",
        tags=["text", "TEXTJOIN"],
        formula='=TEXTJOIN(",",TRUE,A1:A4)',
        inputs=textjoin_inputs,
    )
    _add_case(
        cases,
        prefix="textjoin",
        tags=["text", "TEXTJOIN"],
        formula='=TEXTJOIN(",",FALSE,A1:A4)',
        inputs=textjoin_inputs,
    )

    # TEXTSPLIT (dynamic array)
    _add_case(
        cases,
        prefix="textsplit_basic",
        tags=["text", "TEXTSPLIT"],
        formula='=TEXTSPLIT("a,b,c",",")',
    )

    # TEXT / VALUE / NUMBERVALUE / DOLLAR
    # TEXT formatting: keep these cases locale-independent by avoiding thousand/decimal separators
    # and currency symbols in the *result* string.
    _add_case(cases, prefix="text_fmt", tags=["text", "TEXT"], formula='=TEXT(1234.567,"0")')
    _add_case(cases, prefix="text_pct", tags=["text", "TEXT"], formula='=TEXT(1.23,"0%")')
    _add_case(cases, prefix="text_int", tags=["text", "TEXT"], formula='=TEXT(-1,"0")')
    _add_case(cases, prefix="value", tags=["text", "VALUE", "coercion"], formula='=VALUE("1234")')
    _add_case(cases, prefix="value", tags=["text", "VALUE"], formula='=VALUE("(1000)")')
    _add_case(cases, prefix="value", tags=["text", "VALUE"], formula='=VALUE("10%")')
    _add_case(cases, prefix="value", tags=["text", "VALUE", "error"], formula='=VALUE("nope")')
    _add_case(
        cases,
        prefix="numbervalue",
        tags=["text", "NUMBERVALUE", "coercion"],
        formula='=NUMBERVALUE("1.234,5", ",", ".")',
    )
    _add_case(
        cases,
        prefix="numbervalue",
        tags=["text", "NUMBERVALUE", "error"],
        formula='=NUMBERVALUE("1,23", ",", ",")',
    )
    # DOLLAR returns a localized currency string. Wrap in N(...) so the case result is
    # locale-independent while still exercising the function.
    _add_case(cases, prefix="dollar", tags=["text", "DOLLAR"], formula="=N(DOLLAR(1234.567,2))")
    _add_case(cases, prefix="dollar", tags=["text", "DOLLAR"], formula="=N(DOLLAR(-1234.567,2))")

    # TEXT: Excel number format codes (dates/sections/conditions). These are
    # locale-sensitive, but high-signal for Excel compatibility.
    _add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format"],
        formula=r'=TEXT(A1,"0")',
        inputs=[CellInput("A1", 1234.567)],
    )
    multi_section = '"0.00;(0.00);""zero"";""text:""@"'
    _add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "sections"],
        formula=f"=TEXT(A1,{multi_section})",
        inputs=[CellInput("A1", 1.2)],
    )
    _add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "sections"],
        formula=f"=TEXT(A1,{multi_section})",
        inputs=[CellInput("A1", -1.2)],
    )
    _add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "sections"],
        formula=f"=TEXT(A1,{multi_section})",
        inputs=[CellInput("A1", 0)],
    )
    _add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "sections"],
        formula=f"=TEXT(A1,{multi_section})",
        inputs=[CellInput("A1", "hi")],
    )
    _add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "conditions"],
        formula=r'=TEXT(A1,"[<0]""neg"";""pos""")',
        inputs=[CellInput("A1", -1)],
    )
    _add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "conditions"],
        formula=r'=TEXT(A1,"[<0]""neg"";""pos""")',
        inputs=[CellInput("A1", 1)],
    )
    _add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "date"],
        formula=r'=TEXT(A1,"yyyy-mm-dd")',
        inputs=[CellInput("A1", _excel_serial_1900(2024, 1, 10))],
    )
    _add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "date", "time"],
        formula=r'=TEXT(A1,"yyyy-mm-dd hh:mm")',
        inputs=[CellInput("A1", _excel_serial_1900(2024, 1, 10) + 0.5)],
    )
    _add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "locale"],
        formula=r'=TEXT(A1,"[$€-407]0")',
        inputs=[CellInput("A1", 1)],
        description="Currency symbol bracket token + locale code (LCID 0x0407 de-DE)",
    )
    _add_case(
        cases,
        prefix="text_fmt",
        tags=["text", "TEXT", "format", "invalid"],
        formula=r'=TEXT(A1,"")',
        inputs=[CellInput("A1", 1234.5)],
        description="Empty format_text should fall back to General (Excel behavior)",
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

    # Additional date/time functions (keep results numeric to avoid locale-dependent display text).
    _add_case(cases, prefix="datevalue", tags=["date", "DATEVALUE", "coercion"], formula='=DATEVALUE("2020-01-01")')
    _add_case(cases, prefix="datevalue", tags=["date", "DATEVALUE", "error"], formula='=DATEVALUE("nope")')

    _add_case(cases, prefix="time", tags=["date", "TIME"], formula="=TIME(1,30,0)")
    _add_case(cases, prefix="time", tags=["date", "TIME"], formula="=TIME(24,0,0)")
    _add_case(cases, prefix="time", tags=["date", "TIME", "error"], formula="=TIME(-1,0,0)")

    _add_case(cases, prefix="timevalue", tags=["date", "TIMEVALUE"], formula='=TIMEVALUE("1:30")')
    _add_case(cases, prefix="timevalue", tags=["date", "TIMEVALUE", "coercion"], formula='=TIMEVALUE("13:30")')
    _add_case(cases, prefix="timevalue", tags=["date", "TIMEVALUE", "error"], formula='=TIMEVALUE("nope")')

    _add_case(cases, prefix="hour", tags=["date", "HOUR"], formula="=HOUR(TIME(1,2,3))")
    _add_case(cases, prefix="minute", tags=["date", "MINUTE"], formula="=MINUTE(TIME(1,2,3))")
    _add_case(cases, prefix="second", tags=["date", "SECOND"], formula="=SECOND(TIME(1,2,3))")

    _add_case(cases, prefix="edate", tags=["date", "EDATE"], formula="=EDATE(DATE(2020,1,31),1)")
    _add_case(cases, prefix="eomonth", tags=["date", "EOMONTH"], formula="=EOMONTH(DATE(2020,1,15),0)")
    _add_case(cases, prefix="eomonth", tags=["date", "EOMONTH"], formula="=EOMONTH(DATE(2020,1,15),1)")

    _add_case(cases, prefix="weekday", tags=["date", "WEEKDAY"], formula="=WEEKDAY(1)")
    _add_case(cases, prefix="weekday", tags=["date", "WEEKDAY"], formula="=WEEKDAY(1,2)")
    _add_case(cases, prefix="weekday", tags=["date", "WEEKDAY", "error"], formula="=WEEKDAY(1,0)")

    _add_case(cases, prefix="weeknum", tags=["date", "WEEKNUM"], formula="=WEEKNUM(DATE(2020,1,1),1)")
    _add_case(cases, prefix="weeknum", tags=["date", "WEEKNUM"], formula="=WEEKNUM(DATE(2020,1,5),2)")
    _add_case(cases, prefix="weeknum", tags=["date", "WEEKNUM"], formula="=WEEKNUM(DATE(2021,1,1),21)")
    _add_case(cases, prefix="weeknum", tags=["date", "WEEKNUM", "error"], formula="=WEEKNUM(1,9)")

    _add_case(cases, prefix="workday", tags=["date", "WORKDAY"], formula="=WORKDAY(DATE(2020,1,1),1)")
    _add_case(
        cases,
        prefix="workday",
        tags=["date", "WORKDAY"],
        formula="=WORKDAY(DATE(2020,1,1),1,DATE(2020,1,2))",
    )

    _add_case(
        cases,
        prefix="networkdays",
        tags=["date", "NETWORKDAYS"],
        formula="=NETWORKDAYS(DATE(2020,1,1),DATE(2020,1,10))",
    )
    _add_case(
        cases,
        prefix="networkdays",
        tags=["date", "NETWORKDAYS"],
        formula="=NETWORKDAYS(DATE(2020,1,1),DATE(2020,1,10),{DATE(2020,1,2),DATE(2020,1,3)})",
    )

    _add_case(cases, prefix="workday_intl", tags=["date", "WORKDAY.INTL"], formula="=WORKDAY.INTL(DATE(2020,1,3),1,11)")
    _add_case(cases, prefix="workday_intl", tags=["date", "WORKDAY.INTL", "error"], formula="=WORKDAY.INTL(DATE(2020,1,3),1,99)")
    _add_case(cases, prefix="workday_intl", tags=["date", "WORKDAY.INTL", "error"], formula='=WORKDAY.INTL(DATE(2020,1,3),1,"abc")')

    _add_case(
        cases,
        prefix="networkdays_intl",
        tags=["date", "NETWORKDAYS.INTL"],
        formula="=NETWORKDAYS.INTL(DATE(2020,1,1),DATE(2020,1,10),11)",
    )
    _add_case(
        cases,
        prefix="networkdays_intl",
        tags=["date", "NETWORKDAYS.INTL", "error"],
        formula="=NETWORKDAYS.INTL(DATE(2020,1,1),DATE(2020,1,10),99)",
    )

    _add_case(cases, prefix="days", tags=["date", "DAYS"], formula="=DAYS(DATE(2020,1,10),DATE(2020,1,1))")
    _add_case(cases, prefix="days360", tags=["date", "DAYS360"], formula="=DAYS360(DATE(2020,1,1),DATE(2020,2,1))")
    _add_case(cases, prefix="datedif", tags=["date", "DATEDIF"], formula='=DATEDIF(DATE(2020,1,1),DATE(2021,2,1),"y")')
    _add_case(cases, prefix="yearfrac", tags=["date", "YEARFRAC"], formula="=YEARFRAC(DATE(2020,1,1),DATE(2021,1,1))")
    _add_case(cases, prefix="iso_weeknum", tags=["date", "ISO.WEEKNUM"], formula="=ISO.WEEKNUM(DATE(2021,1,1))")
    _add_case(cases, prefix="isoweeknum", tags=["date", "ISOWEEKNUM"], formula="=ISOWEEKNUM(DATE(2021,1,1))")

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

    # HLOOKUP uses a horizontal table (keys in first row).
    h_keys = [1, 2, 3, 10, 20]
    h_vals = ["one", "two", "three", "ten", "twenty"]
    h_inputs = []
    for col, (k, v) in enumerate(zip(h_keys, h_vals), start=1):
        col_letter = chr(ord("A") + col - 1)
        h_inputs.append(CellInput(f"{col_letter}1", k))
        h_inputs.append(CellInput(f"{col_letter}2", v))
    for key in [1, 2, 3, 10, 20, 4]:
        _add_case(
            cases,
            prefix="hlookup",
            tags=["lookup", "HLOOKUP"],
            formula="=HLOOKUP(F1,A1:E2,2,FALSE)",
            inputs=[*h_inputs, CellInput("F1", key)],
            output_cell="G1",
        )

    # XLOOKUP / XMATCH (Excel 365+)
    for key in [1, 2, 3, 10, 20, 4]:
        _add_case(
            cases,
            prefix="xlookup",
            tags=["lookup", "XLOOKUP"],
            formula='=XLOOKUP(D1,A1:A5,B1:B5,"NF")',
            inputs=[*table_inputs, CellInput("D1", key)],
            output_cell="E1",
        )
        _add_case(
            cases,
            prefix="xmatch",
            tags=["lookup", "XMATCH"],
            formula="=XMATCH(D1,A1:A5,0)",
            inputs=[*table_inputs, CellInput("D1", key)],
            output_cell="E1",
        )

    # GETPIVOTDATA requires a real pivot table; without one, Excel returns #REF!.
    _add_case(
        cases,
        prefix="getpivotdata",
        tags=["lookup", "GETPIVOTDATA", "error"],
        formula='=GETPIVOTDATA("Sales",A1)',
        inputs=[CellInput("A1", 0)],
        output_cell="C1",
        description="No pivot table present; GETPIVOTDATA should return #REF!",
    )

    # Reference helpers
    _add_case(cases, prefix="address", tags=["ref", "ADDRESS"], formula="=ADDRESS(2,3)")
    _add_case(cases, prefix="row", tags=["ref", "ROW"], formula="=ROW(A10)")
    _add_case(cases, prefix="column", tags=["ref", "COLUMN"], formula="=COLUMN(C5)")
    _add_case(cases, prefix="rows", tags=["ref", "ROWS"], formula="=ROWS(A1:B3)")
    _add_case(cases, prefix="columns", tags=["ref", "COLUMNS"], formula="=COLUMNS(A1:B3)")

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
    _add_case(cases, prefix="ipmt", tags=["financial", "IPMT"], formula="=IPMT(0.05, 1, 10, 1000)")
    _add_case(cases, prefix="ipmt", tags=["financial", "IPMT"], formula="=IPMT(0.05, 10, 10, 1000, 0, 1)")
    _add_case(cases, prefix="ppmt", tags=["financial", "PPMT"], formula="=PPMT(0.05, 1, 10, 1000)")
    _add_case(cases, prefix="ppmt", tags=["financial", "PPMT"], formula="=PPMT(0.05, 10, 10, 1000, 0, 1)")
    _add_case(cases, prefix="nper", tags=["financial", "NPER"], formula="=NPER(0, -10, 100)")
    _add_case(cases, prefix="nper", tags=["financial", "NPER"], formula="=NPER(0.05, -100, 1000)")
    _add_case(cases, prefix="rate", tags=["financial", "RATE"], formula="=RATE(10, -100, 1000)")
    _add_case(cases, prefix="rate", tags=["financial", "RATE"], formula="=RATE(12, -50, 500)")
    _add_case(cases, prefix="effect", tags=["financial", "EFFECT"], formula="=EFFECT(0.1,12)")
    _add_case(cases, prefix="nominal", tags=["financial", "NOMINAL"], formula="=NOMINAL(0.1,12)")
    _add_case(cases, prefix="rri", tags=["financial", "RRI"], formula="=RRI(10,-100,200)")
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
        prefix="mirr",
        tags=["financial", "MIRR"],
        formula="=MIRR(A1:A4,0.1,0.12)",
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
    _add_case(
        cases,
        prefix="spill_textsplit",
        tags=["spill", "TEXTSPLIT", "dynarr"],
        formula='=TEXTSPLIT("a,b,c",",")',
        description="Dynamic array function (Excel 365+)",
    )

    # FILTER / SORT / UNIQUE (simple spill cases)
    filter_inputs = [CellInput(f"A{i}", i) for i in range(1, 6)]
    _add_case(
        cases,
        prefix="spill_filter",
        tags=["spill", "FILTER", "dynarr"],
        formula="=FILTER(A1:A5,A1:A5>2)",
        inputs=filter_inputs,
        output_cell="C1",
    )
    _add_case(
        cases,
        prefix="spill_filter",
        tags=["spill", "FILTER", "dynarr"],
        formula='=FILTER(A1:A5,A1:A5>10,"none")',
        inputs=filter_inputs,
        output_cell="C1",
        description="if_empty fallback (no matches)",
    )

    sort_inputs = [CellInput("A1", 3), CellInput("A2", 1), CellInput("A3", 2)]
    _add_case(
        cases,
        prefix="spill_sort",
        tags=["spill", "SORT", "dynarr"],
        formula="=SORT(A1:A3)",
        inputs=sort_inputs,
        output_cell="C1",
    )
    _add_case(
        cases,
        prefix="spill_sort",
        tags=["spill", "SORT", "dynarr"],
        formula="=SORT(A1:A3,1,-1)",
        inputs=sort_inputs,
        output_cell="C1",
    )

    unique_inputs = [
        CellInput("A1", 1),
        CellInput("A2", 1),
        CellInput("A3", 2),
        CellInput("A4", 3),
        CellInput("A5", 3),
        CellInput("A6", 3),
    ]
    _add_case(
        cases,
        prefix="spill_unique",
        tags=["spill", "UNIQUE", "dynarr"],
        formula="=UNIQUE(A1:A6)",
        inputs=unique_inputs,
        output_cell="C1",
    )
    _add_case(
        cases,
        prefix="spill_unique",
        tags=["spill", "UNIQUE", "dynarr"],
        formula="=UNIQUE(A1:A6,FALSE,TRUE)",
        inputs=unique_inputs,
        output_cell="C1",
        description="Return values that occur exactly once",
    )

    # Dynamic array helpers / shape functions (function-catalog backfill)
    _add_case(cases, prefix="choose", tags=["lookup", "CHOOSE"], formula='=CHOOSE(2,"one","two","three")')
    _add_case(
        cases,
        prefix="choosecols",
        tags=["spill", "CHOOSECOLS", "dynarr"],
        formula="=CHOOSECOLS({1,2,3;4,5,6},1,3)",
        output_cell="C1",
    )
    _add_case(
        cases,
        prefix="chooserows",
        tags=["spill", "CHOOSEROWS", "dynarr"],
        formula="=CHOOSEROWS({1,2;3,4;5,6},1,3)",
        output_cell="C1",
    )
    _add_case(cases, prefix="hstack", tags=["spill", "HSTACK", "dynarr"], formula="=HSTACK({1,2},{3,4})", output_cell="C1")
    _add_case(cases, prefix="vstack", tags=["spill", "VSTACK", "dynarr"], formula="=VSTACK({1,2},{3,4})", output_cell="C1")
    _add_case(cases, prefix="take", tags=["spill", "TAKE", "dynarr"], formula="=TAKE({1,2,3;4,5,6},1,2)", output_cell="C1")
    _add_case(cases, prefix="drop", tags=["spill", "DROP", "dynarr"], formula="=DROP({1,2,3;4,5,6},1,1)", output_cell="C1")
    _add_case(cases, prefix="tocol", tags=["spill", "TOCOL", "dynarr"], formula="=TOCOL({1,2;3,4})", output_cell="C1")
    _add_case(cases, prefix="torow", tags=["spill", "TOROW", "dynarr"], formula="=TOROW({1,2;3,4})", output_cell="C1")
    _add_case(cases, prefix="wraprows", tags=["spill", "WRAPROWS", "dynarr"], formula="=WRAPROWS({1,2,3,4,5,6},2)", output_cell="C1")
    _add_case(cases, prefix="wrapcols", tags=["spill", "WRAPCOLS", "dynarr"], formula="=WRAPCOLS({1,2,3,4,5,6},2)", output_cell="C1")
    _add_case(
        cases,
        prefix="expand",
        tags=["spill", "EXPAND", "dynarr"],
        formula="=EXPAND({1,2;3,4},3,3,0)",
        output_cell="C1",
    )
    _add_case(
        cases,
        prefix="sortby",
        tags=["spill", "SORTBY", "dynarr"],
        formula='=SORTBY({"b";"a";"c"},{2;1;3})',
        output_cell="C1",
    )
    _add_case(
        cases,
        prefix="makearray",
        tags=["spill", "MAKEARRAY", "LAMBDA", "dynarr"],
        formula="=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))",
        output_cell="C1",
    )
    _add_case(
        cases,
        prefix="map",
        tags=["spill", "MAP", "LAMBDA", "dynarr"],
        formula="=MAP({1,2,3},LAMBDA(x,x*2))",
        output_cell="C1",
    )
    _add_case(
        cases,
        prefix="reduce",
        tags=["spill", "REDUCE", "LAMBDA", "dynarr"],
        formula="=REDUCE(0,{1,2,3},LAMBDA(acc,x,acc+x))",
        output_cell="C1",
    )
    _add_case(
        cases,
        prefix="scan",
        tags=["spill", "SCAN", "LAMBDA", "dynarr"],
        formula="=SCAN(0,{1,2,3},LAMBDA(acc,x,acc+x))",
        output_cell="C1",
    )
    _add_case(
        cases,
        prefix="byrow",
        tags=["spill", "BYROW", "LAMBDA", "dynarr"],
        formula="=BYROW({1,2;3,4},LAMBDA(r,SUM(r)))",
        output_cell="C1",
    )
    _add_case(
        cases,
        prefix="bycol",
        tags=["spill", "BYCOL", "LAMBDA", "dynarr"],
        formula="=BYCOL({1,2;3,4},LAMBDA(c,SUM(c)))",
        output_cell="C1",
    )
    _add_case(cases, prefix="let", tags=["lambda", "LET"], formula="=LET(a,2,b,a*3,c,b+1,c)")
    _add_case(cases, prefix="lambda", tags=["lambda", "LAMBDA"], formula="=LAMBDA(x,x+1)(2)")
    _add_case(
        cases,
        prefix="isomitted",
        tags=["lambda", "LAMBDA", "ISOMITTED"],
        formula="=LAMBDA(x,y,ISOMITTED(y))(1)",
        description="Missing LAMBDA arguments bind as blank and are detectable via ISOMITTED",
    )

    # ------------------------------------------------------------------
    # Information functions
    # ------------------------------------------------------------------
    _add_case(cases, prefix="isblank", tags=["info", "ISBLANK"], formula="=ISBLANK(A1)")
    _add_case(
        cases,
        prefix="isblank",
        tags=["info", "ISBLANK"],
        formula="=ISBLANK(A1)",
        inputs=[CellInput("A1", "")],
        description="Empty string is not considered blank",
    )

    _add_case(cases, prefix="isnumber", tags=["info", "ISNUMBER"], formula="=ISNUMBER(A1)", inputs=[CellInput("A1", 1)])
    _add_case(cases, prefix="isnumber", tags=["info", "ISNUMBER"], formula="=ISNUMBER(A1)", inputs=[CellInput("A1", "1")])
    _add_case(cases, prefix="istext", tags=["info", "ISTEXT"], formula="=ISTEXT(A1)", inputs=[CellInput("A1", "x")])
    _add_case(cases, prefix="istext", tags=["info", "ISTEXT"], formula="=ISTEXT(A1)", inputs=[CellInput("A1", 1)])
    _add_case(cases, prefix="islogical", tags=["info", "ISLOGICAL"], formula="=ISLOGICAL(A1)", inputs=[CellInput("A1", True)])
    _add_case(cases, prefix="islogical", tags=["info", "ISLOGICAL"], formula="=ISLOGICAL(A1)", inputs=[CellInput("A1", 0)])

    _add_case(cases, prefix="iserror", tags=["info", "ISERROR"], formula="=ISERROR(A1)", inputs=[CellInput("A1", formula="=1/0")])
    _add_case(cases, prefix="iserror", tags=["info", "ISERROR"], formula="=ISERROR(A1)", inputs=[CellInput("A1", 1)])
    _add_case(cases, prefix="iserr", tags=["info", "ISERR"], formula="=ISERR(A1)", inputs=[CellInput("A1", formula="=1/0")])
    _add_case(cases, prefix="iserr", tags=["info", "ISERR"], formula="=ISERR(A1)", inputs=[CellInput("A1", formula="=NA()")])
    _add_case(cases, prefix="isna", tags=["info", "ISNA"], formula="=ISNA(A1)", inputs=[CellInput("A1", formula="=NA()")])
    _add_case(cases, prefix="isna", tags=["info", "ISNA"], formula="=ISNA(A1)", inputs=[CellInput("A1", formula="=1/0")])

    _add_case(cases, prefix="errtype", tags=["info", "ERROR.TYPE"], formula="=ERROR.TYPE(1/0)")
    _add_case(cases, prefix="errtype", tags=["info", "ERROR.TYPE"], formula="=ERROR.TYPE(NA())")
    _add_case(cases, prefix="errtype", tags=["info", "ERROR.TYPE", "error"], formula="=ERROR.TYPE(1)")

    _add_case(cases, prefix="type", tags=["info", "TYPE"], formula="=TYPE(1)")
    _add_case(cases, prefix="type", tags=["info", "TYPE"], formula='=TYPE("x")')
    _add_case(cases, prefix="type", tags=["info", "TYPE"], formula="=TYPE(TRUE)")
    _add_case(cases, prefix="type", tags=["info", "TYPE"], formula="=TYPE(NA())")

    _add_case(cases, prefix="n", tags=["info", "N"], formula="=N(TRUE)")
    _add_case(cases, prefix="n", tags=["info", "N"], formula='=N("x")')
    _add_case(cases, prefix="n", tags=["info", "N"], formula="=N(NA())")

    _add_case(cases, prefix="t", tags=["info", "T"], formula='=T("x")')
    _add_case(cases, prefix="t", tags=["info", "T"], formula="=T(1)")
    _add_case(cases, prefix="t", tags=["info", "T"], formula="=T(NA())")

    # ------------------------------------------------------------------
    # LET / LAMBDA + higher-order array functions (Excel 365+)
    # ------------------------------------------------------------------
    # These are intentionally excluded from `compat_gate.py`'s default include-tag list until
    # we have a pinned Excel dataset for them. Keep tags stable so pinned datasets can be
    # regenerated deterministically on Windows + Excel.
    _add_case(
        cases,
        prefix="lambda_let",
        tags=["lambda", "LET"],
        formula="=LET(x,1,x+1)",
        description="LET basic binding",
    )
    _add_case(
        cases,
        prefix="lambda_let",
        tags=["lambda", "LET"],
        formula="=LET(x,1,y,2,x+y)",
        description="LET multiple bindings",
    )
    _add_case(
        cases,
        prefix="lambda_let",
        tags=["lambda", "LET"],
        formula="=LET(x,1,x,2,x)",
        description="LET shadowing (last binding wins)",
    )

    _add_case(
        cases,
        prefix="lambda_call",
        tags=["lambda", "LAMBDA"],
        formula="=LAMBDA(x,x+1)(5)",
        description="Postfix lambda invocation (call expression)",
    )
    _add_case(
        cases,
        prefix="lambda_call",
        tags=["lambda", "LAMBDA"],
        formula="=LET(inc,LAMBDA(x,x+1),(inc)(5))",
        description="Call a lambda via a LET-bound name (call expression)",
    )

    _add_case(
        cases,
        prefix="lambda_map",
        tags=["lambda", "MAP", "spill"],
        formula="=MAP({1,2,3},LAMBDA(x,x*2))",
        description="MAP over a 1x3 array literal",
    )
    _add_case(
        cases,
        prefix="lambda_reduce",
        tags=["lambda", "REDUCE"],
        formula="=REDUCE(0,{1,2,3},LAMBDA(acc,x,acc+x))",
        description="REDUCE with explicit initial value",
    )
    _add_case(
        cases,
        prefix="lambda_scan",
        tags=["lambda", "SCAN", "spill"],
        formula="=SCAN(0,{1,2,3},LAMBDA(acc,x,acc+x))",
        description="SCAN running sum",
    )
    _add_case(
        cases,
        prefix="lambda_byrow",
        tags=["lambda", "BYROW", "spill"],
        formula="=BYROW({1,2;3,4},LAMBDA(r,SUM(r)))",
        description="BYROW with row-wise SUM",
    )
    _add_case(
        cases,
        prefix="lambda_bycol",
        tags=["lambda", "BYCOL", "spill"],
        formula="=BYCOL({1,2;3,4},LAMBDA(c,SUM(c)))",
        description="BYCOL with column-wise SUM",
    )
    _add_case(
        cases,
        prefix="lambda_makearray",
        tags=["lambda", "MAKEARRAY", "spill"],
        formula="=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))",
        description="MAKEARRAY with index-based values",
    )

    # ------------------------------------------------------------------
    # Explicit error values
    # ------------------------------------------------------------------
    _add_case(cases, prefix="err_div0", tags=["error"], formula="=1/0", inputs=[], output_cell="A1")
    _add_case(cases, prefix="err_na", tags=["error"], formula="=NA()", inputs=[], output_cell="A1")
    _add_case(cases, prefix="err_name", tags=["error"], formula="=NO_SUCH_FUNCTION(1)", inputs=[], output_cell="A1")

    return {
        "schemaVersion": 1,
        "caseSet": "p0-p1-curated-2k",
        "defaultSheet": "Sheet1",
        "cases": cases,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", required=True, help="Output path for cases.json")
    args = parser.parse_args()

    payload = generate_cases()
    _validate_against_function_catalog(payload)

    # Keep the corpus bounded so it remains runnable in real Excel (COM automation) and in CI.
    max_cases = 2000
    cases = payload.get("cases", [])
    if not isinstance(cases, list):
        raise SystemExit("Generated payload.cases must be an array")
    if len(cases) > max_cases:
        raise SystemExit(f"Generated oracle corpus too large: {len(cases)} cases (max {max_cases})")

    # The stable case id should be unique; duplicates indicate accidentally identical cases.
    from collections import Counter

    ids = [c.get("id") for c in cases if isinstance(c, dict)]
    counts = Counter(cid for cid in ids if isinstance(cid, str))
    dup_ids = sorted([cid for cid, n in counts.items() if n > 1])
    if dup_ids:
        dup_preview = ", ".join(dup_ids[:10])
        raise SystemExit(f"Generated oracle corpus contains duplicate case ids: {dup_preview}")

    # Stable JSON formatting for review diffs.
    out_path = args.out
    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, sort_keys=False)
        f.write("\n")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
