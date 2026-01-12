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
import json
import re
from pathlib import Path
from typing import Any, Iterable

from case_generators import (
    arith,
    coercion,
    database,
    date_time,
    engineering,
    errors,
    financial,
    info,
    lambda_cases,
    logical,
    lookup,
    math as math_cases,
    spill,
    statistical,
    text,
)


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

    # NOTE: Module invocation order is part of output stability; keep it deterministic.
    arith.generate(cases, add_case=_add_case, CellInput=CellInput)
    math_cases.generate(cases, add_case=_add_case, CellInput=CellInput)
    engineering.generate(cases, add_case=_add_case, CellInput=CellInput)
    statistical.generate(cases, add_case=_add_case, CellInput=CellInput, excel_serial_1900=_excel_serial_1900)

    # Regression / forecasting (Excel BIFF FTAB: LINEST, LOGEST, TREND, GROWTH).
    #
    # These functions are deterministic but easy to accidentally omit because they return arrays
    # and are less commonly used in "simple" spreadsheets. Include a few small, stable cases so
    # they show up in corpus function coverage.
    _add_case(
        cases,
        prefix="stat_linest",
        tags=["statistical", "LINEST"],
        formula="=LINEST({1;2;3},{1;2;3})",
    )
    _add_case(
        cases,
        prefix="stat_logest",
        tags=["statistical", "LOGEST"],
        formula="=LOGEST({2;6;18},{0;1;2})",
    )
    _add_case(
        cases,
        prefix="stat_trend",
        tags=["statistical", "TREND"],
        formula="=TREND({1;2;3},{1;2;3},{4;5})",
    )
    _add_case(
        cases,
        prefix="stat_growth",
        tags=["statistical", "GROWTH"],
        formula="=GROWTH({2;6;18},{0;1;2},{3;4})",
    )

    logical.generate(cases, add_case=_add_case, CellInput=CellInput)
    coercion.generate(cases, add_case=_add_case, CellInput=CellInput)
    text.generate(cases, add_case=_add_case, CellInput=CellInput, excel_serial_1900=_excel_serial_1900)
    date_time.generate(cases, add_case=_add_case, CellInput=CellInput)
    lookup.generate(cases, add_case=_add_case, CellInput=CellInput)
    database.generate(cases, add_case=_add_case, CellInput=CellInput)
    financial.generate(cases, add_case=_add_case, CellInput=CellInput, excel_serial_1900=_excel_serial_1900)
    spill.generate(cases, add_case=_add_case, CellInput=CellInput)
    info.generate(cases, add_case=_add_case, CellInput=CellInput)
    lambda_cases.generate(cases, add_case=_add_case, CellInput=CellInput)
    errors.generate(cases, add_case=_add_case, CellInput=CellInput)

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
