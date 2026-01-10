#!/usr/bin/env python3
"""
Compare formula engine results against an Excel oracle dataset.

Input files:
- cases.json: the canonical case corpus (for looking up formula + inputs)
- expected.json: results from tools/excel-oracle/run-excel-oracle.ps1
- actual.json: results from the formula engine under test (same schema as expected)

Output:
- report JSON with mismatches and a summary.
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable


def _load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _index_results(results: Iterable[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    out: dict[str, dict[str, Any]] = {}
    for r in results:
        cid = r.get("caseId")
        if not isinstance(cid, str):
            continue
        out[cid] = r
    return out


def _pretty_input(cell_input: dict[str, Any]) -> dict[str, Any]:
    if "formula" in cell_input:
        return {"cell": cell_input.get("cell"), "formula": cell_input.get("formula")}
    return {"cell": cell_input.get("cell"), "value": cell_input.get("value")}


@dataclass(frozen=True)
class CompareConfig:
    abs_tol: float
    rel_tol: float


def _is_number(value_obj: Any) -> bool:
    return isinstance(value_obj, dict) and value_obj.get("t") == "n" and isinstance(value_obj.get("v"), (int, float))


def _numbers_close(a: float, b: float, cfg: CompareConfig) -> bool:
    # Handle NaN explicitly (Excel doesn't produce NaN, but engines might).
    if math.isnan(a) or math.isnan(b):
        return False
    return math.isclose(a, b, rel_tol=cfg.rel_tol, abs_tol=cfg.abs_tol)


def _compare_value(expected: Any, actual: Any, cfg: CompareConfig) -> tuple[bool, str]:
    if expected == actual:
        return True, "ok"

    if not isinstance(expected, dict) or not isinstance(actual, dict):
        return False, "type-mismatch"

    et = expected.get("t")
    at = actual.get("t")
    if et != at:
        return False, "type-mismatch"

    if et == "n":
        av = float(actual.get("v"))
        ev = float(expected.get("v"))
        return (_numbers_close(ev, av, cfg), "number-mismatch")

    if et in ("s", "b", "e"):
        return (expected.get("v") == actual.get("v"), f"{et}-mismatch")

    if et == "blank":
        return True, "ok"

    if et == "arr":
        erows = expected.get("rows")
        arows = actual.get("rows")
        if not isinstance(erows, list) or not isinstance(arows, list):
            return False, "array-shape-mismatch"
        if len(erows) != len(arows):
            return False, "array-shape-mismatch"
        for r in range(len(erows)):
            if not isinstance(erows[r], list) or not isinstance(arows[r], list):
                return False, "array-shape-mismatch"
            if len(erows[r]) != len(arows[r]):
                return False, "array-shape-mismatch"
            for c in range(len(erows[r])):
                ok, reason = _compare_value(erows[r][c], arows[r][c], cfg)
                if not ok:
                    return False, f"array-mismatch:{reason}"
        return True, "ok"

    return False, "unknown-type"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--cases", required=True, help="Path to cases.json")
    parser.add_argument("--expected", required=True, help="Path to Excel oracle results JSON")
    parser.add_argument("--actual", required=True, help="Path to engine results JSON")
    parser.add_argument("--report", required=True, help="Path to write mismatch report JSON")
    parser.add_argument("--abs-tol", type=float, default=1e-9)
    parser.add_argument("--rel-tol", type=float, default=1e-9)
    parser.add_argument(
        "--max-mismatch-rate",
        type=float,
        default=0.0,
        help="Fail if mismatches / total > this threshold (default 0).",
    )
    args = parser.parse_args()

    cases_path = Path(args.cases)
    expected_path = Path(args.expected)
    actual_path = Path(args.actual)
    report_path = Path(args.report)

    cases = _load_json(cases_path)
    expected = _load_json(expected_path)
    actual = _load_json(actual_path)

    if cases.get("schemaVersion") != 1:
        raise SystemExit(f"Unsupported cases schemaVersion: {cases.get('schemaVersion')}")
    if expected.get("schemaVersion") != 1:
        raise SystemExit(f"Unsupported expected schemaVersion: {expected.get('schemaVersion')}")
    if actual.get("schemaVersion") != 1:
        raise SystemExit(f"Unsupported actual schemaVersion: {actual.get('schemaVersion')}")

    expected_index = _index_results(expected.get("results", []))
    actual_index = _index_results(actual.get("results", []))

    cfg = CompareConfig(abs_tol=args.abs_tol, rel_tol=args.rel_tol)

    mismatches: list[dict[str, Any]] = []

    for case in cases.get("cases", []):
        case_id = case.get("id")
        if not isinstance(case_id, str):
            continue

        exp = expected_index.get(case_id)
        act = actual_index.get(case_id)

        if exp is None:
            mismatches.append(
                {
                    "caseId": case_id,
                    "reason": "missing-expected",
                    "formula": case.get("formula"),
                    "inputs": [_pretty_input(i) for i in case.get("inputs", [])],
                }
            )
            continue

        if act is None:
            mismatches.append(
                {
                    "caseId": case_id,
                    "reason": "missing-actual",
                    "formula": case.get("formula"),
                    "inputs": [_pretty_input(i) for i in case.get("inputs", [])],
                    "expected": exp.get("result"),
                }
            )
            continue

        ok, reason = _compare_value(exp.get("result"), act.get("result"), cfg)
        if not ok:
            mismatches.append(
                {
                    "caseId": case_id,
                    "reason": reason,
                    "formula": case.get("formula"),
                    "inputs": [_pretty_input(i) for i in case.get("inputs", [])],
                    "expected": exp.get("result"),
                    "actual": act.get("result"),
                }
            )

    total = len([c for c in cases.get("cases", []) if isinstance(c.get("id"), str)])
    mismatch_count = len(mismatches)
    mismatch_rate = (mismatch_count / total) if total else 0.0

    report = {
        "schemaVersion": 1,
        "summary": {
            "totalCases": total,
            "mismatches": mismatch_count,
            "mismatchRate": mismatch_rate,
            "maxMismatchRate": args.max_mismatch_rate,
        },
        "expectedSource": expected.get("source"),
        "actualSource": actual.get("source"),
        "mismatches": mismatches,
    }

    report_path.parent.mkdir(parents=True, exist_ok=True)
    with report_path.open("w", encoding="utf-8", newline="\n") as f:
        json.dump(report, f, ensure_ascii=False, indent=2, sort_keys=False)
        f.write("\n")

    # Exit code based on threshold.
    if mismatch_rate > args.max_mismatch_rate:
        sys.stderr.write(
            f"Excel compatibility mismatch rate {mismatch_rate:.4%} exceeded threshold {args.max_mismatch_rate:.4%}\n"
        )
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

