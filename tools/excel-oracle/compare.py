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
import hashlib
import json
import math
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable


def _load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


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
    return (
        isinstance(value_obj, dict)
        and value_obj.get("t") == "n"
        and isinstance(value_obj.get("v"), (int, float))
    )


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
    parser.add_argument(
        "--include-tag",
        action="append",
        default=[],
        help="Only include cases that contain this tag (can be repeated).",
    )
    parser.add_argument(
        "--exclude-tag",
        action="append",
        default=[],
        help="Exclude cases that contain this tag (can be repeated).",
    )
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

    expected_source = expected.get("source", {})
    if isinstance(expected_source, dict) and expected_source.get("kind") != "excel":
        raise SystemExit(
            "Expected dataset must be produced by real Excel (source.kind == 'excel'). "
            f"Got: {expected_source.get('kind')!r}"
        )

    cases_sha = _sha256_file(cases_path)
    expected_case_set = expected.get("caseSet")
    actual_case_set = actual.get("caseSet")
    expected_sha = expected_case_set.get("sha256") if isinstance(expected_case_set, dict) else None
    actual_sha = actual_case_set.get("sha256") if isinstance(actual_case_set, dict) else None

    if isinstance(expected_sha, str) and expected_sha.lower() != cases_sha.lower():
        raise SystemExit(
            "Expected dataset caseSet.sha256 does not match cases.json. "
            f"expected={expected_sha} cases={cases_sha}"
        )

    if isinstance(actual_sha, str) and actual_sha.lower() != cases_sha.lower():
        raise SystemExit(
            "Actual dataset caseSet.sha256 does not match cases.json. "
            f"actual={actual_sha} cases={cases_sha}"
        )

    expected_index = _index_results(expected.get("results", []))
    actual_index = _index_results(actual.get("results", []))

    cfg = CompareConfig(abs_tol=args.abs_tol, rel_tol=args.rel_tol)

    mismatches: list[dict[str, Any]] = []
    reason_counts: dict[str, int] = {}

    include_tags = set(args.include_tag)
    exclude_tags = set(args.exclude_tag)

    included_cases: list[dict[str, Any]] = []
    for case in cases.get("cases", []):
        case_id = case.get("id")
        if not isinstance(case_id, str):
            continue

        tags = case.get("tags", [])
        if not isinstance(tags, list):
            tags = []
        tag_set = {t for t in tags if isinstance(t, str)}

        if include_tags and not (include_tags & tag_set):
            continue
        if exclude_tags and (exclude_tags & tag_set):
            continue

        included_cases.append(case)

    for case in included_cases:
        case_id = case["id"]

        exp = expected_index.get(case_id)
        act = actual_index.get(case_id)

        if exp is None:
            reason = "missing-expected"
            mismatches.append(
                {
                    "caseId": case_id,
                    "reason": reason,
                    "formula": case.get("formula"),
                    "inputs": [_pretty_input(i) for i in case.get("inputs", [])],
                }
            )
            reason_counts[reason] = reason_counts.get(reason, 0) + 1
            continue

        if act is None:
            reason = "missing-actual"
            mismatches.append(
                {
                    "caseId": case_id,
                    "reason": reason,
                    "formula": case.get("formula"),
                    "inputs": [_pretty_input(i) for i in case.get("inputs", [])],
                    "expected": exp.get("result"),
                }
            )
            reason_counts[reason] = reason_counts.get(reason, 0) + 1
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
            reason_counts[reason] = reason_counts.get(reason, 0) + 1

    total = len(included_cases)
    mismatch_count = len(mismatches)
    mismatch_rate = (mismatch_count / total) if total else 0.0

    report = {
        "schemaVersion": 1,
        "summary": {
            "totalCases": total,
            "includeTags": sorted(include_tags),
            "excludeTags": sorted(exclude_tags),
            "mismatches": mismatch_count,
            "mismatchRate": mismatch_rate,
            "maxMismatchRate": args.max_mismatch_rate,
            "reasonCounts": dict(sorted(reason_counts.items(), key=lambda kv: (-kv[1], kv[0]))),
            "casesSha256": cases_sha,
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
