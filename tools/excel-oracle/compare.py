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
import re
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


@dataclass(frozen=True)
class CompareConfig:
    abs_tol: float
    rel_tol: float


def _parse_tag_tolerances(values: list[str], *, flag_name: str) -> dict[str, float]:
    """
    Parse `TAG=FLOAT` pairs into a mapping, taking the maximum for duplicate tags.
    """

    out: dict[str, float] = {}
    for raw in values:
        if not isinstance(raw, str) or "=" not in raw:
            raise SystemExit(
                f"Invalid {flag_name} value {raw!r}. Expected TAG=FLOAT (example: odd_coupon=1e-6)."
            )
        tag, value_str = raw.split("=", 1)
        tag = tag.strip()
        if not tag:
            raise SystemExit(
                f"Invalid {flag_name} value {raw!r}. Tag must be non-empty (example: odd_coupon=1e-6)."
            )
        try:
            value = float(value_str)
        except ValueError:
            raise SystemExit(
                f"Invalid {flag_name} value {raw!r}. {value_str!r} is not a float (example: odd_coupon=1e-6)."
            ) from None
        if not math.isfinite(value) or value < 0.0:
            raise SystemExit(
                f"Invalid {flag_name} value {raw!r}. Tolerance must be a finite, non-negative float."
            )

        prev = out.get(tag)
        if prev is None or value > prev:
            out[tag] = value
    return out


def _effective_cfg_for_tags(
    default: CompareConfig,
    *,
    tags: set[str],
    tag_abs_tol: dict[str, float],
    tag_rel_tol: dict[str, float],
) -> CompareConfig:
    abs_tol = default.abs_tol
    rel_tol = default.rel_tol
    for t in tags:
        v = tag_abs_tol.get(t)
        if v is not None and v > abs_tol:
            abs_tol = v
        v = tag_rel_tol.get(t)
        if v is not None and v > rel_tol:
            rel_tol = v
    return CompareConfig(abs_tol=abs_tol, rel_tol=rel_tol)


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
    parser.add_argument(
        "--max-cases",
        type=int,
        default=0,
        help="Optional cap (after tag filtering): compare only the first N cases (0 = all).",
    )
    parser.add_argument("--abs-tol", type=float, default=1e-9)
    parser.add_argument("--rel-tol", type=float, default=1e-9)
    parser.add_argument(
        "--tag-abs-tol",
        action="append",
        default=[],
        help=(
            "Override numeric abs tolerance for cases that contain a tag. Format TAG=FLOAT "
            "(example: odd_coupon=1e-6). Can be repeated; the maximum across matching tags wins."
        ),
    )
    parser.add_argument(
        "--tag-rel-tol",
        action="append",
        default=[],
        help=(
            "Override numeric rel tolerance for cases that contain a tag. Format TAG=FLOAT "
            "(example: odd_coupon=1e-6). Can be repeated; the maximum across matching tags wins."
        ),
    )
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
    actual_source = actual.get("source")
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

    expected_results = expected.get("results", [])
    if not isinstance(expected_results, list):
        raise SystemExit("Expected dataset 'results' must be an array.")
    expected_count = expected_case_set.get("count") if isinstance(expected_case_set, dict) else None
    if isinstance(expected_count, int) and expected_count != len(expected_results):
        raise SystemExit(
            "Expected dataset caseSet.count does not match results length. "
            f"count={expected_count} results={len(expected_results)}"
        )

    actual_results = actual.get("results", [])
    if not isinstance(actual_results, list):
        raise SystemExit("Actual dataset 'results' must be an array.")
    actual_count = actual_case_set.get("count") if isinstance(actual_case_set, dict) else None
    if isinstance(actual_count, int) and actual_count != len(actual_results):
        raise SystemExit(
            "Actual dataset caseSet.count does not match results length. "
            f"count={actual_count} results={len(actual_results)}"
        )

    # Developer ergonomics: `formula-excel-oracle` is frequently run with tag filters (or `--max-cases`)
    # to keep iteration fast. If the user then runs `compare.py` without the same filters, the report
    # is dominated by "missing-actual" noise and can look like a catastrophic regression.
    #
    # When compare has no filters enabled, sanity-check that the actual dataset appears to cover the
    # full corpus before continuing.
    if (
        not args.include_tag
        and not args.exclude_tag
        and args.max_cases == 0
        and len(actual_results) != len(cases.get("cases", []))
    ):
        raise SystemExit(
            "Actual dataset does not cover the full case corpus. "
            f"cases={len(cases.get('cases', []))} actual_results={len(actual_results)}. "
            "If you generated the engine results with --include-tag/--exclude-tag or --max-cases, "
            "re-run compare.py with the same filters, or regenerate the engine results without filtering."
        )

    expected_index = _index_results(expected_results)
    actual_index = _index_results(actual_results)

    default_cfg = CompareConfig(abs_tol=args.abs_tol, rel_tol=args.rel_tol)
    tag_abs_tol = _parse_tag_tolerances(args.tag_abs_tol, flag_name="--tag-abs-tol")
    tag_rel_tol = _parse_tag_tolerances(args.tag_rel_tol, flag_name="--tag-rel-tol")

    mismatches: list[dict[str, Any]] = []
    reason_counts: dict[str, int] = {}
    tag_totals: dict[str, int] = {}
    tag_fails: dict[str, int] = {}

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

    if args.max_cases and args.max_cases > 0:
        included_cases = included_cases[: args.max_cases]

    for case in included_cases:
        case_id = case["id"]
        tags = case.get("tags", [])
        if not isinstance(tags, list):
            tags = []
        tag_set = {t for t in tags if isinstance(t, str)}

        exp = expected_index.get(case_id)
        act = actual_index.get(case_id)

        mismatch_reason: str | None = None
        if exp is None:
            mismatch_reason = "missing-expected"
            actual_value = act.get("result") if isinstance(act, dict) else None
            mismatches.append(
                {
                    "caseId": case_id,
                    "reason": mismatch_reason,
                    "formula": case.get("formula"),
                    "inputs": [_pretty_input(i) for i in case.get("inputs", [])],
                    "tags": sorted(tag_set),
                    # When the expected dataset is missing a case (common when new deterministic
                    # cases are added to cases.json but the pinned dataset wasn't updated yet),
                    # include the engine-computed value to make patching/regeneration easier.
                    **({"actual": actual_value} if actual_value is not None else {}),
                }
            )
            reason_counts[mismatch_reason] = reason_counts.get(mismatch_reason, 0) + 1

        elif act is None:
            mismatch_reason = "missing-actual"
            mismatches.append(
                {
                    "caseId": case_id,
                    "reason": mismatch_reason,
                    "formula": case.get("formula"),
                    "inputs": [_pretty_input(i) for i in case.get("inputs", [])],
                    "tags": sorted(tag_set),
                    "expected": exp.get("result"),
                }
            )
            reason_counts[mismatch_reason] = reason_counts.get(mismatch_reason, 0) + 1

        else:
            cfg = _effective_cfg_for_tags(
                default_cfg,
                tags=tag_set,
                tag_abs_tol=tag_abs_tol,
                tag_rel_tol=tag_rel_tol,
            )
            ok, reason = _compare_value(exp.get("result"), act.get("result"), cfg)
            if not ok:
                mismatch_reason = reason
                mismatches.append(
                    {
                        "caseId": case_id,
                        "reason": mismatch_reason,
                        "formula": case.get("formula"),
                        "inputs": [_pretty_input(i) for i in case.get("inputs", [])],
                        "tags": sorted(tag_set),
                        "expected": exp.get("result"),
                        "actual": act.get("result"),
                    }
                )
                reason_counts[mismatch_reason] = reason_counts.get(mismatch_reason, 0) + 1

        # Per-tag accounting (a case can contribute to multiple tags).
        if not tag_set:
            tag_set = {"<untagged>"}

        for t in tag_set:
            tag_totals[t] = tag_totals.get(t, 0) + 1
            if mismatch_reason is not None:
                tag_fails[t] = tag_fails.get(t, 0) + 1

    total = len(included_cases)
    mismatch_count = len(mismatches)
    mismatch_rate = (mismatch_count / total) if total else 0.0

    tag_summary: list[dict[str, Any]] = []
    for tag, tot in tag_totals.items():
        fails = tag_fails.get(tag, 0)
        passes = tot - fails
        tag_summary.append(
            {
                "tag": tag,
                "total": tot,
                "passes": passes,
                "mismatches": fails,
                "mismatchRate": (fails / tot) if tot else 0.0,
            }
        )
    tag_summary.sort(key=lambda x: (-x["mismatches"], -x["total"], x["tag"]))

    # Derived aggregates over mismatches: missing functions and error kinds.
    missing_functions: dict[str, int] = {}
    actual_error_kinds: dict[str, int] = {}
    for m in mismatches:
        mismatch_actual = m.get("actual")
        if isinstance(mismatch_actual, dict) and mismatch_actual.get("t") == "e":
            code = mismatch_actual.get("v")
            if isinstance(code, str):
                actual_error_kinds[code] = actual_error_kinds.get(code, 0) + 1
                if code == "#NAME?":
                    for fn in _extract_function_names(m.get("formula")):
                        missing_functions[fn] = missing_functions.get(fn, 0) + 1

    top_missing_functions = [
        {"name": k, "count": v}
        for k, v in sorted(missing_functions.items(), key=lambda kv: (-kv[1], kv[0]))
    ][:20]
    top_actual_error_kinds = [
        {"code": k, "count": v}
        for k, v in sorted(actual_error_kinds.items(), key=lambda kv: (-kv[1], kv[0]))
    ][:20]

    report = {
        "schemaVersion": 1,
        "summary": {
            "totalCases": total,
            "includeTags": sorted(include_tags),
            "excludeTags": sorted(exclude_tags),
            "maxCases": args.max_cases,
            "absTol": args.abs_tol,
            "relTol": args.rel_tol,
            "tagAbsTol": tag_abs_tol,
            "tagRelTol": tag_rel_tol,
            "mismatches": mismatch_count,
            "mismatchRate": mismatch_rate,
            "maxMismatchRate": args.max_mismatch_rate,
            "reasonCounts": dict(sorted(reason_counts.items(), key=lambda kv: (-kv[1], kv[0]))),
            "tagSummary": tag_summary,
            "topMissingFunctions": top_missing_functions,
            "topActualErrorKinds": top_actual_error_kinds,
            "casesSha256": cases_sha,
            # Make reports self-contained: consumers (CI artifacts, local debugging) should be able
            # to see exactly which datasets were compared without having to reconstruct CLI args.
            "casesPath": str(cases_path),
            "expectedPath": str(expected_path),
            "actualPath": str(actual_path),
        },
        "expectedSource": expected.get("source"),
        "actualSource": actual_source,
        "mismatches": mismatches,
    }

    report_path.parent.mkdir(parents=True, exist_ok=True)
    with report_path.open("w", encoding="utf-8", newline="\n") as f:
        json.dump(report, f, ensure_ascii=False, indent=2, sort_keys=False)
        f.write("\n")

    # Human-friendly summary (stdout) for CI/dev ergonomics.
    print(f"Excel compatibility: {total} cases, {mismatch_count} mismatches ({mismatch_rate:.4%})")
    if tag_summary:
        print("")
        print("Tag summary (mismatches/total):")
        # Print tags with failures; if none, print the largest tags for context.
        interesting = [t for t in tag_summary if t["mismatches"] > 0]
        if not interesting:
            interesting = tag_summary[: min(10, len(tag_summary))]
        for row in interesting[: min(25, len(interesting))]:
            print(
                f"  {row['tag']}: {row['mismatches']}/{row['total']} ({row['mismatchRate']:.2%})"
            )

    if top_missing_functions:
        print("")
        print("Top missing functions (mismatches where actual is #NAME?):")
        for row in top_missing_functions[: min(10, len(top_missing_functions))]:
            print(f"  {row['name']}: {row['count']}")

    if top_actual_error_kinds:
        print("")
        print("Top actual error kinds (in mismatches):")
        for row in top_actual_error_kinds[: min(10, len(top_actual_error_kinds))]:
            print(f"  {row['code']}: {row['count']}")

    # Exit code based on threshold.
    if mismatch_rate > args.max_mismatch_rate:
        sys.stderr.write(
            f"Excel compatibility mismatch rate {mismatch_rate:.4%} exceeded threshold {args.max_mismatch_rate:.4%}\n"
        )
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
