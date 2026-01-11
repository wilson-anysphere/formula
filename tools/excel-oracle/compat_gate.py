#!/usr/bin/env python3
"""
End-to-end Excel-oracle compatibility gate for the Rust formula engine.

This is intentionally lightweight and CI-friendly:

  1) Run `crates/formula-excel-oracle` to produce engine-results.json
  2) Compare against a pinned Excel dataset via `tools/excel-oracle/compare.py`

The compare step emits `tests/compatibility/excel-oracle/reports/mismatch-report.json`
and exits non-zero if the mismatch rate exceeds the configured threshold.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path


DEFAULT_INCLUDE_TAGS = [
    # Keep CI bounded to a small, high-signal slice of the corpus.
    "add",
    "sub",
    "mul",
    "div",
    "cmp",
    "SUM",
    "IF",
    "IFERROR",
    "error",
    # Minimal spill coverage (range reference + a couple of array functions).
    "range",
    "TRANSPOSE",
    "SEQUENCE",
]


def _default_expected_dataset() -> Path:
    versioned_dir = Path("tests/compatibility/excel-oracle/datasets/versioned")
    candidates = sorted(p for p in versioned_dir.glob("*.json") if p.is_file())
    if candidates:
        return candidates[-1]

    pinned = Path("tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json")
    if pinned.is_file():
        return pinned

    raise SystemExit(
        "No pinned Excel oracle dataset found. Expected either:\n"
        "  - tests/compatibility/excel-oracle/datasets/versioned/*.json\n"
        "  - tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json\n"
        "\n"
        "See tests/compatibility/excel-oracle/README.md for how to generate/pin datasets."
    )


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument(
        "--cases",
        default="tests/compatibility/excel-oracle/cases.json",
        help="Path to cases.json (default: %(default)s)",
    )
    p.add_argument(
        "--expected",
        default="",
        help="Path to pinned Excel results JSON. Defaults to the newest file in "
        "tests/compatibility/excel-oracle/datasets/versioned/ if present, else "
        "tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json.",
    )
    p.add_argument(
        "--actual",
        default="tests/compatibility/excel-oracle/datasets/engine-results.json",
        help="Where to write engine results JSON (default: %(default)s)",
    )
    p.add_argument(
        "--report",
        default="tests/compatibility/excel-oracle/reports/mismatch-report.json",
        help="Where to write mismatch report JSON (default: %(default)s)",
    )
    p.add_argument(
        "--report-md",
        default="tests/compatibility/excel-oracle/reports/summary.md",
        help="Where to write a human-readable markdown summary (default: %(default)s)",
    )
    p.add_argument(
        "--max-cases",
        type=int,
        default=0,
        help="Optional cap (after tag filtering): evaluate/compare only first N cases (0 = all).",
    )
    p.add_argument(
        "--include-tag",
        action="append",
        default=[],
        help="Only include cases containing this tag (can be repeated). Defaults to a curated set.",
    )
    p.add_argument(
        "--exclude-tag",
        action="append",
        default=[],
        help="Exclude cases containing this tag (can be repeated).",
    )
    p.add_argument("--abs-tol", type=float, default=1e-9)
    p.add_argument("--rel-tol", type=float, default=1e-9)
    p.add_argument(
        "--max-mismatch-rate",
        type=float,
        default=0.0,
        help="Fail if mismatches / total exceeds this threshold (default 0).",
    )
    args = p.parse_args()

    cases_path = Path(args.cases)
    expected_path = Path(args.expected) if args.expected else _default_expected_dataset()
    actual_path = Path(args.actual)
    report_path = Path(args.report)

    include_tags = args.include_tag or list(DEFAULT_INCLUDE_TAGS)
    exclude_tags = args.exclude_tag

    engine_cmd = [
        "cargo",
        "run",
        "-p",
        "formula-excel-oracle",
        "--quiet",
        "--locked",
        "--",
        "--cases",
        str(cases_path),
        "--out",
        str(actual_path),
    ]
    if args.max_cases and args.max_cases > 0:
        engine_cmd += ["--max-cases", str(args.max_cases)]
    for t in include_tags:
        engine_cmd += ["--include-tag", t]
    for t in exclude_tags:
        engine_cmd += ["--exclude-tag", t]

    subprocess.run(engine_cmd, check=True)

    compare_cmd = [
        sys.executable,
        "tools/excel-oracle/compare.py",
        "--cases",
        str(cases_path),
        "--expected",
        str(expected_path),
        "--actual",
        str(actual_path),
        "--report",
        str(report_path),
        "--max-mismatch-rate",
        str(args.max_mismatch_rate),
        "--abs-tol",
        str(args.abs_tol),
        "--rel-tol",
        str(args.rel_tol),
    ]
    if args.max_cases and args.max_cases > 0:
        compare_cmd += ["--max-cases", str(args.max_cases)]
    for t in include_tags:
        compare_cmd += ["--include-tag", t]
    for t in exclude_tags:
        compare_cmd += ["--exclude-tag", t]

    proc = subprocess.run(compare_cmd)

    # Produce a markdown summary alongside the JSON report for easy viewing in CI.
    try:
        report_payload = json.loads(report_path.read_text(encoding="utf-8"))
        summary = report_payload.get("summary", {}) if isinstance(report_payload, dict) else {}
        md_path = Path(args.report_md)

        lines: list[str] = []
        lines.append("# Excel oracle compatibility report")
        lines.append("")
        lines.append(f"* Total cases: {summary.get('totalCases')}")
        lines.append(f"* Mismatches: {summary.get('mismatches')}")
        lines.append(f"* Mismatch rate: {summary.get('mismatchRate')}")
        lines.append(f"* Max mismatch rate: {summary.get('maxMismatchRate')}")
        lines.append("")

        tag_summary = summary.get("tagSummary")
        if isinstance(tag_summary, list) and tag_summary:
            lines.append("## Tag summary")
            lines.append("")
            lines.append("| Tag | Passes | Mismatches | Total | Mismatch rate |")
            lines.append("| --- | ---: | ---: | ---: | ---: |")
            for row in tag_summary[:50]:
                if not isinstance(row, dict):
                    continue
                tag = row.get("tag")
                passes = row.get("passes")
                mismatches = row.get("mismatches")
                total = row.get("total")
                rate = row.get("mismatchRate")
                lines.append(f"| {tag} | {passes} | {mismatches} | {total} | {rate:.4%} |")
            lines.append("")

        top_missing = summary.get("topMissingFunctions")
        if isinstance(top_missing, list) and top_missing:
            lines.append("## Top missing functions")
            lines.append("")
            for row in top_missing[:20]:
                if isinstance(row, dict) and "name" in row and "count" in row:
                    lines.append(f"* `{row['name']}`: {row['count']}")
            lines.append("")

        top_errors = summary.get("topActualErrorKinds")
        if isinstance(top_errors, list) and top_errors:
            lines.append("## Top actual error kinds (in mismatches)")
            lines.append("")
            for row in top_errors[:20]:
                if isinstance(row, dict) and "code" in row and "count" in row:
                    lines.append(f"* `{row['code']}`: {row['count']}")
            lines.append("")

        mismatches = report_payload.get("mismatches") if isinstance(report_payload, dict) else None
        if isinstance(mismatches, list) and mismatches:
            lines.append("## Sample mismatches")
            lines.append("")
            for m in mismatches[:10]:
                if not isinstance(m, dict):
                    continue
                lines.append(f"* `{m.get('caseId')}` `{m.get('reason')}` `{m.get('formula')}`")
            lines.append("")

        md_path.parent.mkdir(parents=True, exist_ok=True)
        md_path.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")
    except Exception:
        # Don't fail the gate if the summary couldn't be generated (the compare step already
        # enforces correctness via exit code + JSON report).
        pass

    return proc.returncode


if __name__ == "__main__":
    raise SystemExit(main())
