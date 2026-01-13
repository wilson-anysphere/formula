#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
from typing import Any

from .util import ensure_dir, github_commit_sha, github_run_url, utc_now_iso, write_json


def _status(value: Any) -> str:
    if value is True:
        return "PASS"
    if value is False:
        return "FAIL"
    return "SKIP"


def _rate(passed: int, total: int) -> float:
    if total == 0:
        return 0.0
    return passed / total


def _attempted(value: Any) -> bool:
    # `*_ok` fields in triage results use `True`/`False`/`None` to mean PASS/FAIL/SKIP.
    # Treat booleans as "attempted" and `None`/missing as skipped.
    return value is True or value is False


def _load_reports(reports_dir: Path) -> list[dict[str, Any]]:
    reports: list[dict[str, Any]] = []
    for path in sorted(reports_dir.glob("*.json")):
        reports.append(json.loads(path.read_text(encoding="utf-8")))
    return reports


def _markdown_summary(summary: dict[str, Any], reports: list[dict[str, Any]]) -> str:
    counts = summary["counts"]
    rates = summary["rates"]
    lines: list[str] = []
    lines.append("# Compatibility corpus scorecard")
    lines.append("")
    lines.append(f"- Timestamp: `{summary['timestamp']}`")
    if summary.get("commit"):
        lines.append(f"- Commit: `{summary['commit']}`")
    if summary.get("run_url"):
        lines.append(f"- Run: {summary['run_url']}")
    lines.append("")
    lines.append("## Overall")
    lines.append("")
    lines.append(f"- Total workbooks: **{counts['total']}**")
    lines.append(
        f"- Open: **{counts['open_ok']} / {counts['total']}** ({rates['open']:.1%})"
    )

    for key, label in [("calculate", "Calculate"), ("render", "Render")]:
        ok = int(counts.get(f"{key}_ok", 0))
        attempted = int(counts.get(f"{key}_attempted", 0))
        skipped = int(counts.get("total", 0)) - attempted
        rate = rates.get(key)
        if attempted <= 0 or rate is None:
            extra = "SKIP"
        else:
            extra = f"{float(rate):.1%}"
        if skipped > 0:
            extra = f"{extra}, {skipped} skipped"
        lines.append(f"- {label}: **{ok} / {attempted}** ({extra})")

    lines.append(
        f"- Round-trip: **{counts['round_trip_ok']} / {counts['total']}** ({rates['round_trip']:.1%})"
    )
    diff_totals = summary.get("diff_totals") or {}
    if diff_totals:
        lines.append(
            f"- Diff totals (critical/warn/info): **{diff_totals.get('critical', 0)} / {diff_totals.get('warning', 0)} / {diff_totals.get('info', 0)}**"
        )
    lines.append("")
    lines.append("## Per-workbook")
    lines.append("")
    lines.append("| Workbook | Open | Calculate | Render | Round-trip | Diff (C/W/I) | Failure category |")
    lines.append("|---|---:|---:|---:|---:|---:|---|")
    for r in reports:
        res = r.get("result", {})
        diff_cell = ""
        if any(k in res for k in ("diff_critical_count", "diff_warning_count", "diff_info_count")):
            diff_cell = (
                f"{res.get('diff_critical_count', 0)}/"
                f"{res.get('diff_warning_count', 0)}/"
                f"{res.get('diff_info_count', 0)}"
            )
        lines.append(
            "| "
            + " | ".join(
                [
                    r.get("display_name", "?"),
                    _status(res.get("open_ok")),
                    _status(res.get("calculate_ok")),
                    _status(res.get("render_ok")),
                    _status(res.get("round_trip_ok")),
                    diff_cell,
                    r.get("failure_category", ""),
                ]
            )
            + " |"
        )
    lines.append("")

    failures = summary.get("failures_by_category", {})
    if failures:
        lines.append("## Failures by category")
        lines.append("")
        lines.append("| Category | Count |")
        lines.append("|---|---:|")
        for k, v in sorted(failures.items(), key=lambda kv: (-kv[1], kv[0])):
            lines.append(f"| {k} | {v} |")
        lines.append("")

    top_functions = summary.get("top_functions_in_failures", [])
    if top_functions:
        lines.append("## Top functions in failing workbooks")
        lines.append("")
        lines.append("| Function | Count |")
        lines.append("|---|---:|")
        for row in top_functions[:20]:
            lines.append(f"| {row['function']} | {row['count']} |")
        lines.append("")

    top_features = summary.get("top_features_in_failures", [])
    if top_features:
        lines.append("## Top features in failing workbooks")
        lines.append("")
        lines.append("| Feature | Count |")
        lines.append("|---|---:|")
        for row in top_features[:20]:
            lines.append(f"| {row['feature']} | {row['count']} |")
        lines.append("")

    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate corpus compatibility dashboard.")
    parser.add_argument("--triage-dir", type=Path, required=True)
    parser.add_argument("--out-dir", type=Path, help="Defaults to --triage-dir")
    args = parser.parse_args()

    triage_dir = args.triage_dir
    out_dir = args.out_dir or triage_dir
    ensure_dir(out_dir)

    reports_dir = triage_dir / "reports"
    reports = _load_reports(reports_dir)

    total = len(reports)
    open_ok = sum(1 for r in reports if r.get("result", {}).get("open_ok") is True)
    calc_ok = sum(
        1 for r in reports if r.get("result", {}).get("calculate_ok") is True
    )
    render_ok = sum(1 for r in reports if r.get("result", {}).get("render_ok") is True)
    rt_ok = sum(1 for r in reports if r.get("result", {}).get("round_trip_ok") is True)

    calc_attempted = sum(
        1 for r in reports if _attempted(r.get("result", {}).get("calculate_ok"))
    )
    render_attempted = sum(
        1 for r in reports if _attempted(r.get("result", {}).get("render_ok"))
    )

    failures_by_category: Counter[str] = Counter()
    failing_function_counts: Counter[str] = Counter()
    failing_feature_counts: Counter[str] = Counter()
    diff_totals: Counter[str] = Counter()

    for r in reports:
        res = r.get("result", {})
        failed = any(
            res.get(k) is False
            for k in ("open_ok", "calculate_ok", "render_ok", "round_trip_ok")
        )
        if failed:
            failures_by_category[r.get("failure_category", "unknown")] += 1
            for fn, cnt in (r.get("functions") or {}).items():
                failing_function_counts[fn] += int(cnt)
            for feat, enabled in (r.get("features") or {}).items():
                if enabled is True:
                    failing_feature_counts[feat] += 1

        for key, out_key in [
            ("diff_critical_count", "critical"),
            ("diff_warning_count", "warning"),
            ("diff_info_count", "info"),
        ]:
            val = res.get(key)
            if isinstance(val, int):
                diff_totals[out_key] += val

    summary: dict[str, Any] = {
        "timestamp": utc_now_iso(),
        "commit": github_commit_sha(),
        "run_url": github_run_url(),
        "counts": {
            "total": total,
            "open_ok": open_ok,
            "calculate_ok": calc_ok,
            "calculate_attempted": calc_attempted,
            "render_ok": render_ok,
            "render_attempted": render_attempted,
            "round_trip_ok": rt_ok,
        },
        "rates": {
            "open": _rate(open_ok, total),
            # Calculate/render steps are optional. When disabled, their results are `SKIP` and
            # should not be counted as failures.
            "calculate": (calc_ok / calc_attempted) if calc_attempted else None,
            "render": (render_ok / render_attempted) if render_attempted else None,
            "round_trip": _rate(rt_ok, total),
        },
        "failures_by_category": dict(failures_by_category),
        "diff_totals": dict(diff_totals),
        "top_functions_in_failures": [
            {"function": fn, "count": cnt}
            for fn, cnt in failing_function_counts.most_common(50)
        ],
        "top_features_in_failures": [
            {"feature": feat, "count": cnt}
            for feat, cnt in failing_feature_counts.most_common(50)
        ],
    }

    write_json(out_dir / "summary.json", summary)
    (out_dir / "summary.md").write_text(
        _markdown_summary(summary, reports), encoding="utf-8"
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
