#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import math
import statistics
import sys
from collections import Counter
from pathlib import Path
from typing import Any

from .triage import _PRIVACY_PRIVATE, _PRIVACY_PUBLIC, _redact_run_url, infer_round_trip_failure_kind
from .util import ensure_dir, github_commit_sha, github_run_url, utc_now_iso, write_json


TIMING_STEPS: tuple[str, ...] = ("load", "round_trip", "diff", "recalc", "render")


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


def _mean(values: list[int]) -> float | None:
    if not values:
        return None
    return sum(values) / len(values)


def _median(values: list[int]) -> float | None:
    if not values:
        return None
    values_sorted = sorted(values)
    mid = len(values_sorted) // 2
    if len(values_sorted) % 2 == 1:
        return float(values_sorted[mid])
    return (values_sorted[mid - 1] + values_sorted[mid]) / 2.0


def _load_reports(reports_dir: Path) -> list[dict[str, Any]]:
    """Load triage reports from `reports_dir`.

    If `index.json` is present in the parent directory, prefer its ordered report list so the
    per-workbook table is stable and matches the original triage input ordering (rather than being
    an artifact of report filename sorting).
    """

    triage_dir = reports_dir.parent
    index_path = triage_dir / "index.json"

    if index_path.exists():
        try:
            index = json.loads(index_path.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001 (tooling)
            index = None
        if isinstance(index, dict):
            entries = index.get("reports")
            if isinstance(entries, list):
                ordered: list[dict[str, Any]] = []
                for entry in entries:
                    if not isinstance(entry, dict):
                        continue
                    filename = entry.get("file")
                    if not isinstance(filename, str) or not filename:
                        # Back-compat: older indices only recorded the report id, and the report
                        # file was named `<id>.json`.
                        report_id = entry.get("id")
                        if isinstance(report_id, str) and report_id:
                            filename = f"{report_id}.json"
                    if not isinstance(filename, str) or not filename:
                        continue
                    path = reports_dir / filename
                    if not path.exists():
                        continue
                    try:
                        ordered.append(json.loads(path.read_text(encoding="utf-8")))
                    except Exception:  # noqa: BLE001 (tooling)
                        continue
                if ordered:
                    return ordered

    reports: list[dict[str, Any]] = []
    for path in sorted(reports_dir.glob("*.json")):
        reports.append(json.loads(path.read_text(encoding="utf-8")))
    return reports


def _percentile(sorted_values: list[float], p: float) -> float:
    """Return the `p` percentile for pre-sorted values.

    Uses linear interpolation between points (matches common spreadsheet percentile behavior and
    produces stable results for small sample sizes).
    """

    if not sorted_values:
        raise ValueError("cannot compute percentile of empty data")
    if p <= 0.0:
        return sorted_values[0]
    if p >= 1.0:
        return sorted_values[-1]
    if len(sorted_values) == 1:
        return sorted_values[0]

    # Linear interpolation between the surrounding ranks.
    idx = p * (len(sorted_values) - 1)
    lo = int(math.floor(idx))
    hi = int(math.ceil(idx))
    if lo == hi:
        return sorted_values[lo]
    frac = idx - lo
    return (sorted_values[lo] * (1.0 - frac)) + (sorted_values[hi] * frac)


def _round_trip_size_overhead(reports: list[dict[str, Any]]) -> dict[str, Any]:
    ratios: list[float] = []
    for r in reports:
        input_size = r.get("size_bytes")
        if not isinstance(input_size, int) or input_size <= 0:
            continue

        steps = r.get("steps") or {}
        if not isinstance(steps, dict):
            continue
        rt_step = steps.get("round_trip") or {}
        if not isinstance(rt_step, dict):
            continue
        if rt_step.get("status") != "ok":
            continue

        details = rt_step.get("details") or {}
        if not isinstance(details, dict):
            continue
        output_size = details.get("output_size_bytes")
        if not isinstance(output_size, int):
            continue

        ratios.append(output_size / input_size)

    ratios.sort()
    count = len(ratios)
    if count == 0:
        return {
            "count": 0,
            "mean": None,
            "p50": None,
            "p90": None,
            "max": None,
            "count_over_1_05": 0,
            "count_over_1_10": 0,
        }

    return {
        "count": count,
        "mean": statistics.fmean(ratios),
        "p50": _percentile(ratios, 0.50),
        "p90": _percentile(ratios, 0.90),
        "max": ratios[-1],
        "count_over_1_05": sum(1 for r in ratios if r > 1.05),
        "count_over_1_10": sum(1 for r in ratios if r > 1.10),
    }


def _part_change_ratio_summary(reports: list[dict[str, Any]]) -> dict[str, Any]:
    """Aggregate privacy-safe part-level diff ratios from triage reports.

    These ratios are stable across runs and do not leak workbook content: they only count whether
    a part changed, not what changed.
    """

    ratios: list[float] = []
    critical_ratios: list[float] = []

    for r in reports:
        diff_step = (r.get("steps") or {}).get("diff") or {}
        if not isinstance(diff_step, dict):
            continue
        if diff_step.get("status") != "ok":
            continue

        details = diff_step.get("details") or {}
        if not isinstance(details, dict):
            continue
        part_stats = details.get("part_stats") or {}
        if not isinstance(part_stats, dict):
            continue

        parts_total = part_stats.get("parts_total")
        parts_changed = part_stats.get("parts_changed")
        parts_changed_critical = part_stats.get("parts_changed_critical")
        if not (
            isinstance(parts_total, int)
            and isinstance(parts_changed, int)
            and isinstance(parts_changed_critical, int)
        ):
            continue
        if parts_total <= 0:
            continue

        ratios.append(parts_changed / parts_total)
        critical_ratios.append(parts_changed_critical / parts_total)

    def _summarize(values: list[float]) -> dict[str, Any]:
        values.sort()
        if not values:
            return {"count": 0, "mean": None, "p50": None, "p90": None, "max": None}
        return {
            "count": len(values),
            "mean": statistics.fmean(values),
            "p50": _percentile(values, 0.50),
            "p90": _percentile(values, 0.90),
            "max": values[-1],
        }

    return {
        "part_change_ratio": _summarize(ratios),
        "part_change_ratio_critical": _summarize(critical_ratios),
    }

TREND_MAX_ENTRIES = 90


def _trend_entry(summary: dict[str, Any]) -> dict[str, Any]:
    """Return a compact, machine-readable snapshot for trend tracking."""

    counts = summary.get("counts") or {}
    rates = summary.get("rates") or {}

    diff_totals = dict(summary.get("diff_totals") or {})
    diff_totals.setdefault("critical", 0)
    diff_totals.setdefault("warning", 0)
    diff_totals.setdefault("info", 0)
    diff_totals["total"] = (
        int(diff_totals.get("critical") or 0)
        + int(diff_totals.get("warning") or 0)
        + int(diff_totals.get("info") or 0)
    )

    overhead = summary.get("round_trip_size_overhead") or {}
    if not isinstance(overhead, dict):
        overhead = {}

    part_change_ratio = summary.get("part_change_ratio") or {}
    if not isinstance(part_change_ratio, dict):
        part_change_ratio = {}
    part_change_ratio_critical = summary.get("part_change_ratio_critical") or {}
    if not isinstance(part_change_ratio_critical, dict):
        part_change_ratio_critical = {}

    timings = summary.get("timings") or {}
    if not isinstance(timings, dict):
        timings = {}

    def _timing_val(step: str, key: str) -> float | None:
        row = timings.get(step) if isinstance(timings, dict) else None
        if not isinstance(row, dict):
            return None
        val = row.get(key)
        if isinstance(val, bool):
            return None
        if isinstance(val, (int, float)):
            return float(val)
        return None

    entry: dict[str, Any] = {
        "timestamp": summary.get("timestamp"),
        "commit": summary.get("commit"),
        "run_url": summary.get("run_url"),
        "total": int(counts.get("total") or 0),
        "open_ok": int(counts.get("open_ok") or 0),
        "round_trip_ok": int(counts.get("round_trip_ok") or 0),
        "open_rate": float(rates.get("open") or 0.0),
        "round_trip_rate": float(rates.get("round_trip") or 0.0),
        # Optional checks: rate among attempted workbooks only.
        "calc_ok": int(counts.get("calculate_ok") or 0),
        "calc_attempted": int(counts.get("calculate_attempted") or 0),
        "calc_rate": rates.get("calculate"),
        "render_ok": int(counts.get("render_ok") or 0),
        "render_attempted": int(counts.get("render_attempted") or 0),
        "render_rate": rates.get("render"),
        "diff_totals": diff_totals,
        "failures_by_category": summary.get("failures_by_category") or {},
        # Optional higher-signal breakdown for round-trip failures.
        "failures_by_round_trip_failure_kind": summary.get(
            "failures_by_round_trip_failure_kind"
        )
        or {},
        # Size ratio: output_size / input_size for successful round-trips.
        "size_overhead_mean": overhead.get("mean"),
        "size_overhead_p50": overhead.get("p50"),
        "size_overhead_p90": overhead.get("p90"),
        "size_overhead_samples": int(overhead.get("count") or 0),
        # Fraction of package parts that changed (any severity / critical-only).
        "part_change_ratio_p90": part_change_ratio.get("p90"),
        "part_change_ratio_critical_p90": part_change_ratio_critical.get("p90"),
    }

    def _top_list(
        key: str,
        *,
        list_key: str,
        max_entries: int = 5,
    ) -> None:
        raw = summary.get(key)
        if not isinstance(raw, list):
            return
        out: list[dict[str, Any]] = []
        for row in raw:
            if not isinstance(row, dict):
                continue
            name = row.get(list_key)
            count = row.get("count")
            if not isinstance(name, str) or not name:
                continue
            if isinstance(count, bool) or not isinstance(count, int):
                continue
            out.append({list_key: name, "count": int(count)})
            if len(out) >= max_entries:
                break
        if out:
            entry[key] = out

    # Diff part/group breakdowns: keep only a small top-N in trend entries so the file stays
    # compact even as the summary expands.
    _top_list("top_diff_parts_critical", list_key="part")
    _top_list("top_diff_parts_total", list_key="part")
    _top_list("top_diff_part_groups_critical", list_key="group")
    _top_list("top_diff_part_groups_total", list_key="group")

    # Perf trend signals (optional; only present when dashboard has timings data).
    for key, step, stat_key in [
        ("load_p50_ms", "load", "p50_ms"),
        ("load_p90_ms", "load", "p90_ms"),
        ("round_trip_p50_ms", "round_trip", "p50_ms"),
        ("round_trip_p90_ms", "round_trip", "p90_ms"),
    ]:
        v = _timing_val(step, stat_key)
        if v is not None:
            entry[key] = v

    # Keep entries compact by eliding empty metadata keys.
    if not entry.get("commit"):
        entry.pop("commit", None)
    if not entry.get("run_url"):
        entry.pop("run_url", None)
    return entry


def _append_trend_file(
    trend_path: Path,
    *,
    summary: dict[str, Any],
    max_entries: int = TREND_MAX_ENTRIES,
) -> tuple[list[dict[str, Any]], dict[str, Any] | None]:
    """Append a trend entry for this run and write back the updated JSON list.

    Returns: (updated_entries, previous_entry)
    """

    prev: dict[str, Any] | None = None
    entries: list[dict[str, Any]] = []
    if trend_path.exists():
        raw_text = trend_path.read_text(encoding="utf-8").strip()
        try:
            raw = json.loads(raw_text or "[]")
        except json.JSONDecodeError:
            # Trend files are persisted across scheduled runs via cache. Be resilient to partial
            # writes/corruption so one bad cache entry doesn't break the scheduled job.
            print(
                f"warning: trend file contained invalid JSON; overwriting: {trend_path}",
                file=sys.stderr,
            )
            raw = []
        if not isinstance(raw, list):
            raise ValueError(f"trend file must be a JSON list: {trend_path}")
        entries = [e for e in raw if isinstance(e, dict)]
        if entries:
            prev = entries[-1]

    entries.append(_trend_entry(summary))
    if max_entries > 0 and len(entries) > max_entries:
        entries = entries[-max_entries:]

    write_json(trend_path, entries)
    return entries, prev

def _timing_stats(values: list[int]) -> dict[str, Any]:
    values_sorted = sorted(values)
    if not values_sorted:
        return {
            "count": 0,
            "mean_ms": None,
            "p50_ms": None,
            "p90_ms": None,
            "max_ms": None,
        }

    float_values = [float(v) for v in values_sorted]
    return {
        "count": len(values_sorted),
        "mean_ms": statistics.fmean(float_values),
        "p50_ms": _percentile(float_values, 0.50),
        "p90_ms": _percentile(float_values, 0.90),
        "max_ms": values_sorted[-1],
    }


def _compute_timings(reports: list[dict[str, Any]]) -> dict[str, Any]:
    """Aggregate per-step `duration_ms` metrics across triage reports."""

    durations: dict[str, list[int]] = {step: [] for step in TIMING_STEPS}
    for r in reports:
        steps = r.get("steps")
        if not isinstance(steps, dict):
            continue
        for step in TIMING_STEPS:
            step_out = steps.get(step)
            if not isinstance(step_out, dict):
                continue
            # Timing metrics should reflect successful work only; failures are surfaced separately
            # via the compatibility counts/gates.
            if step_out.get("status") != "ok":
                continue
            duration = step_out.get("duration_ms")
            # JSON booleans are ints in python, so explicitly exclude bools.
            if isinstance(duration, bool):
                continue
            if isinstance(duration, (int, float)):
                durations[step].append(int(duration))

    return {step: _timing_stats(vals) for step, vals in durations.items()}


def _round_trip_failure_kind(report: dict[str, Any]) -> str | None:
    kind = report.get("round_trip_failure_kind")
    if isinstance(kind, str) and kind:
        return kind
    return infer_round_trip_failure_kind(report)


def _compute_summary(reports: list[dict[str, Any]]) -> dict[str, Any]:
    """Compute the JSON summary structure consumed by the dashboard markdown/trend files."""

    total = len(reports)
    open_ok = sum(1 for r in reports if r.get("result", {}).get("open_ok") is True)
    calc_ok = sum(1 for r in reports if r.get("result", {}).get("calculate_ok") is True)
    render_ok = sum(1 for r in reports if r.get("result", {}).get("render_ok") is True)
    rt_ok = sum(1 for r in reports if r.get("result", {}).get("round_trip_ok") is True)

    calc_attempted = sum(
        1 for r in reports if _attempted(r.get("result", {}).get("calculate_ok"))
    )
    render_attempted = sum(
        1 for r in reports if _attempted(r.get("result", {}).get("render_ok"))
    )

    failures_by_category: Counter[str] = Counter()
    failures_by_round_trip_failure_kind: Counter[str] = Counter()
    failing_function_counts: Counter[str] = Counter()
    failing_feature_counts: Counter[str] = Counter()
    failing_diff_fingerprint_counts: Counter[str] = Counter()
    failing_diff_fingerprint_samples: dict[str, dict[str, str]] = {}
    diff_totals: Counter[str] = Counter()
    diff_part_critical: Counter[str] = Counter()
    diff_part_total: Counter[str] = Counter()
    diff_part_group_critical: Counter[str] = Counter()
    diff_part_group_total: Counter[str] = Counter()
    passing_cellxfs: list[int] = []
    failing_cellxfs: list[int] = []
    failing_cellxfs_by_workbook: list[tuple[int, str]] = []
    round_trip_fail_on_values: set[str] = set()

    for r in reports:
        res = r.get("result", {})
        rt_fail_on = res.get("round_trip_fail_on")
        if isinstance(rt_fail_on, str) and rt_fail_on:
            round_trip_fail_on_values.add(rt_fail_on)

        failed = any(
            res.get(k) is False for k in ("open_ok", "calculate_ok", "render_ok", "round_trip_ok")
        )
        cellxfs_val = (r.get("style_stats") or {}).get("cellXfs")
        cellxfs: int | None = cellxfs_val if isinstance(cellxfs_val, int) else None

        if failed:
            failures_by_category[r.get("failure_category", "unknown")] += 1
            for fn, cnt in (r.get("functions") or {}).items():
                failing_function_counts[fn] += int(cnt)
            for feat, enabled in (r.get("features") or {}).items():
                if enabled is True:
                    failing_feature_counts[feat] += 1
            if cellxfs is not None:
                failing_cellxfs.append(cellxfs)
                failing_cellxfs_by_workbook.append((cellxfs, r.get("display_name", "?")))

            # Fingerprint top diffs (privacy-safe) to find the most common diff patterns across
            # failing workbooks.
            diff_step = (r.get("steps") or {}).get("diff") or {}
            diff_details = diff_step.get("details") if isinstance(diff_step, dict) else None
            if isinstance(diff_details, dict):
                top_diffs = diff_details.get("top_differences")
                if isinstance(top_diffs, list):
                    for entry in top_diffs:
                        if not isinstance(entry, dict):
                            continue
                        fp = entry.get("fingerprint")
                        if not isinstance(fp, str) or not fp:
                            continue
                        failing_diff_fingerprint_counts[fp] += 1
                        if fp not in failing_diff_fingerprint_samples:
                            failing_diff_fingerprint_samples[fp] = {
                                "part": entry.get("part") or "",
                                "path": entry.get("path") or "",
                                "kind": entry.get("kind") or "",
                            }
        else:
            if cellxfs is not None:
                passing_cellxfs.append(cellxfs)

        # Higher-signal round-trip bucket based on which OPC parts changed.
        if r.get("failure_category") == "round_trip_diff":
            kind = _round_trip_failure_kind(r) or "round_trip_other"
            failures_by_round_trip_failure_kind[kind] += 1

        for key, out_key in [
            ("diff_critical_count", "critical"),
            ("diff_warning_count", "warning"),
            ("diff_info_count", "info"),
        ]:
            val = res.get(key)
            if isinstance(val, int):
                diff_totals[out_key] += val

        diff_step = (r.get("steps") or {}).get("diff") or {}
        diff_details = diff_step.get("details") or {}
        parts_with_diffs = diff_details.get("parts_with_diffs")
        part_groups = diff_details.get("part_groups") or {}
        if not isinstance(part_groups, dict):
            part_groups = {}
        if isinstance(parts_with_diffs, list):
            for row in parts_with_diffs:
                if not isinstance(row, dict):
                    continue
                part = row.get("part")
                if not isinstance(part, str) or not part:
                    continue
                group = row.get("group")
                if not isinstance(group, str) or not group:
                    group = part_groups.get(part)
                if not isinstance(group, str) or not group:
                    group = None
                critical = row.get("critical")
                warning = row.get("warning")
                info = row.get("info")
                total_part = row.get("total")
                if isinstance(critical, int):
                    diff_part_critical[part] += critical
                    if group is not None:
                        diff_part_group_critical[group] += critical
                if isinstance(total_part, int):
                    diff_part_total[part] += total_part
                    if group is not None:
                        diff_part_group_total[group] += total_part
                else:
                    # Back-compat: older schemas might not provide `total`.
                    part_sum = 0
                    for v in (critical, warning, info):
                        if isinstance(v, int):
                            part_sum += v
                    if part_sum:
                        diff_part_total[part] += part_sum
                        if group is not None:
                            diff_part_group_total[group] += part_sum

    top_diff_parts_critical = [
        {"part": part, "count": count}
        for part, count in sorted(diff_part_critical.items(), key=lambda kv: (-kv[1], kv[0]))[:10]
        if count > 0
    ]
    top_diff_parts_total = [
        {"part": part, "count": count}
        for part, count in sorted(diff_part_total.items(), key=lambda kv: (-kv[1], kv[0]))[:10]
        if count > 0
    ]

    top_diff_part_groups_critical = [
        {"group": group, "count": count}
        for group, count in sorted(
            diff_part_group_critical.items(), key=lambda kv: (-kv[1], kv[0])
        )[:10]
        if count > 0
    ]
    top_diff_part_groups_total = [
        {"group": group, "count": count}
        for group, count in sorted(
            diff_part_group_total.items(), key=lambda kv: (-kv[1], kv[0])
        )[:10]
        if count > 0
    ]

    style_summary: dict[str, Any] = {}
    if passing_cellxfs or failing_cellxfs:
        style_summary["cellXfs"] = {
            "passing": {
                "count": len(passing_cellxfs),
                "avg": _mean(passing_cellxfs),
                "median": _median(passing_cellxfs),
            },
            "failing": {
                "count": len(failing_cellxfs),
                "avg": _mean(failing_cellxfs),
                "median": _median(failing_cellxfs),
            },
        }
    if failing_cellxfs_by_workbook:
        failing_cellxfs_by_workbook.sort(key=lambda x: (-x[0], x[1]))
        style_summary["top_failing_by_cellXfs"] = [
            {"workbook": name, "cellXfs": cellxfs}
            for cellxfs, name in failing_cellxfs_by_workbook[:20]
        ]

    timings = _compute_timings(reports)
    summary: dict[str, Any] = {
        "timestamp": utc_now_iso(),
        "commit": github_commit_sha(),
        "run_url": github_run_url(),
        "round_trip_fail_on": (
            sorted(round_trip_fail_on_values)[0]
            if len(round_trip_fail_on_values) == 1
            else sorted(round_trip_fail_on_values)
            if round_trip_fail_on_values
            else None
        ),
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
        "round_trip_size_overhead": _round_trip_size_overhead(reports),
        "failures_by_category": dict(failures_by_category),
        "failures_by_round_trip_failure_kind": dict(failures_by_round_trip_failure_kind),
        "diff_totals": dict(diff_totals),
        "timings": timings,
        "top_diff_parts_critical": top_diff_parts_critical,
        "top_diff_parts_total": top_diff_parts_total,
        "top_diff_part_groups_critical": top_diff_part_groups_critical,
        "top_diff_part_groups_total": top_diff_part_groups_total,
        "top_functions_in_failures": [
            {"function": fn, "count": cnt} for fn, cnt in failing_function_counts.most_common(50)
        ],
        "top_features_in_failures": [
            {"feature": feat, "count": cnt} for feat, cnt in failing_feature_counts.most_common(50)
        ],
        "top_diff_fingerprints_in_failures": [
            {
                "fingerprint": fp,
                "count": cnt,
                "part": (failing_diff_fingerprint_samples.get(fp) or {}).get("part", ""),
                "path": (failing_diff_fingerprint_samples.get(fp) or {}).get("path", ""),
                "kind": (failing_diff_fingerprint_samples.get(fp) or {}).get("kind", ""),
            }
            for fp, cnt in failing_diff_fingerprint_counts.most_common(20)
        ],
    }
    if style_summary:
        summary["style"] = style_summary

    return summary

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
    if summary.get("round_trip_fail_on"):
        fail_on = summary["round_trip_fail_on"]
        if isinstance(fail_on, list):
            fail_on_str = ", ".join(str(v) for v in fail_on)
        else:
            fail_on_str = str(fail_on)
        lines.append(f"- Round-trip fail-on: `{fail_on_str}`")
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
        lines.append(f"- {label}: **{ok} / {attempted} attempted** ({extra})")

    lines.append(
        f"- Round-trip: **{counts['round_trip_ok']} / {counts['total']}** ({rates['round_trip']:.1%})"
    )
    diff_totals = summary.get("diff_totals") or {}
    if diff_totals:
        lines.append(
            f"- Diff totals (critical/warn/info): **{diff_totals.get('critical', 0)} / {diff_totals.get('warning', 0)} / {diff_totals.get('info', 0)}**"
        )
    lines.append("")

    timings = summary.get("timings") or {}
    if timings:
        lines.append("## Timings")
        lines.append("")
        lines.append("_Computed from successful (`status: ok`) step executions only._")
        lines.append("")
        lines.append("| Step | Count | Mean (ms) | P50 (ms) | P90 (ms) | Max (ms) |")
        lines.append("|---|---:|---:|---:|---:|---:|")

        def _fmt_ms(v: Any) -> str:
            if v is None:
                return "—"
            if isinstance(v, float):
                if v.is_integer():
                    return str(int(v))
                return f"{v:.1f}"
            return str(v)

        for step in TIMING_STEPS:
            row = timings.get(step) or {}
            lines.append(
                "| "
                + " | ".join(
                    [
                        step,
                        str(row.get("count", 0)),
                        _fmt_ms(row.get("mean_ms")),
                        _fmt_ms(row.get("p50_ms")),
                        _fmt_ms(row.get("p90_ms")),
                        _fmt_ms(row.get("max_ms")),
                    ]
                )
                + " |"
            )
        lines.append("")

    overhead = summary.get("round_trip_size_overhead") or {}
    lines.append("## Round-trip size overhead")
    lines.append("")
    if isinstance(overhead, dict) and overhead.get("count", 0):
        lines.append(f"- Workbooks with size data: **{overhead.get('count', 0)}**")
        mean = overhead.get("mean")
        p50 = overhead.get("p50")
        p90 = overhead.get("p90")
        max_ratio = overhead.get("max")
        if all(isinstance(v, (int, float)) for v in (mean, p50, p90, max_ratio)):
            lines.append(
                "- Size ratio (output/input): "
                f"mean **{float(mean):.3f}**, "
                f"p50 **{float(p50):.3f}**, "
                f"p90 **{float(p90):.3f}**, "
                f"max **{float(max_ratio):.3f}**"
            )
        lines.append(
            "- Exceeding ratio thresholds (>1.05 / >1.10): "
            f"**{overhead.get('count_over_1_05', 0)} / {overhead.get('count_over_1_10', 0)}**"
        )
    else:
        lines.append("_No successful round-trip size metrics found._")
    lines.append("")

    top_critical_parts = summary.get("top_diff_parts_critical") or []
    if top_critical_parts:
        lines.append("## Top diff parts (CRITICAL)")
        lines.append("")
        lines.append("| Part | Critical diffs |")
        lines.append("|---|---:|")
        for row in top_critical_parts[:10]:
            lines.append(f"| {row['part']} | {row['count']} |")
        lines.append("")

    top_any_parts = summary.get("top_diff_parts_total") or []
    if top_any_parts:
        lines.append("## Top diff parts (all severities)")
        lines.append("")
        lines.append("| Part | Total diffs |")
        lines.append("|---|---:|")
        for row in top_any_parts[:10]:
            lines.append(f"| {row['part']} | {row['count']} |")
        lines.append("")

    top_critical_groups = summary.get("top_diff_part_groups_critical") or []
    if top_critical_groups:
        lines.append("## Top diff part groups (CRITICAL)")
        lines.append("")
        lines.append("| Group | Critical diffs |")
        lines.append("|---|---:|")
        for row in top_critical_groups[:10]:
            lines.append(f"| {row['group']} | {row['count']} |")
        lines.append("")

    top_any_groups = summary.get("top_diff_part_groups_total") or []
    if top_any_groups:
        lines.append("## Top diff part groups (all severities)")
        lines.append("")
        lines.append("| Group | Total diffs |")
        lines.append("|---|---:|")
        for row in top_any_groups[:10]:
            lines.append(f"| {row['group']} | {row['count']} |")
        lines.append("")

    part_ratio = summary.get("part_change_ratio") or {}
    part_ratio_critical = summary.get("part_change_ratio_critical") or {}
    if isinstance(part_ratio, dict) and part_ratio.get("count", 0):
        lines.append("## Part-level change ratio (privacy-safe)")
        lines.append("")
        lines.append("| Metric | Mean | P50 | P90 | N |")
        lines.append("|---|---:|---:|---:|---:|")

        def _fmt_pct(v: Any) -> str:
            if v is None:
                return "—"
            try:
                return f"{float(v):.1%}"
            except Exception:  # noqa: BLE001
                return "—"

        lines.append(
            "| parts_changed / parts_total | "
            + " | ".join(
                [
                    _fmt_pct(part_ratio.get("mean")),
                    _fmt_pct(part_ratio.get("p50")),
                    _fmt_pct(part_ratio.get("p90")),
                    str(part_ratio.get("count", 0)),
                ]
            )
            + " |"
        )

        if not isinstance(part_ratio_critical, dict):
            part_ratio_critical = {}
        lines.append(
            "| parts_changed_critical / parts_total | "
            + " | ".join(
                [
                    _fmt_pct(part_ratio_critical.get("mean")),
                    _fmt_pct(part_ratio_critical.get("p50")),
                    _fmt_pct(part_ratio_critical.get("p90")),
                    str(part_ratio_critical.get("count", 0)),
                ]
            )
            + " |"
        )
        lines.append("")
    lines.append("## Per-workbook")
    lines.append("")
    lines.append(
        "| Workbook | Open | Calculate | Render | Round-trip | Diff (C/W/I) | Failure category | Round-trip kind |"
    )
    lines.append("|---|---:|---:|---:|---:|---:|---|---|")
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
                    (_round_trip_failure_kind(r) or "") if r.get("failure_category") == "round_trip_diff" else "",
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

    rt_kinds = summary.get("failures_by_round_trip_failure_kind", {})
    if rt_kinds:
        lines.append("## Round-trip failures by kind")
        lines.append("")
        lines.append("| Kind | Count |")
        lines.append("|---|---:|")
        for k, v in sorted(rt_kinds.items(), key=lambda kv: (-kv[1], kv[0])):
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

    top_diff_fingerprints = summary.get("top_diff_fingerprints_in_failures", [])
    if top_diff_fingerprints:
        lines.append("## Top diff fingerprints in failing workbooks")
        lines.append("")
        lines.append("| Fingerprint | Count | Part | Kind | Path |")
        lines.append("|---|---:|---|---|---|")
        for row in top_diff_fingerprints[:20]:
            fp = row.get("fingerprint") or ""
            fp_prefix = fp[:16] if isinstance(fp, str) else ""
            lines.append(
                "| "
                + " | ".join(
                    [
                        fp_prefix,
                        str(row.get("count", 0)),
                        row.get("part", "") or "",
                        row.get("kind", "") or "",
                        row.get("path", "") or "",
                    ]
                )
                + " |"
            )
        lines.append("")

    style = summary.get("style") or {}
    cellxfs = (style.get("cellXfs") or {}) if isinstance(style, dict) else {}
    if cellxfs:
        def _fmt_float(v: Any) -> str:
            if v is None:
                return ""
            try:
                return f"{float(v):.1f}"
            except Exception:  # noqa: BLE001
                return ""

        lines.append("## Style complexity (cellXfs)")
        lines.append("")
        lines.append("| Group | Workbooks | Avg cellXfs | Median cellXfs |")
        lines.append("|---|---:|---:|---:|")
        for group in ("passing", "failing"):
            row = cellxfs.get(group) or {}
            lines.append(
                "| "
                + " | ".join(
                    [
                        group,
                        str(row.get("count", 0)),
                        _fmt_float(row.get("avg")),
                        _fmt_float(row.get("median")),
                    ]
                )
                + " |"
            )
        lines.append("")

    top_failing_by_cellxfs = (
        style.get("top_failing_by_cellXfs") if isinstance(style, dict) else None
    ) or []
    if top_failing_by_cellxfs:
        lines.append("## Top failing workbooks by cellXfs")
        lines.append("")
        lines.append("| Workbook | cellXfs |")
        lines.append("|---|---:|")
        for row in top_failing_by_cellxfs[:20]:
            lines.append(f"| {row.get('workbook', '?')} | {row.get('cellXfs', 0)} |")
        lines.append("")

    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate corpus compatibility dashboard.")
    parser.add_argument("--triage-dir", type=Path, required=True)
    parser.add_argument("--out-dir", type=Path, help="Defaults to --triage-dir")
    parser.add_argument(
        "--privacy-mode",
        choices=[_PRIVACY_PUBLIC, _PRIVACY_PRIVATE],
        default=_PRIVACY_PUBLIC,
        help="Redact potentially sensitive fields (e.g. hash non-github.com run URLs).",
    )
    parser.add_argument(
        "--append-trend",
        type=Path,
        help="Append a compact time-series entry for this run to the given JSON list file.",
    )
    parser.add_argument(
        "--trend-max-entries",
        type=int,
        default=TREND_MAX_ENTRIES,
        help=f"Maximum number of entries to keep in --append-trend output (default: {TREND_MAX_ENTRIES}).",
    )
    parser.add_argument(
        "--gate-load-p90-ms",
        type=int,
        help="Optional CI gate: fail if load p90 exceeds this threshold (ms).",
    )
    parser.add_argument(
        "--gate-round-trip-p90-ms",
        type=int,
        help="Optional CI gate: fail if round_trip p90 exceeds this threshold (ms).",
    )
    args = parser.parse_args()
    if args.trend_max_entries < 0:
        parser.error("--trend-max-entries must be >= 0")

    triage_dir = args.triage_dir
    out_dir = args.out_dir or triage_dir
    ensure_dir(out_dir)

    reports_dir = triage_dir / "reports"
    reports = _load_reports(reports_dir)
    summary = _compute_summary(reports)

    # If the dashboard is generated outside of CI (e.g. from a downloaded artifact), GitHub env
    # vars may be missing. Prefer the triage run metadata from `index.json` when available.
    index_path = triage_dir / "index.json"
    if index_path.exists():
        try:
            index = json.loads(index_path.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001 (tooling)
            index = None
        if isinstance(index, dict):
            # Prefer the triage run's metadata over local fallbacks (e.g. util.github_commit_sha()
            # may fall back to `git rev-parse HEAD` on local machines, which can be misleading when
            # analyzing an artifact from a different revision).
            commit = index.get("commit")
            if isinstance(commit, str) and commit:
                summary["commit"] = commit
            run_url = index.get("run_url")
            if isinstance(run_url, str) and run_url:
                summary["run_url"] = run_url

    run_url = summary.get("run_url")
    if isinstance(run_url, str) and run_url.startswith("sha256="):
        # Already redacted by triage (privacy-mode=private). Avoid double hashing.
        pass
    else:
        summary["run_url"] = _redact_run_url(run_url, privacy_mode=args.privacy_mode)
    timings = summary.get("timings") or {}
    if not isinstance(timings, dict):
        timings = {}

    summary.update(_part_change_ratio_summary(reports))

    write_json(out_dir / "summary.json", summary)
    (out_dir / "summary.md").write_text(
        _markdown_summary(summary, reports), encoding="utf-8"
    )

    if args.append_trend:
        _append_trend_file(args.append_trend, summary=summary, max_entries=args.trend_max_entries)

    gate_failures: list[str] = []
    if args.gate_load_p90_ms is not None:
        load_p90 = (timings.get("load") or {}).get("p90_ms")
        if not isinstance(load_p90, (int, float)) or isinstance(load_p90, bool):
            print(
                "TIMING GATE ERROR: load p90 unavailable (no successful 'load' samples)."
            )
            return 2
        if load_p90 > args.gate_load_p90_ms:
            gate_failures.append(
                f"load_p90_ms={load_p90} exceeds threshold {args.gate_load_p90_ms}"
            )
    if args.gate_round_trip_p90_ms is not None:
        rt_p90 = (timings.get("round_trip") or {}).get("p90_ms")
        if not isinstance(rt_p90, (int, float)) or isinstance(rt_p90, bool):
            print(
                "TIMING GATE ERROR: round_trip p90 unavailable (no successful 'round_trip' samples)."
            )
            return 2
        if rt_p90 > args.gate_round_trip_p90_ms:
            gate_failures.append(
                f"round_trip_p90_ms={rt_p90} exceeds threshold {args.gate_round_trip_p90_ms}"
            )

    if gate_failures:
        for msg in gate_failures:
            print(f"TIMING REGRESSION: {msg}")
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
