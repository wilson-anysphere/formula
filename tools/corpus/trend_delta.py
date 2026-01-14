#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def _load_json(path: Path) -> Any | None:
    try:
        text = path.read_text(encoding="utf-8").strip()
    except FileNotFoundError:
        return None
    if not text:
        return None
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return None


def _load_trend_entries(path: Path) -> list[dict[str, Any]] | None:
    raw = _load_json(path)
    if not isinstance(raw, list):
        return None
    return [e for e in raw if isinstance(e, dict)]


def _load_summary(path: Path) -> dict[str, Any] | None:
    raw = _load_json(path)
    if not isinstance(raw, dict):
        return None
    return raw


def _is_num(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _pct(v: Any) -> str:
    if not _is_num(v):
        return "n/a"
    return f"{float(v):.2%}"


def _delta_pct(a: Any, b: Any) -> str:
    if not (_is_num(a) and _is_num(b)):
        return "n/a"
    return f"{(float(b) - float(a)) * 100:+.2f}pp"


def _ms(v: Any) -> str:
    if not _is_num(v):
        return "n/a"
    f = float(v)
    if abs(f - round(f)) < 1e-9:
        return f"{int(round(f))}ms"
    return f"{f:.1f}ms"


def _delta_ms(a: Any, b: Any) -> str:
    if not (_is_num(a) and _is_num(b)):
        return "n/a"
    d = float(b) - float(a)
    if abs(d - round(d)) < 1e-9:
        return f"{int(round(d)):+d}ms"
    return f"{d:+.1f}ms"


def _ratio(v: Any) -> str:
    if not _is_num(v):
        return "n/a"
    return f"{float(v):.3f}"


def _delta_ratio(a: Any, b: Any) -> str:
    if not (_is_num(a) and _is_num(b)):
        return "n/a"
    return f"{float(b) - float(a):+.3f}"


def _delta_int(a: Any, b: Any) -> str:
    if not (isinstance(a, int) and isinstance(b, int) and not isinstance(a, bool) and not isinstance(b, bool)):
        return "n/a"
    return f"{b - a:+d}"


def _top_delta_list(
    prev_map: dict[str, Any], cur_map: dict[str, Any], *, max_items: int = 3
) -> str | None:
    if not (isinstance(prev_map, dict) and isinstance(cur_map, dict)):
        return None
    items: list[tuple[str, int]] = []
    for k, v in cur_map.items():
        if isinstance(v, bool) or not isinstance(v, int):
            continue
        items.append((str(k), int(v)))
    items.sort(key=lambda kv: (-kv[1], kv[0]))
    top = items[:max_items]
    if not top:
        return None

    parts: list[str] = []
    for k, v in top:
        prev_v = prev_map.get(k)
        prev_i = int(prev_v) if isinstance(prev_v, int) and not isinstance(prev_v, bool) else 0
        parts.append(f"{k}={v} ({v - prev_i:+d})")
    return ", ".join(parts)


def trend_delta_markdown(
    entries: list[dict[str, Any]], *, summary: dict[str, Any] | None = None
) -> str | None:
    if len(entries) < 2:
        return None

    prev, cur = entries[-2], entries[-1]

    if summary is not None:
        ts = summary.get("timestamp")
        if isinstance(ts, str) and ts and cur.get("timestamp") != ts:
            # Trend file didn't get a new entry for this run; skip to avoid publishing a stale delta.
            return None

    lines: list[str] = []
    lines.append("## Trend delta (vs previous private run)")
    lines.append("")
    lines.append(
        f"- Open rate: **{_pct(cur.get('open_rate'))}** ({_delta_pct(prev.get('open_rate'), cur.get('open_rate'))})"
    )
    lines.append(
        f"- Round-trip rate: **{_pct(cur.get('round_trip_rate'))}** ({_delta_pct(prev.get('round_trip_rate'), cur.get('round_trip_rate'))})"
    )
    lines.append(
        f"- Load p90: **{_ms(cur.get('load_p90_ms'))}** ({_delta_ms(prev.get('load_p90_ms'), cur.get('load_p90_ms'))})"
    )
    lines.append(
        f"- Round-trip p90: **{_ms(cur.get('round_trip_p90_ms'))}** ({_delta_ms(prev.get('round_trip_p90_ms'), cur.get('round_trip_p90_ms'))})"
    )
    lines.append(
        f"- Calc rate (attempted): **{_pct(cur.get('calc_rate'))}** ({_delta_pct(prev.get('calc_rate'), cur.get('calc_rate'))}); attempted {cur.get('calc_attempted', 0)}"
    )
    lines.append(
        f"- Render rate (attempted): **{_pct(cur.get('render_rate'))}** ({_delta_pct(prev.get('render_rate'), cur.get('render_rate'))}); attempted {cur.get('render_attempted', 0)}"
    )
    lines.append(
        "- Size ratio p90 (output/input): "
        f"**{_ratio(cur.get('size_overhead_p90'))}** "
        f"({_delta_ratio(prev.get('size_overhead_p90'), cur.get('size_overhead_p90'))}); "
        f"samples {cur.get('size_overhead_samples', 0)}"
    )
    lines.append(
        f"- Part change ratio p90: **{_pct(cur.get('part_change_ratio_p90'))}** ({_delta_pct(prev.get('part_change_ratio_p90'), cur.get('part_change_ratio_p90'))})"
    )
    lines.append(
        f"- Part change ratio p90 (critical): **{_pct(cur.get('part_change_ratio_critical_p90'))}** ({_delta_pct(prev.get('part_change_ratio_critical_p90'), cur.get('part_change_ratio_critical_p90'))})"
    )

    top_kinds = _top_delta_list(
        prev.get("failures_by_round_trip_failure_kind") or {},
        cur.get("failures_by_round_trip_failure_kind") or {},
    )
    if top_kinds:
        lines.append(f"- Top round-trip failure kinds: {top_kinds}")

    top_categories = _top_delta_list(
        prev.get("failures_by_category") or {},
        cur.get("failures_by_category") or {},
    )
    if top_categories:
        lines.append(f"- Top failure categories: {top_categories}")

    prev_diff = prev.get("diff_totals") or {}
    cur_diff = cur.get("diff_totals") or {}
    lines.append(
        "- Diff totals (critical/warn/info): "
        f"**{cur_diff.get('critical', 0)}/{cur_diff.get('warning', 0)}/{cur_diff.get('info', 0)}** "
        "("
        f"{_delta_int(prev_diff.get('critical'), cur_diff.get('critical'))}/"
        f"{_delta_int(prev_diff.get('warning'), cur_diff.get('warning'))}/"
        f"{_delta_int(prev_diff.get('info'), cur_diff.get('info'))}"
        ")"
    )

    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description="Emit a Markdown trend delta from a corpus trend.json.")
    parser.add_argument("--trend-json", type=Path, required=True)
    parser.add_argument(
        "--summary-json",
        type=Path,
        help="Optional: when provided, only emit output if the last trend entry matches summary.json timestamp.",
    )
    args = parser.parse_args()

    entries = _load_trend_entries(args.trend_json)
    if not entries:
        return 0

    summary = _load_summary(args.summary_json) if args.summary_json else None
    md = trend_delta_markdown(entries, summary=summary)
    if md:
        print(md, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

