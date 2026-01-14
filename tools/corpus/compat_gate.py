#!/usr/bin/env python3

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

from .util import load_json


def _rate(passed: int, total: int) -> float:
    if total == 0:
        return 0.0
    return passed / total


def _fmt_rate(passed: int, total: int) -> str:
    rate = _rate(passed, total)
    return f"{passed}/{total} ({rate:.2%})"


def _fmt_pct(rate: float | None) -> str:
    if rate is None:
        return "SKIP"
    return f"{rate:.2%}"


def _get_int(obj: dict[str, Any], key: str) -> int:
    val = obj.get(key)
    if isinstance(val, bool):
        # Guard against accidentally treating booleans as integers.
        raise TypeError(f"summary.json field {key!r} must be an int, got bool")
    if isinstance(val, (int, float)):
        return int(val)
    raise TypeError(f"summary.json field {key!r} must be an int, got {type(val).__name__}")


def _get_optional_int(obj: dict[str, Any], key: str) -> int | None:
    if key not in obj:
        return None
    val = obj.get(key)
    if val is None:
        return None
    if isinstance(val, bool):
        # Guard against accidentally treating booleans as integers.
        raise TypeError(f"summary.json field {key!r} must be an int, got bool")
    if isinstance(val, (int, float)):
        return int(val)
    raise TypeError(f"summary.json field {key!r} must be an int, got {type(val).__name__}")


def _resolve_summary_path(*, triage_dir: Path | None, summary_json: Path | None) -> Path:
    if summary_json is not None:
        return summary_json
    if triage_dir is None:
        raise ValueError("Expected --triage-dir or --summary-json")
    return triage_dir / "summary.json"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Fail CI if corpus compatibility rates drop below configured thresholds."
    )
    parser.add_argument("--triage-dir", type=Path)
    parser.add_argument("--summary-json", type=Path)

    parser.add_argument("--min-open-rate", type=float)
    parser.add_argument("--min-calc-rate", type=float)
    parser.add_argument("--min-render-rate", type=float)
    parser.add_argument("--min-round-trip-rate", type=float)

    args = parser.parse_args(argv)

    if (
        args.min_open_rate is None
        and args.min_calc_rate is None
        and args.min_render_rate is None
        and args.min_round_trip_rate is None
    ):
        print(
            "CORPUS GATE ERROR: No thresholds configured. Pass at least one of "
            "--min-open-rate/--min-round-trip-rate/--min-calc-rate/--min-render-rate."
        )
        return 2

    try:
        summary_path = _resolve_summary_path(
            triage_dir=args.triage_dir, summary_json=args.summary_json
        )
    except ValueError as e:
        print(f"CORPUS GATE ERROR: {e}")
        return 2

    if not summary_path.exists():
        print(f"CORPUS GATE ERROR: summary.json not found: {summary_path}")
        return 2

    try:
        summary = load_json(summary_path)
    except Exception as e:  # noqa: BLE001 (tooling)
        print(f"CORPUS GATE ERROR: Failed to read {summary_path}: {e}")
        return 2

    if not isinstance(summary, dict):
        print(f"CORPUS GATE ERROR: Expected {summary_path} to contain a JSON object")
        return 2

    counts = summary.get("counts")
    if not isinstance(counts, dict):
        print(f"CORPUS GATE ERROR: Expected {summary_path} to contain a 'counts' object")
        return 2

    try:
        total = _get_int(counts, "total")
        open_ok = _get_int(counts, "open_ok")
        calc_ok = _get_int(counts, "calculate_ok")
        render_ok = _get_int(counts, "render_ok")
        rt_ok = _get_int(counts, "round_trip_ok")
        # Newer summaries include `*_attempted` so calculate/render are measured among attempted
        # workbooks only. Older summaries implicitly treated all workbooks as attempted.
        calc_attempted = _get_optional_int(counts, "calculate_attempted")
        render_attempted = _get_optional_int(counts, "render_attempted")
    except Exception as e:  # noqa: BLE001 (tooling)
        print(f"CORPUS GATE ERROR: Invalid counts in {summary_path}: {e}")
        return 2

    if total <= 0:
        print(
            "CORPUS GATE ERROR: summary.json has total=0 workbooks; refusing to pass on an empty corpus."
        )
        return 2

    # Back-compat: older summaries did not include `*_attempted`.
    if calc_attempted is None:
        calc_attempted = total
    if render_attempted is None:
        render_attempted = total

    if calc_attempted < 0 or calc_attempted > total:
        print(
            "CORPUS GATE ERROR: Invalid counts in summary.json: "
            f"calculate_attempted={calc_attempted} must be in [0, total={total}]"
        )
        return 2
    if render_attempted < 0 or render_attempted > total:
        print(
            "CORPUS GATE ERROR: Invalid counts in summary.json: "
            f"render_attempted={render_attempted} must be in [0, total={total}]"
        )
        return 2

    if args.min_calc_rate is not None and calc_attempted == 0:
        print(
            "CORPUS GATE ERROR: --min-calc-rate was set but no calculate results were attempted "
            "(counts.calculate_attempted=0). Either enable the calculate step in triage or "
            "remove --min-calc-rate."
        )
        return 2

    if args.min_render_rate is not None and render_attempted == 0:
        print(
            "CORPUS GATE ERROR: --min-render-rate was set but no render results were attempted "
            "(counts.render_attempted=0). Either enable the render step in triage or "
            "remove --min-render-rate."
        )
        return 2

    actual = {
        "open": _rate(open_ok, total),
        # Calculate/render are optional steps: compute rates among attempted workbooks only.
        "calculate": (calc_ok / calc_attempted) if calc_attempted else None,
        "render": (render_ok / render_attempted) if render_attempted else None,
        "round_trip": _rate(rt_ok, total),
    }

    labels = {
        "open": "open",
        "calculate": "calculate",
        "render": "render",
        "round_trip": "round-trip",
    }

    thresholds: dict[str, float | None] = {
        "open": args.min_open_rate,
        "calculate": args.min_calc_rate,
        "render": args.min_render_rate,
        "round_trip": args.min_round_trip_rate,
    }

    violations: list[str] = []
    for metric, min_rate in thresholds.items():
        if min_rate is None:
            continue
        metric_rate = actual[metric]
        if metric_rate is None:
            # Should be unreachable because we guard config errors above, but keep this safe.
            violations.append(
                f"{labels[metric]} SKIP (0 attempted) < min {min_rate:.2%}"
            )
            continue
        if metric_rate + 1e-12 < min_rate:
            if metric == "open":
                details = _fmt_rate(open_ok, total)
            elif metric == "calculate":
                details = _fmt_rate(calc_ok, calc_attempted)
            elif metric == "render":
                details = _fmt_rate(render_ok, render_attempted)
            else:
                details = _fmt_rate(rt_ok, total)
            violations.append(
                f"{labels[metric]} {details} < min {min_rate:.2%}"
            )

    if violations:
        # CI-friendly (single line + compact counts/rates).
        print("CORPUS GATE FAIL: " + "; ".join(violations))
        print(
            "Actual rates: "
            + ", ".join(
                [
                    f"open={_fmt_pct(actual['open'])}",
                    f"round-trip={_fmt_pct(actual['round_trip'])}",
                    f"calculate={_fmt_pct(actual['calculate'])}",
                    f"render={_fmt_pct(actual['render'])}",
                ]
            )
        )
        return 1

    checked = {k: v for k, v in thresholds.items() if v is not None}
    checked_str = ", ".join(f"{labels[k]}>={v:.2%}" for k, v in checked.items())
    print(
        f"CORPUS GATE PASS: {checked_str} "
        "("
        + ", ".join(
            [
                f"open={_fmt_pct(actual['open'])}",
                f"round-trip={_fmt_pct(actual['round_trip'])}",
                f"calculate={_fmt_pct(actual['calculate'])}",
                f"render={_fmt_pct(actual['render'])}",
            ]
        )
        + ")"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
