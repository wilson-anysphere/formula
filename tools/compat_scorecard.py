#!/usr/bin/env python3
"""
Unified compatibility scorecard generator.

This script merges:

* The XLSX compatibility corpus harness (file I/O) summary:
    tools/corpus/out/**/summary.json
* The Excel-oracle mismatch report (calculation fidelity):
    tests/compatibility/excel-oracle/reports/mismatch-report.json

Output:
* Markdown scorecard (default: compat_scorecard.md)
* Optional JSON output (via --out-json)
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _github_commit_sha() -> str | None:
    sha = os.environ.get("GITHUB_SHA")
    return sha or None


def _github_run_url() -> str | None:
    server = os.environ.get("GITHUB_SERVER_URL")
    repo = os.environ.get("GITHUB_REPOSITORY")
    run_id = os.environ.get("GITHUB_RUN_ID")
    if server and repo and run_id:
        return f"{server}/{repo}/actions/runs/{run_id}"
    return None


def _git_commit_sha(repo_root: Path) -> str | None:
    """
    Best-effort local fallback when not running in GitHub Actions.

    This is intentionally optional: if the repo isn't a git checkout (e.g. running against
    downloaded artifacts), we simply omit the commit field.
    """

    try:
        out = subprocess.check_output(
            ["git", "rev-parse", "HEAD"],
            cwd=repo_root,
            stderr=subprocess.DEVNULL,
            text=True,
        )
    except Exception:
        return None
    sha = out.strip()
    return sha or None


def _load_json(path: Path) -> Any:
    try:
        with path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        raise SystemExit(f"File not found: {path}") from None
    except json.JSONDecodeError as e:
        raise SystemExit(f"Invalid JSON in {path}: {e}") from None


def _as_int(value: Any, *, label: str, path: Path) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise SystemExit(f"Expected int for {label} in {path}, got {type(value).__name__}")
    return value


def _as_nonneg_int(value: Any, *, label: str, path: Path) -> int:
    out = _as_int(value, label=label, path=path)
    if out < 0:
        raise SystemExit(f"Expected non-negative int for {label} in {path}, got {out}")
    return out


def _as_float(value: Any, *, label: str, path: Path) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise SystemExit(f"Expected float for {label} in {path}, got {type(value).__name__}")
    out = float(value)
    if not (0.0 <= out <= 1.0):
        raise SystemExit(f"Expected {label} in {path} to be in [0, 1], got {out}")
    return out


def _as_str_or_none(value: Any) -> str | None:
    if isinstance(value, str) and value:
        return value
    return None


@dataclass(frozen=True)
class CorpusMetrics:
    path: Path
    timestamp: str | None
    commit: str | None
    run_url: str | None
    total_workbooks: int
    open_ok: int
    open_rate: float
    round_trip_ok: int
    round_trip_rate: float


@dataclass(frozen=True)
class OracleMetrics:
    path: Path
    total_cases: int
    mismatches: int
    mismatch_rate: float
    max_mismatch_rate: float | None
    include_tags: list[str] | None
    exclude_tags: list[str] | None
    max_cases: int | None
    cases_sha256: str | None
    expected_path: str | None
    actual_path: str | None


def _parse_corpus_summary(path: Path, payload: Any) -> CorpusMetrics:
    if not isinstance(payload, dict):
        raise SystemExit(f"Corpus summary JSON must be an object: {path}")

    counts = payload.get("counts")
    rates = payload.get("rates")
    if not isinstance(counts, dict):
        raise SystemExit(f"Corpus summary missing counts object: {path}")
    if rates is not None and not isinstance(rates, dict):
        raise SystemExit(f"Corpus summary rates must be an object when present: {path}")

    total = _as_nonneg_int(counts.get("total"), label="counts.total", path=path)
    open_ok = _as_nonneg_int(counts.get("open_ok"), label="counts.open_ok", path=path)
    rt_ok = _as_nonneg_int(counts.get("round_trip_ok"), label="counts.round_trip_ok", path=path)

    if open_ok > total:
        raise SystemExit(
            f"Corpus summary has inconsistent counts in {path}: open_ok={open_ok} > total={total}"
        )
    if rt_ok > total:
        raise SystemExit(
            f"Corpus summary has inconsistent counts in {path}: round_trip_ok={rt_ok} > total={total}"
        )

    open_rate_raw = rates.get("open") if isinstance(rates, dict) else None
    rt_rate_raw = rates.get("round_trip") if isinstance(rates, dict) else None

    open_rate = (
        _as_float(open_rate_raw, label="rates.open", path=path)
        if open_rate_raw is not None
        else (open_ok / total if total else 0.0)
    )
    rt_rate = (
        _as_float(rt_rate_raw, label="rates.round_trip", path=path)
        if rt_rate_raw is not None
        else (rt_ok / total if total else 0.0)
    )

    # Sanity check: when rates are present, they should match the derived ratios.
    if open_rate_raw is not None and total:
        expected = open_ok / total
        if abs(open_rate - expected) > 1e-9:
            raise SystemExit(
                f"Corpus summary has inconsistent open rate in {path}: rates.open={open_rate} != open_ok/total={expected}"
            )
    if rt_rate_raw is not None and total:
        expected = rt_ok / total
        if abs(rt_rate - expected) > 1e-9:
            raise SystemExit(
                f"Corpus summary has inconsistent round-trip rate in {path}: rates.round_trip={rt_rate} != round_trip_ok/total={expected}"
            )

    return CorpusMetrics(
        path=path,
        timestamp=_as_str_or_none(payload.get("timestamp")),
        commit=_as_str_or_none(payload.get("commit")),
        run_url=_as_str_or_none(payload.get("run_url")),
        total_workbooks=total,
        open_ok=open_ok,
        open_rate=open_rate,
        round_trip_ok=rt_ok,
        round_trip_rate=rt_rate,
    )


def _parse_oracle_report(path: Path, payload: Any) -> OracleMetrics:
    if not isinstance(payload, dict):
        raise SystemExit(f"Mismatch report JSON must be an object: {path}")

    summary = payload.get("summary")
    if not isinstance(summary, dict):
        raise SystemExit(f"Mismatch report missing summary object: {path}")

    total = _as_nonneg_int(summary.get("totalCases"), label="summary.totalCases", path=path)
    mismatches = _as_nonneg_int(summary.get("mismatches"), label="summary.mismatches", path=path)
    if mismatches > total:
        raise SystemExit(
            f"Mismatch report has inconsistent counts in {path}: mismatches={mismatches} > totalCases={total}"
        )
    mismatch_rate_raw = summary.get("mismatchRate")
    mismatch_rate = (
        _as_float(mismatch_rate_raw, label="summary.mismatchRate", path=path)
        if mismatch_rate_raw is not None
        else (mismatches / total if total else 0.0)
    )
    if mismatch_rate_raw is not None and total:
        expected = mismatches / total
        if abs(mismatch_rate - expected) > 1e-9:
            raise SystemExit(
                f"Mismatch report has inconsistent mismatchRate in {path}: mismatchRate={mismatch_rate} != mismatches/totalCases={expected}"
            )
    max_rate_raw = summary.get("maxMismatchRate")
    max_rate = None
    if max_rate_raw is not None:
        max_rate = _as_float(max_rate_raw, label="summary.maxMismatchRate", path=path)

    include_tags = summary.get("includeTags")
    if include_tags is not None and not (
        isinstance(include_tags, list) and all(isinstance(t, str) for t in include_tags)
    ):
        raise SystemExit(f"Expected summary.includeTags to be an array of strings in {path}")

    exclude_tags = summary.get("excludeTags")
    if exclude_tags is not None and not (
        isinstance(exclude_tags, list) and all(isinstance(t, str) for t in exclude_tags)
    ):
        raise SystemExit(f"Expected summary.excludeTags to be an array of strings in {path}")

    max_cases_raw = summary.get("maxCases")
    max_cases = None
    if max_cases_raw is not None:
        max_cases = _as_nonneg_int(max_cases_raw, label="summary.maxCases", path=path)

    cases_sha = _as_str_or_none(summary.get("casesSha256"))
    expected_path = _as_str_or_none(summary.get("expectedPath"))
    actual_path = _as_str_or_none(summary.get("actualPath"))

    return OracleMetrics(
        path=path,
        total_cases=total,
        mismatches=mismatches,
        mismatch_rate=mismatch_rate,
        max_mismatch_rate=max_rate,
        include_tags=include_tags,
        exclude_tags=exclude_tags,
        max_cases=max_cases,
        cases_sha256=cases_sha,
        expected_path=expected_path,
        actual_path=actual_path,
    )


def _find_default_corpus_summary(repo_root: Path) -> Path | None:
    out_root = repo_root / "tools" / "corpus" / "out"
    if not out_root.is_dir():
        return None

    candidates = [p for p in out_root.glob("**/summary.json") if p.is_file()]
    if not candidates:
        return None
    # Multiple corpora (public/private/strict variants). Pick the newest output so local runs
    # "just work" even when both public + private corpora are present.
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]


def _fmt_pct(value: float | None) -> str:
    if value is None:
        return "—"
    return f"{value:.2%}"


def _fmt_ratio(passed: int | None, total: int | None) -> str:
    if passed is None or total is None:
        return "—"
    return f"{passed} / {total}"


def _fmt_path(repo_root: Path, path: Path) -> str:
    try:
        return str(path.relative_to(repo_root))
    except ValueError:
        return str(path)


def _rate_status(rate: float | None, *, target: float) -> str:
    if rate is None:
        return "MISSING"
    return "PASS" if rate >= target else "FAIL"


def _fmt_tags(tags: list[str] | None, *, empty_text: str) -> str:
    if tags is None:
        return "—"
    tags = [t for t in tags if t]
    if not tags:
        return empty_text
    return ", ".join(tags)


def _fmt_max_cases(value: int | None) -> str:
    if value is None:
        return "—"
    if value <= 0:
        return "all"
    return str(value)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate a unified compatibility scorecard (corpus + Excel oracle)."
    )
    parser.add_argument(
        "--corpus-summary",
        default="",
        help=(
            "Path to corpus summary.json. If omitted, selects the newest "
            "tools/corpus/out/**/summary.json (public/private/strict variants)."
        ),
    )
    parser.add_argument(
        "--oracle-report",
        default="tests/compatibility/excel-oracle/reports/mismatch-report.json",
        help="Path to Excel-oracle mismatch report JSON (default: %(default)s)",
    )
    parser.add_argument(
        "--out-md",
        default="compat_scorecard.md",
        help="Where to write the markdown scorecard (default: %(default)s)",
    )
    parser.add_argument(
        "--out-json",
        default="",
        help="Optional path to write the scorecard as JSON (default: disabled)",
    )
    parser.add_argument(
        "--target-read",
        type=float,
        default=1.0,
        help="Target pass rate for L1 Read (0-1). Default: 1.0",
    )
    parser.add_argument(
        "--target-calc",
        type=float,
        default=0.999,
        help="Target pass rate for L2 Calculate (0-1). Default: 0.999 (99.9%%)",
    )
    parser.add_argument(
        "--target-round-trip",
        type=float,
        default=0.97,
        help="Target pass rate for L4 Round-trip (0-1). Default: 0.97 (97%%)",
    )
    parser.add_argument(
        "--allow-missing-inputs",
        action="store_true",
        help=(
            "Write a partial scorecard even if one or both inputs are missing (missing metrics "
            "are rendered as '—'). Exit code is 0."
        ),
    )
    args = parser.parse_args()

    for key in ("target_read", "target_calc", "target_round_trip"):
        value = getattr(args, key)
        if not isinstance(value, (int, float)) or isinstance(value, bool) or not (0.0 <= value <= 1.0):
            raise SystemExit(f"--{key.replace('_', '-')} must be a float in [0, 1]. Got: {value!r}")

    repo_root = Path(__file__).resolve().parents[1]

    corpus_path: Path | None
    if args.corpus_summary:
        corpus_path = Path(args.corpus_summary)
        if not corpus_path.is_absolute():
            corpus_path = repo_root / corpus_path
    else:
        corpus_path = _find_default_corpus_summary(repo_root)

    oracle_path = Path(args.oracle_report)
    if not oracle_path.is_absolute():
        oracle_path = repo_root / oracle_path

    missing: list[str] = []
    if corpus_path is None or not corpus_path.is_file():
        searched = repo_root / "tools" / "corpus" / "out"
        msg = (
            "Missing corpus summary.json. "
            "Expected a file under tools/corpus/out/**/summary.json "
            "(run: python -m tools.corpus.triage ... then python -m tools.corpus.dashboard ...)."
        )
        if corpus_path is not None:
            msg += f" Provided/selected path: {corpus_path}"
        else:
            msg += f" Searched under: {searched}"
        missing.append(msg)

    if not oracle_path.is_file():
        missing.append(
            "Missing Excel-oracle mismatch report. "
            "Expected tests/compatibility/excel-oracle/reports/mismatch-report.json "
            "(run: python tools/excel-oracle/compat_gate.py). "
            f"Path: {oracle_path}"
        )

    if missing and not args.allow_missing_inputs:
        for msg in missing:
            sys.stderr.write(msg + "\n")
        return 1

    corpus: CorpusMetrics | None = None
    oracle: OracleMetrics | None = None
    if corpus_path is not None and corpus_path.is_file():
        corpus = _parse_corpus_summary(corpus_path, _load_json(corpus_path))
    if oracle_path.is_file():
        oracle = _parse_oracle_report(oracle_path, _load_json(oracle_path))

    # Targets (project goals). Defaults match the repo's published targets, but can be overridden.
    target_read = float(args.target_read)
    target_calc = float(args.target_calc)  # 99.9% calc fidelity target.
    target_round_trip = float(args.target_round_trip)

    read_total = corpus.total_workbooks if corpus else None
    read_pass = corpus.open_ok if corpus else None
    read_rate = corpus.open_rate if corpus and corpus.total_workbooks > 0 else None

    rt_total = corpus.total_workbooks if corpus else None
    rt_pass = corpus.round_trip_ok if corpus else None
    rt_rate = corpus.round_trip_rate if corpus and corpus.total_workbooks > 0 else None

    corpus_label = corpus.path.parent.name if corpus else None
    corpus_notes_parts: list[str] = []
    if corpus_label:
        corpus_notes_parts.append(f"corpus={corpus_label}")
    if corpus is not None and corpus.total_workbooks == 0:
        corpus_notes_parts.append("no workbooks")
    corpus_notes = ", ".join(corpus_notes_parts) if corpus_notes_parts else "—"

    calc_total = oracle.total_cases if oracle else None
    calc_mismatches = oracle.mismatches if oracle else None
    calc_mismatch_rate = oracle.mismatch_rate if oracle else None
    calc_mismatch_rate_output = (
        calc_mismatch_rate
        if calc_mismatch_rate is not None and calc_total is not None and calc_total > 0
        else None
    )
    calc_pass_rate = (
        (1.0 - calc_mismatch_rate_output)
        if calc_mismatch_rate_output is not None
        else None
    )
    calc_passes = (
        (calc_total - calc_mismatches)
        if calc_total is not None and calc_mismatches is not None
        else None
    )

    commit = _github_commit_sha() or (corpus.commit if corpus else None)
    if not commit:
        commit = _git_commit_sha(repo_root)
    run_url = _github_run_url() or (corpus.run_url if corpus else None)

    out_md = Path(args.out_md)
    if not out_md.is_absolute():
        out_md = repo_root / out_md
    out_md.parent.mkdir(parents=True, exist_ok=True)

    lines: list[str] = []
    lines.append("# Compatibility scorecard")
    lines.append("")
    lines.append(f"- Generated: `{_utc_now_iso()}`")
    if commit:
        lines.append(f"- Commit: `{commit}`")
    if run_url:
        lines.append(f"- Run: {run_url}")
    lines.append("")
    lines.append("## Inputs")
    lines.append("")
    if corpus:
        corpus_meta_parts: list[str] = []
        if corpus.timestamp:
            corpus_meta_parts.append(f"timestamp: `{corpus.timestamp}`")
        if corpus.commit:
            corpus_meta_parts.append(f"commit: `{corpus.commit}`")
        if corpus.run_url:
            corpus_meta_parts.append(f"run: {corpus.run_url}")
        extra = f" ({', '.join(corpus_meta_parts)})" if corpus_meta_parts else ""
        lines.append(f"- Corpus summary: `{_fmt_path(repo_root, corpus.path)}`{extra}")
    else:
        lines.append("- Corpus summary: **MISSING**")
    if oracle:
        oracle_meta_parts: list[str] = []
        oracle_meta_parts.append(f"cases: {oracle.total_cases}")
        oracle_meta_parts.append(f"mismatches: {oracle.mismatches}")
        if oracle.cases_sha256:
            oracle_meta_parts.append(f"casesSha256: `{oracle.cases_sha256[:8]}`")
        oracle_meta_parts.append(f"includeTags: {_fmt_tags(oracle.include_tags, empty_text='<all>')}")
        oracle_meta_parts.append(f"excludeTags: {_fmt_tags(oracle.exclude_tags, empty_text='<none>')}")
        oracle_meta_parts.append(f"maxCases: {_fmt_max_cases(oracle.max_cases)}")
        if oracle.expected_path:
            oracle_meta_parts.append(f"expected: `{oracle.expected_path}`")
        if oracle.actual_path:
            oracle_meta_parts.append(f"actual: `{oracle.actual_path}`")
        extra = f" ({', '.join(oracle_meta_parts)})" if oracle_meta_parts else ""
        lines.append(f"- Excel-oracle mismatch report: `{_fmt_path(repo_root, oracle.path)}`{extra}")
    else:
        lines.append("- Excel-oracle mismatch report: **MISSING**")
    lines.append("")
    lines.append("## Scorecard")
    lines.append("")
    lines.append("| Level | Metric | Status | Pass rate | Passes / Total | Target | Notes |")
    lines.append("| --- | --- | --- | ---: | ---: | ---: | --- |")
    lines.append(
        "| L1 | Read (corpus open) | "
        + _rate_status(read_rate, target=target_read)
        + " | "
        + _fmt_pct(read_rate)
        + " | "
        + _fmt_ratio(read_pass, read_total)
        + " | "
        + _fmt_pct(target_read)
        + " | "
        + corpus_notes
        + " |"
    )

    calc_notes_parts: list[str] = []
    if oracle is not None:
        if calc_total == 0:
            calc_notes_parts.append("no cases")
        else:
            calc_notes_parts.append(f"mismatch rate={_fmt_pct(calc_mismatch_rate_output)}")
            if calc_mismatches is not None and calc_total is not None:
                calc_notes_parts.append(f"mismatches={calc_mismatches}/{calc_total}")
        if oracle.max_mismatch_rate is not None:
            calc_notes_parts.append(f"max={_fmt_pct(oracle.max_mismatch_rate)}")
    calc_notes = ", ".join(calc_notes_parts) if calc_notes_parts else "—"
    lines.append(
        "| L2 | Calculate (Excel oracle) | "
        + _rate_status(calc_pass_rate, target=target_calc)
        + " | "
        + _fmt_pct(calc_pass_rate)
        + " | "
        + _fmt_ratio(calc_passes, calc_total)
        + " | "
        + _fmt_pct(target_calc)
        + " | "
        + calc_notes
        + " |"
    )

    lines.append(
        "| L4 | Round-trip (corpus) | "
        + _rate_status(rt_rate, target=target_round_trip)
        + " | "
        + _fmt_pct(rt_rate)
        + " | "
        + _fmt_ratio(rt_pass, rt_total)
        + " | "
        + _fmt_pct(target_round_trip)
        + " | "
        + corpus_notes
        + " |"
    )
    lines.append("")

    out_md.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")

    if args.out_json:
        out_json = Path(args.out_json)
        if not out_json.is_absolute():
            out_json = repo_root / out_json
        out_json.parent.mkdir(parents=True, exist_ok=True)

        payload: dict[str, Any] = {
            "schemaVersion": 1,
            "generatedAt": _utc_now_iso(),
            "commit": commit,
            "runUrl": run_url,
            "inputs": {
                "corpusSummaryPath": _fmt_path(repo_root, corpus.path) if corpus else None,
                "oracleReportPath": _fmt_path(repo_root, oracle.path) if oracle else None,
                "corpus": {
                    "label": corpus_label,
                    "timestamp": corpus.timestamp if corpus else None,
                    "commit": corpus.commit if corpus else None,
                    "runUrl": corpus.run_url if corpus else None,
                },
                "oracle": {
                    "includeTags": oracle.include_tags if oracle else None,
                    "excludeTags": oracle.exclude_tags if oracle else None,
                    "maxCases": oracle.max_cases if oracle else None,
                    "casesSha256": oracle.cases_sha256 if oracle else None,
                    "expectedPath": oracle.expected_path if oracle else None,
                    "actualPath": oracle.actual_path if oracle else None,
                },
            },
            "metrics": {
                "l1Read": {
                    "status": _rate_status(read_rate, target=target_read),
                    "passRate": read_rate,
                    "passes": read_pass,
                    "total": read_total,
                    "targetPassRate": target_read,
                },
                "l2Calculate": {
                    "status": _rate_status(calc_pass_rate, target=target_calc),
                    "passRate": calc_pass_rate,
                    "mismatchRate": calc_mismatch_rate_output,
                    "maxMismatchRate": oracle.max_mismatch_rate if oracle else None,
                    "passes": calc_passes,
                    "mismatches": calc_mismatches,
                    "total": calc_total,
                    "targetPassRate": target_calc,
                },
                "l4RoundTrip": {
                    "status": _rate_status(rt_rate, target=target_round_trip),
                    "passRate": rt_rate,
                    "passes": rt_pass,
                    "total": rt_total,
                    "targetPassRate": target_round_trip,
                },
            },
        }
        out_json.write_text(
            json.dumps(payload, indent=2, sort_keys=False) + "\n",
            encoding="utf-8",
            newline="\n",
        )

    if missing:
        for msg in missing:
            sys.stderr.write(msg + "\n")
        return 0

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
