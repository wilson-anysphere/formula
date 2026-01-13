#!/usr/bin/env python3

from __future__ import annotations

import argparse
import os
from collections import defaultdict
from pathlib import Path
from typing import Any, DefaultDict

from . import triage as triage_mod
from .util import (
    WorkbookInput,
    github_commit_sha,
    github_run_url,
    read_workbook_input,
    sha256_hex,
    utc_now_iso,
    write_json,
)


def _relationship_id_from_diff_path(path: str) -> str | None:
    """Extract `rId...` relationship ids from xlsx-diff XML paths.

    `xlsx-diff` uses XPath-ish paths for `.rels` parts, e.g.:

    `/Relationships/Relationship[@Id="rId3"]@Target`

    We only surface the Id itself (privacy-safe) and ignore all other path data.
    """

    needle = '[@Id="'
    start = path.find(needle)
    if start < 0:
        return None
    rest = path[start + len(needle) :]
    end = rest.find('"')
    if end < 0:
        return None
    rid = rest[:end].strip()
    if not rid:
        return None
    return rid


def _diff_details_from_rust_output(rust_out: dict[str, Any]) -> dict[str, Any]:
    steps = rust_out.get("steps") or {}
    diff_step = steps.get("diff") or {}
    details = diff_step.get("details") or {}
    if not isinstance(details, dict):
        return {}
    return details


def _ensure_full_diff_entries(
    rust_exe: Path,
    workbook: WorkbookInput,
    *,
    diff_ignore: set[str],
    diff_limit: int,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """Run the Rust triage helper and return (diff_details, differences).

    The Rust helper always computes the full diff, but it only *emits* up to `diff_limit`
    entries in `top_differences`. For per-part counts we need the complete set of emitted
    differences, so we detect truncation and re-run with `diff_limit=total`.
    """

    rust_out = triage_mod._run_rust_triage(  # noqa: SLF001 (internal reuse)
        rust_exe,
        workbook.data,
        diff_ignore=diff_ignore,
        diff_limit=diff_limit,
        recalc=False,
        render_smoke=False,
    )
    details = _diff_details_from_rust_output(rust_out)

    diffs = details.get("top_differences") or []
    if not isinstance(diffs, list):
        diffs = []

    total = ((details.get("counts") or {}).get("total")) if isinstance(details.get("counts"), dict) else None
    if isinstance(total, int) and total >= 0 and len(diffs) < total:
        rust_out = triage_mod._run_rust_triage(  # noqa: SLF001 (internal reuse)
            rust_exe,
            workbook.data,
            diff_ignore=diff_ignore,
            diff_limit=total,
            recalc=False,
            render_smoke=False,
        )
        details = _diff_details_from_rust_output(rust_out)
        diffs = details.get("top_differences") or []
        if not isinstance(diffs, list):
            diffs = []

    return details, diffs


def summarize_differences(
    diffs: list[dict[str, Any]],
) -> tuple[dict[str, dict[str, int]], list[str], dict[str, list[str]]]:
    """Return (per_part_counts, critical_parts, rels_part_to_rIds).

    This function is privacy-safe by construction: it only looks at the safe diff fields
    emitted by the Rust triage helper (severity/part/path/kind) and never returns raw
    expected/actual text.
    """

    per_part: DefaultDict[str, dict[str, int]] = defaultdict(
        lambda: {"critical": 0, "warning": 0, "info": 0, "total": 0}
    )
    rel_ids_by_part: DefaultDict[str, set[str]] = defaultdict(set)

    for d in diffs:
        if not isinstance(d, dict):
            continue
        part = d.get("part")
        if not isinstance(part, str) or not part:
            continue
        severity = d.get("severity")
        if not isinstance(severity, str):
            severity = ""

        counts = per_part[part]
        counts["total"] += 1
        if severity == "CRITICAL":
            counts["critical"] += 1
        elif severity == "WARN":
            counts["warning"] += 1
        elif severity == "INFO":
            counts["info"] += 1

        if part.endswith(".rels") and severity == "CRITICAL":
            path = d.get("path")
            if isinstance(path, str) and path:
                rid = _relationship_id_from_diff_path(path)
                if rid:
                    rel_ids_by_part[part].add(rid)

    per_part_sorted = {k: per_part[k] for k in sorted(per_part)}
    critical_parts = sorted([k for k, v in per_part.items() if v.get("critical", 0) > 0])
    rels_sorted = {k: sorted(rel_ids_by_part[k]) for k in sorted(rel_ids_by_part)}

    return per_part_sorted, critical_parts, rels_sorted


def _format_counts(counts: dict[str, int]) -> str:
    return f'{counts.get("critical", 0)}/{counts.get("warning", 0)}/{counts.get("info", 0)} (total {counts.get("total", 0)})'


def minimize_workbook(
    workbook: WorkbookInput,
    *,
    rust_exe: Path,
    diff_ignore: set[str],
    diff_limit: int,
) -> dict[str, Any]:
    diff_details, diffs = _ensure_full_diff_entries(
        rust_exe,
        workbook,
        diff_ignore=diff_ignore,
        diff_limit=diff_limit,
    )
    per_part_counts, critical_parts, rels_ids = summarize_differences(diffs)

    # Overall counts from the Rust helper (computed on the full diff report regardless of
    # truncation). Keep these as a cross-check.
    overall_counts = diff_details.get("counts") if isinstance(diff_details, dict) else None
    if not isinstance(overall_counts, dict):
        overall_counts = {}

    return {
        "display_name": workbook.display_name,
        "sha256": sha256_hex(workbook.data),
        "size_bytes": len(workbook.data),
        "timestamp": utc_now_iso(),
        "commit": github_commit_sha(),
        "run_url": github_run_url(),
        "diff_ignore": sorted(diff_ignore),
        "diff_counts": {
            "critical": int(overall_counts.get("critical") or 0),
            "warning": int(overall_counts.get("warning") or 0),
            "info": int(overall_counts.get("info") or 0),
            "total": int(overall_counts.get("total") or 0),
        },
        "critical_parts": critical_parts,
        "part_counts": per_part_counts,
        "rels_critical_ids": rels_ids,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Minimize a round-trip diff by reporting which XLSX parts are responsible."
    )
    parser.add_argument("--input", type=Path, required=True, help="Workbook path (*.xlsx, *.b64, *.enc)")
    parser.add_argument(
        "--out",
        type=Path,
        help="Write JSON summary to this path (default: tools/corpus/out/minimize/<sha16>.json).",
    )
    parser.add_argument(
        "--fernet-key-env",
        default="CORPUS_ENCRYPTION_KEY",
        help="Env var containing Fernet key used to decrypt *.enc workbook inputs.",
    )
    parser.add_argument(
        "--diff-ignore",
        action="append",
        default=[],
        help="Additional part path to ignore during diff (can be repeated).",
    )
    parser.add_argument(
        "--diff-limit",
        type=int,
        default=25000,
        help=(
            "Max diff entries to request from the Rust helper in the first pass. The tool will "
            "auto-rerun with the exact total if truncation is detected."
        ),
    )
    args = parser.parse_args()

    fernet_key = os.environ.get(args.fernet_key_env)

    workbook = read_workbook_input(args.input, fernet_key=fernet_key)

    rust_exe = triage_mod._build_rust_helper()  # noqa: SLF001 (internal reuse)
    diff_ignore = set(triage_mod.DEFAULT_DIFF_IGNORE) | {p for p in args.diff_ignore if p}

    summary = minimize_workbook(
        workbook,
        rust_exe=rust_exe,
        diff_ignore=diff_ignore,
        diff_limit=max(0, int(args.diff_limit)),
    )

    workbook_id = summary["sha256"][:16]
    out_path = args.out or (Path("tools/corpus/out/minimize") / f"{workbook_id}.json")
    write_json(out_path, summary)

    # Privacy-safe stdout summary: parts + counts only (no raw XML/text).
    print(f"{summary['display_name']} sha256={workbook_id} diff critical_parts={len(summary['critical_parts'])}")
    for part in summary["critical_parts"]:
        counts = summary["part_counts"].get(part, {})
        print(f"  {part}: {_format_counts(counts)}")
        if part.endswith(".rels"):
            ids = summary["rels_critical_ids"].get(part) or []
            if ids:
                print(f"    rel_ids: {', '.join(ids)}")

    print(f"Wrote summary: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
