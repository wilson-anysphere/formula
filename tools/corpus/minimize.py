#!/usr/bin/env python3

from __future__ import annotations

import argparse
import io
import os
import posixpath
import zipfile
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

from xml.etree import ElementTree as ET


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
) -> tuple[dict[str, Any], dict[str, Any], list[dict[str, Any]]]:
    """Run the Rust triage helper and return (rust_out, diff_details, differences).

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

    return rust_out, details, diffs


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
    rust_out, diff_details, diffs = _ensure_full_diff_entries(
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

    result = rust_out.get("result") or {}
    if not isinstance(result, dict):
        result = {}

    return {
        "display_name": workbook.display_name,
        "sha256": sha256_hex(workbook.data),
        "size_bytes": len(workbook.data),
        "timestamp": utc_now_iso(),
        "commit": github_commit_sha(),
        "run_url": github_run_url(),
        "open_ok": result.get("open_ok"),
        "round_trip_ok": result.get("round_trip_ok"),
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


def _normalize_part_name(name: str) -> str:
    return name.replace("\\", "/").lstrip("/")


def _normalize_opc_path(path: str) -> str:
    normalized = path.replace("\\", "/")
    out: list[str] = []
    for seg in normalized.split("/"):
        if seg in ("", "."):
            continue
        if seg == "..":
            if out:
                out.pop()
            continue
        out.append(seg)
    return "/".join(out)


def _source_part_from_rels_part(rels_part: str) -> str:
    rels_part = _normalize_part_name(rels_part)
    if rels_part == "_rels/.rels":
        return ""
    if rels_part.startswith("_rels/"):
        rels_file = rels_part[len("_rels/") :]
        return _normalize_opc_path(rels_file[: -len(".rels")] if rels_file.endswith(".rels") else rels_file)
    if "/_rels/" in rels_part:
        dir_part, rels_file = rels_part.rsplit("/_rels/", 1)
        rels_file = rels_file[: -len(".rels")] if rels_file.endswith(".rels") else rels_file
        if dir_part:
            return _normalize_opc_path(f"{dir_part}/{rels_file}")
        return _normalize_opc_path(rels_file)
    return _normalize_opc_path(rels_part[: -len(".rels")] if rels_part.endswith(".rels") else rels_part)


def _resolve_relationship_target(rels_part: str, target: str) -> str:
    target = (target or "").strip().replace("\\", "/")
    if "#" in target:
        target = target.split("#", 1)[0]
    if not target:
        return _source_part_from_rels_part(rels_part)
    if target.startswith("/"):
        return _normalize_opc_path(target.lstrip("/"))
    # Be tolerant of producers that include an `xl/` prefix without a leading `/`, even though
    # relationship targets are supposed to be relative to the source part.
    if target.casefold().startswith("xl/"):
        return _normalize_opc_path(target)

    source_part = _source_part_from_rels_part(rels_part)
    base_dir = f"{posixpath.dirname(source_part)}/" if source_part and "/" in source_part else ""
    return _normalize_opc_path(f"{base_dir}{target}")


def _read_zip_parts(data: bytes) -> dict[str, bytes]:
    parts: dict[str, bytes] = {}
    with zipfile.ZipFile(io.BytesIO(data), "r") as z:
        for info in z.infolist():
            if info.is_dir():
                continue
            normalized = _normalize_part_name(info.filename)
            if normalized in parts:
                raise ValueError(f"duplicate part after normalization: {normalized}")
            parts[normalized] = z.read(info.filename)
    return parts


def _write_zip_parts(parts: dict[str, bytes]) -> bytes:
    # Deterministic-ish ZIP output: stable ordering + stable timestamps.
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for name in sorted(parts):
            info = zipfile.ZipInfo(name)
            info.date_time = (1980, 1, 1, 0, 0, 0)
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, parts[name])
    return buf.getvalue()


def _xml_local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def _prune_content_types_xml(content_types: bytes, keep_parts: set[str]) -> bytes:
    try:
        root = ET.fromstring(content_types)
    except Exception:
        return content_types

    ns = root.tag[1:].split("}", 1)[0] if root.tag.startswith("{") else ""
    if ns:
        ET.register_namespace("", ns)

    removed = []
    for child in list(root):
        if _xml_local_name(child.tag) != "Override":
            continue
        part_name = child.attrib.get("PartName") or ""
        normalized = _normalize_part_name(part_name)
        if normalized not in keep_parts:
            root.remove(child)
            removed.append(normalized)

    if not removed:
        return content_types
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _prune_rels_xml(rels_bytes: bytes, rels_part: str, keep_parts: set[str]) -> bytes:
    try:
        root = ET.fromstring(rels_bytes)
    except Exception:
        return rels_bytes

    ns = root.tag[1:].split("}", 1)[0] if root.tag.startswith("{") else ""
    if ns:
        ET.register_namespace("", ns)

    changed = False
    for child in list(root):
        if _xml_local_name(child.tag) != "Relationship":
            continue
        target_mode = (child.attrib.get("TargetMode") or "").strip()
        if target_mode and target_mode.casefold() == "external":
            continue
        target = child.attrib.get("Target") or ""
        resolved = _resolve_relationship_target(rels_part, target)
        if resolved and resolved not in keep_parts:
            root.remove(child)
            changed = True

    if not changed:
        return rels_bytes
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def prune_xlsx_parts(
    workbook_bytes: bytes,
    *,
    keep_parts: set[str],
    prune_content_types: bool = True,
    prune_rels: bool = True,
) -> bytes:
    parts = _read_zip_parts(workbook_bytes)
    keep_parts_normalized = {_normalize_part_name(p) for p in keep_parts}
    kept: dict[str, bytes] = {k: v for k, v in parts.items() if k in keep_parts_normalized}

    if prune_content_types and "[Content_Types].xml" in kept:
        kept["[Content_Types].xml"] = _prune_content_types_xml(
            kept["[Content_Types].xml"], set(kept.keys())
        )

    if prune_rels:
        for name in list(kept.keys()):
            if not name.endswith(".rels"):
                continue
            kept[name] = _prune_rels_xml(kept[name], name, set(kept.keys()))

    return _write_zip_parts(kept)


def _required_core_parts(parts: dict[str, bytes]) -> set[str]:
    required: set[str] = {
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
    }
    # These are often required for formula-xlsx to parse cells correctly. Keep them when present.
    for name in ("xl/styles.xml", "xl/sharedStrings.xml"):
        if name in parts:
            required.add(name)

    # Keep all worksheets referenced by workbook.xml.rels.
    rels_bytes = parts.get("xl/_rels/workbook.xml.rels")
    if rels_bytes:
        try:
            root = ET.fromstring(rels_bytes)
            for child in root.iter():
                if _xml_local_name(child.tag) != "Relationship":
                    continue
                ty = (child.attrib.get("Type") or "").strip()
                if not ty.endswith("/worksheet"):
                    continue
                target = child.attrib.get("Target") or ""
                resolved = _resolve_relationship_target("xl/_rels/workbook.xml.rels", target)
                if resolved and resolved in parts:
                    required.add(resolved)
        except Exception:
            pass
    return required


def minimize_workbook_package(
    workbook: WorkbookInput,
    *,
    rust_exe: Path,
    diff_ignore: set[str],
    diff_limit: int,
    out_xlsx: Path,
    max_steps: int = 500,
    baseline: dict[str, Any] | None = None,
) -> tuple[dict[str, Any], list[str], bytes]:
    """Attempt to remove non-critical parts while still reproducing the original critical diffs.

    This is a best-effort "Phase 2" minimizer. It is intentionally conservative:
    - never removes required core parts needed for loadability
    - only removes parts *not* in the initial critical part set
    - after each removal, reruns triage and only accepts the change if the critical diff
      signature is unchanged (same set of critical parts + same overall critical count).
    """

    baseline_summary = baseline or minimize_workbook(
        workbook,
        rust_exe=rust_exe,
        diff_ignore=diff_ignore,
        diff_limit=diff_limit,
    )
    baseline_critical = set(baseline_summary.get("critical_parts") or [])
    if not baseline_critical:
        raise RuntimeError("workbook has no critical diffs; nothing to minimize")
    baseline_critical_count = int((baseline_summary.get("diff_counts") or {}).get("critical") or 0)

    parts = _read_zip_parts(workbook.data)
    required = _required_core_parts(parts)
    always_keep = required | baseline_critical

    candidates = [
        p
        for p in parts
        if p not in always_keep
        and not p.endswith("/")  # defensive; parts dict excludes dirs anyway
    ]
    candidates.sort(key=lambda p: (-len(parts.get(p, b"")), p))

    removed: list[str] = []
    current_bytes = workbook.data
    current_summary = baseline_summary
    steps = 0
    for part in candidates:
        if steps >= max_steps:
            break
        steps += 1

        keep = set(_read_zip_parts(current_bytes).keys())
        if part not in keep:
            continue
        keep.remove(part)

        # Always keep baseline critical parts and required core parts.
        keep |= baseline_critical | required

        trial_bytes = prune_xlsx_parts(
            current_bytes,
            keep_parts=keep,
            prune_content_types=True,
            prune_rels=True,
        )
        trial_summary = minimize_workbook(
            WorkbookInput(display_name=workbook.display_name, data=trial_bytes),
            rust_exe=rust_exe,
            diff_ignore=diff_ignore,
            diff_limit=diff_limit,
        )

        trial_open_ok = trial_summary.get("open_ok") is True
        trial_critical = set(trial_summary.get("critical_parts") or [])
        trial_critical_count = int((trial_summary.get("diff_counts") or {}).get("critical") or 0)
        if (
            trial_open_ok
            and trial_summary.get("round_trip_ok") == baseline_summary.get("round_trip_ok")
            and trial_critical == baseline_critical
            and trial_critical_count == baseline_critical_count
        ):
            removed.append(part)
            current_bytes = trial_bytes
            current_summary = trial_summary

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    out_xlsx.write_bytes(current_bytes)
    return current_summary, removed, current_bytes


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
    parser.add_argument(
        "--out-xlsx",
        type=Path,
        help=(
            "Optional: write a minimized workbook variant that attempts to remove parts not needed "
            "to reproduce the original critical diffs."
        ),
    )
    parser.add_argument(
        "--minimize-max-steps",
        type=int,
        default=500,
        help="Maximum number of part-removal attempts when --out-xlsx is used.",
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

    out_xlsx = args.out_xlsx
    if out_xlsx is not None:
        # Allow passing a directory for convenience.
        if out_xlsx.exists() and out_xlsx.is_dir():
            out_xlsx = out_xlsx / f"{workbook_id}.min.xlsx"
        try:
            min_summary, removed_parts, min_bytes = minimize_workbook_package(
                workbook,
                rust_exe=rust_exe,
                diff_ignore=diff_ignore,
                diff_limit=max(0, int(args.diff_limit)),
                out_xlsx=out_xlsx,
                max_steps=max(0, int(args.minimize_max_steps)),
                baseline=summary,
            )
            summary["minimized"] = {
                "path": str(out_xlsx),
                "sha256": sha256_hex(min_bytes),
                "size_bytes": len(min_bytes),
                "removed_parts": removed_parts,
                "diff_counts": min_summary.get("diff_counts"),
                "critical_parts": min_summary.get("critical_parts"),
            }
        except Exception as e:  # noqa: BLE001 (tooling)
            summary["minimized"] = {"error": str(e)}

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
    if out_xlsx is not None:
        minimized = summary.get("minimized")
        if isinstance(minimized, dict) and "path" in minimized:
            print(f"Wrote minimized workbook: {out_xlsx}")
        elif isinstance(minimized, dict) and "error" in minimized:
            print(f"Minimized workbook not written: {minimized['error']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
