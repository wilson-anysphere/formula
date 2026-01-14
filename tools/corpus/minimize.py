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


MAX_DIFF_ENTRIES_AUTO_RERUN = 200_000


def _basename_for_privacy(path_str: str) -> str:
    """Return a privacy-safe basename for a user-supplied path string.

    Note: `path_str` can be a Windows/UNC path or a URI-like string (e.g. file://...). Using
    `Path(path_str).name` directly is OS-dependent and can fail to strip Windows backslashes on
    non-Windows platforms. Normalize slashes first for consistent behavior.
    """

    raw = (path_str or "").strip()
    if not raw:
        return raw
    normalized = raw.replace("\\", "/").rstrip("/")
    if not normalized:
        return raw
    return Path(normalized).name or raw


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
    password: str | None = None,
    password_file: Path | None = None,
    diff_ignore: set[str],
    diff_limit: int,
) -> tuple[dict[str, Any], dict[str, Any], list[dict[str, Any]]]:
    """Run the Rust triage helper and return (rust_out, diff_details, differences).

    The Rust helper always computes the full diff, but it only *emits* up to `diff_limit`
    entries in `top_differences`.

    Older triage outputs don't include per-part summaries, so for accurate per-part counts we
    detect truncation and re-run with `diff_limit=total` (bounded by
    `MAX_DIFF_ENTRIES_AUTO_RERUN`).

    Newer triage outputs include `parts_with_diffs` / `critical_parts`, which are derived from
    the full diff report and are therefore unaffected by `diff_limit`. In that case, callers can
    avoid requesting every diff entry.
    """

    rust_out = triage_mod._run_rust_triage(  # noqa: SLF001 (internal reuse)
        rust_exe,
        workbook.data,
        workbook_name=workbook.display_name,
        password=password,
        password_file=password_file,
        diff_ignore=diff_ignore,
        diff_limit=diff_limit,
        recalc=False,
        render_smoke=False,
    )
    details = _diff_details_from_rust_output(rust_out)

    diffs = details.get("top_differences") or []
    if not isinstance(diffs, list):
        diffs = []

    counts = details.get("counts") or {}
    if not isinstance(counts, dict):
        counts = {}
    total = counts.get("total")

    # Only auto-rerun when we need the full diff entry list. Newer Rust helpers provide per-part
    # diff summaries (`parts_with_diffs` / `critical_parts`) without requiring emitting every
    # difference, and requesting millions of entries can be expensive.
    has_part_summaries = isinstance(details.get("parts_with_diffs"), list) and isinstance(
        details.get("critical_parts"), list
    )
    if (
        not has_part_summaries
        and isinstance(total, int)
        and 0 <= total <= MAX_DIFF_ENTRIES_AUTO_RERUN
        and len(diffs) < total
    ):
        rust_out = triage_mod._run_rust_triage(  # noqa: SLF001 (internal reuse)
            rust_exe,
            workbook.data,
            workbook_name=workbook.display_name,
            password=password,
            password_file=password_file,
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


def _summarize_rels_critical_ids(diffs: list[dict[str, Any]]) -> dict[str, list[str]]:
    rel_ids_by_part: DefaultDict[str, set[str]] = defaultdict(set)
    for d in diffs:
        if not isinstance(d, dict):
            continue
        part = d.get("part")
        severity = d.get("severity")
        if (
            isinstance(part, str)
            and part.endswith(".rels")
            and isinstance(severity, str)
            and severity == "CRITICAL"
        ):
            path = d.get("path")
            if isinstance(path, str) and path:
                rid = _relationship_id_from_diff_path(path)
                if rid:
                    rel_ids_by_part[part].add(rid)
    return {k: sorted(rel_ids_by_part[k]) for k in sorted(rel_ids_by_part)}


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


def _guess_workbook_extension(name: str) -> str:
    lower = (name or "").lower()
    for ext in (".xlsb", ".xlsm", ".xlsx"):
        if lower.endswith(ext):
            return ext
    return ".xlsx"


def minimize_workbook(
    workbook: WorkbookInput,
    *,
    rust_exe: Path,
    password: str | None = None,
    password_file: Path | None = None,
    diff_ignore: set[str],
    diff_limit: int,
    compute_part_hashes: bool = True,
    compute_rels_ids: bool = True,
) -> dict[str, Any]:
    rust_out, diff_details, diffs = _ensure_full_diff_entries(
        rust_exe,
        workbook,
        password=password,
        password_file=password_file,
        diff_ignore=diff_ignore,
        diff_limit=diff_limit,
    )

    # Overall counts from the Rust helper (computed on the full diff report regardless of
    # truncation). Keep these as a cross-check.
    overall_counts = diff_details.get("counts") if isinstance(diff_details, dict) else None
    if not isinstance(overall_counts, dict):
        overall_counts = {}

    total_diffs_raw = overall_counts.get("total")
    total_diffs = int(total_diffs_raw) if isinstance(total_diffs_raw, int) and not isinstance(total_diffs_raw, bool) else 0
    critical_total_raw = overall_counts.get("critical")
    critical_total = int(critical_total_raw) if isinstance(critical_total_raw, int) and not isinstance(critical_total_raw, bool) else 0

    result = rust_out.get("result") or {}
    if not isinstance(result, dict):
        result = {}

    # Prefer per-part summaries provided by newer Rust triage helpers. These are derived from the
    # full diff report even when `top_differences` is truncated.
    per_part_counts: dict[str, dict[str, int]] | None = None
    critical_parts: list[str] | None = None
    rels_ids: dict[str, list[str]] = {}
    parts_with_diffs_out: list[dict[str, Any]] = []
    part_groups: dict[str, str] = {}

    parts_with_diffs = diff_details.get("parts_with_diffs")
    if isinstance(parts_with_diffs, list):
        parsed: dict[str, dict[str, int]] = {}
        for entry in parts_with_diffs:
            if not isinstance(entry, dict):
                continue
            part = entry.get("part")
            if not isinstance(part, str) or not part:
                continue
            group = entry.get("group")
            group_str = group if isinstance(group, str) and group else "other"
            part_groups[part] = group_str
            parsed[part] = {
                "critical": int(entry.get("critical") or 0),
                "warning": int(entry.get("warning") or 0),
                "info": int(entry.get("info") or 0),
                "total": int(entry.get("total") or 0),
            }
            parts_with_diffs_out.append(
                {
                    "part": part,
                    "group": group_str,
                    "critical": int(entry.get("critical") or 0),
                    "warning": int(entry.get("warning") or 0),
                    "info": int(entry.get("info") or 0),
                    "total": int(entry.get("total") or 0),
                }
            )
        per_part_counts = parsed

        # `critical_parts` is optional; recompute from counts to keep ordering stable.
        critical_parts = sorted([p for p, c in parsed.items() if c.get("critical", 0) > 0])

    if per_part_counts is None or critical_parts is None:
        per_part_counts, critical_parts, rels_ids = summarize_differences(diffs)
    elif compute_rels_ids:
        # Relationship Id extraction is best-effort and currently relies on diff paths.
        rels_ids = _summarize_rels_critical_ids(diffs)
    else:
        rels_ids = {}

    emitted_diffs = len(diffs)
    has_critical_rels = any(p.endswith(".rels") for p in critical_parts)

    # If we have critical `.rels` diffs but diff entries were truncated before capturing all
    # CRITICAL diffs, do a bounded rerun to capture all critical diff paths and extract complete
    # relationship Ids. This avoids unbounded reruns on very large workbooks while still producing
    # useful `.rels` metadata for common cases.
    if (
        compute_rels_ids
        and has_critical_rels
        and 0 < critical_total <= MAX_DIFF_ENTRIES_AUTO_RERUN
        and emitted_diffs < critical_total
    ):
        try:
            rerun_out = triage_mod._run_rust_triage(  # noqa: SLF001 (internal reuse)
                rust_exe,
                workbook.data,
                workbook_name=workbook.display_name,
                password=password,
                password_file=password_file,
                diff_ignore=diff_ignore,
                diff_limit=critical_total,
                recalc=False,
                render_smoke=False,
            )
            rerun_details = _diff_details_from_rust_output(rerun_out)
            rerun_diffs = rerun_details.get("top_differences") or []
            if isinstance(rerun_diffs, list):
                diffs = rerun_diffs
                emitted_diffs = len(diffs)
                rels_ids = _summarize_rels_critical_ids(diffs)
        except Exception:
            # Fall back to the partial set extracted from the first run.
            pass

    critical_part_hashes: dict[str, dict[str, Any]] = {}
    critical_part_hashes_error: str | None = None
    if compute_part_hashes and critical_parts:
        try:
            critical_part_hashes = _compute_part_hashes(workbook.data, critical_parts)
        except Exception as e:  # noqa: BLE001 (tooling)
            critical_part_hashes = {}
            critical_part_hashes_error = str(e)

    out: dict[str, Any] = {
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
        "diff_entries_emitted": emitted_diffs,
        "diff_entries_total": total_diffs,
        "diff_entries_truncated": emitted_diffs < total_diffs,
        "critical_parts": critical_parts,
        "part_counts": per_part_counts,
        "parts_with_diffs": parts_with_diffs_out,
        "part_groups": part_groups,
        "rels_critical_ids": rels_ids,
        "rels_critical_ids_complete": (not has_critical_rels)
        or (compute_rels_ids and emitted_diffs >= critical_total),
        "critical_part_hashes": critical_part_hashes,
    }

    # Provide the same actionable round-trip diff bucketing as `tools.corpus.triage` reports when
    # round-trip fails. This is privacy-safe: it only uses OPC part names and part-group labels.
    if out.get("round_trip_ok") is False:
        out["round_trip_failure_kind"] = (
            triage_mod.infer_round_trip_failure_kind(rust_out)  # noqa: SLF001 (internal reuse)
            or "round_trip_other"
        )

    if critical_part_hashes_error:
        out["critical_part_hashes_error"] = critical_part_hashes_error

    return out


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


def _compute_part_hashes(data: bytes, parts: list[str]) -> dict[str, dict[str, Any]]:
    wanted = {_normalize_part_name(p) for p in parts if p}
    out: dict[str, dict[str, Any]] = {}
    if not wanted:
        return out
    with zipfile.ZipFile(io.BytesIO(data), "r") as z:
        for info in z.infolist():
            if info.is_dir():
                continue
            name_norm = _normalize_part_name(info.filename)
            if name_norm not in wanted:
                continue
            b = z.read(info.filename)
            out[name_norm] = {"sha256": sha256_hex(b), "size_bytes": len(b)}
    return out


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
    required: set[str] = {"[Content_Types].xml", "_rels/.rels"}

    # Discover the workbook "main" part by parsing `_rels/.rels` (officeDocument relationship).
    workbook_part: str | None = None
    root_rels = parts.get("_rels/.rels")
    if root_rels:
        try:
            root = ET.fromstring(root_rels)
            for child in root.iter():
                if _xml_local_name(child.tag) != "Relationship":
                    continue
                ty = (child.attrib.get("Type") or "").strip()
                if not ty.lower().endswith("/officedocument"):
                    continue
                target = child.attrib.get("Target") or ""
                resolved = _resolve_relationship_target("_rels/.rels", target)
                if resolved and resolved in parts:
                    workbook_part = resolved
                    break
        except Exception:
            workbook_part = None

    if workbook_part is None:
        # Fallback for malformed packages missing `_rels/.rels` or officeDocument relationship.
        for candidate in ("xl/workbook.xml", "xl/workbook.bin"):
            if candidate in parts:
                workbook_part = candidate
                break

    if workbook_part is None:
        # Last resort: keep the hard-coded XLSX main parts if present.
        for candidate in ("xl/workbook.xml", "xl/_rels/workbook.xml.rels"):
            if candidate in parts:
                required.add(candidate)
        return required

    required.add(workbook_part)

    # Add the workbook relationships part if present.
    wb_dir = posixpath.dirname(workbook_part)
    wb_base = posixpath.basename(workbook_part)
    wb_rels = f"{wb_dir}/_rels/{wb_base}.rels" if wb_dir else f"_rels/{wb_base}.rels"
    if wb_rels in parts:
        required.add(wb_rels)
    else:
        wb_rels = ""

    # Keep parts referenced by workbook relationships that are commonly required for load:
    # - worksheets
    # - styles
    # - shared strings
    if wb_rels and wb_rels in parts:
        try:
            root = ET.fromstring(parts[wb_rels])
            for child in root.iter():
                if _xml_local_name(child.tag) != "Relationship":
                    continue
                ty = (child.attrib.get("Type") or "").strip()
                ty_lower = ty.lower()
                if not (
                    ty_lower.endswith("/worksheet")
                    or ty_lower.endswith("/styles")
                    or ty_lower.endswith("/sharedstrings")
                ):
                    continue
                target = child.attrib.get("Target") or ""
                resolved = _resolve_relationship_target(wb_rels, target)
                if resolved and resolved in parts:
                    required.add(resolved)
        except Exception:
            pass

    # Fallback: keep typical XLSX parts if present (even if not referenced).
    for name in ("xl/styles.xml", "xl/sharedStrings.xml"):
        if name in parts:
            required.add(name)

    return required


def minimize_workbook_package(
    workbook: WorkbookInput,
    *,
    rust_exe: Path,
    password: str | None = None,
    password_file: Path | None = None,
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
        password=password,
        password_file=password_file,
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
            password=password,
            password_file=password_file,
            diff_ignore=diff_ignore,
            diff_limit=diff_limit,
            compute_part_hashes=False,
            compute_rels_ids=False,
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
    parser.add_argument(
        "--input",
        type=Path,
        required=True,
        help="Workbook path (*.xlsx, *.xlsm, *.xlsb, *.b64, *.enc)",
    )
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
    pw_group = parser.add_mutually_exclusive_group()
    pw_group.add_argument(
        "--password",
        help=(
            "Optional password for Office-encrypted workbooks (Excel 'Encrypt with Password'). "
            "Prefer --password-file to avoid exposing secrets in process args."
        ),
    )
    pw_group.add_argument(
        "--password-file",
        type=Path,
        help=(
            "Read Office workbook password from a file (first line). "
            "Enables inspecting Office-encrypted XLSX/XLSM/XLSB workbooks."
        ),
    )
    parser.add_argument(
        "--privacy-mode",
        choices=[triage_mod._PRIVACY_PUBLIC, triage_mod._PRIVACY_PRIVATE],  # noqa: SLF001
        default=triage_mod._PRIVACY_PUBLIC,  # noqa: SLF001
        help=(
            "Control redaction of minimize outputs. `public` preserves filenames/URLs; "
            "`private` anonymizes display_name and hashes non-standard URLs. "
            "In private mode, local output paths and error strings are also redacted."
        ),
    )
    parser.add_argument(
        "--diff-ignore",
        action="append",
        default=[],
        help="Additional part path to ignore during diff (can be repeated).",
    )
    parser.add_argument(
        "--no-default-diff-ignore",
        action="store_true",
        help="Do not ignore default noisy parts (docProps/*).",
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

    password_file: Path | None = None
    if args.password_file:
        password_file = args.password_file.expanduser()
        if not password_file.is_file():
            parser.error(f"--password-file does not exist or is not a file: {password_file}")
        password_file = password_file.resolve()

    workbook = read_workbook_input(args.input, fernet_key=fernet_key)
    if args.privacy_mode == triage_mod._PRIVACY_PRIVATE:  # noqa: SLF001
        # Ensure the Rust helper never sees raw local filenames in privacy mode. The helper only
        # uses the name for extension-based format inference.
        workbook = WorkbookInput(
            display_name=triage_mod._anonymized_display_name(  # noqa: SLF001
                sha256=sha256_hex(workbook.data),
                original_name=workbook.display_name,
            ),
            data=workbook.data,
        )

    rust_exe = triage_mod._build_rust_helper()  # noqa: SLF001 (internal reuse)
    diff_ignore = triage_mod._compute_diff_ignore(  # noqa: SLF001 (internal reuse)
        diff_ignore=args.diff_ignore,
        use_default=not args.no_default_diff_ignore,
    )

    def _apply_privacy_mode(summary: dict[str, Any]) -> None:
        if args.privacy_mode != triage_mod._PRIVACY_PRIVATE:  # noqa: SLF001
            return

        # Match triage privacy-mode behavior to avoid leaking local filenames or corporate domains
        # when summary JSON is uploaded as an artifact.
        summary["display_name"] = triage_mod._anonymized_display_name(  # noqa: SLF001
            sha256=summary.get("sha256") or sha256_hex(workbook.data),
            original_name=workbook.display_name,
        )
        summary["run_url"] = triage_mod._redact_run_url(  # noqa: SLF001
            summary.get("run_url"), privacy_mode=args.privacy_mode
        )
        minimized = summary.get("minimized")
        if isinstance(minimized, dict):
            # Do not leak local output paths; keep only the basename.
            p = minimized.get("path")
            if isinstance(p, str) and p:
                minimized["path"] = _basename_for_privacy(p)
            err = minimized.get("error")
            if isinstance(err, str) and err:
                minimized["error"] = f"sha256={triage_mod._sha256_text(err)}"  # noqa: SLF001
        err = summary.get("critical_part_hashes_error")
        if isinstance(err, str) and err:
            summary["critical_part_hashes_error"] = f"sha256={triage_mod._sha256_text(err)}"  # noqa: SLF001

    summary = minimize_workbook(
        workbook,
        rust_exe=rust_exe,
        password=args.password,
        password_file=password_file,
        diff_ignore=diff_ignore,
        diff_limit=max(0, int(args.diff_limit)),
    )

    workbook_id = summary["sha256"][:16]
    out_path = args.out or (Path("tools/corpus/out/minimize") / f"{workbook_id}.json")

    out_xlsx = args.out_xlsx
    if out_xlsx is not None:
        # Allow passing a directory for convenience, and preserve the original workbook extension.
        ext = _guess_workbook_extension(workbook.display_name)
        raw_out_xlsx = str(args.out_xlsx)
        if out_xlsx.exists() and out_xlsx.is_dir():
            out_xlsx = out_xlsx / f"{workbook_id}.min{ext}"
        elif raw_out_xlsx.endswith(("/", "\\")):
            out_dir = Path(raw_out_xlsx)
            out_xlsx = out_dir / f"{workbook_id}.min{ext}"
        elif out_xlsx.suffix == "":
            out_xlsx = out_xlsx.with_suffix(ext)
        try:
            min_summary, removed_parts, min_bytes = minimize_workbook_package(
                workbook,
                rust_exe=rust_exe,
                password=args.password,
                password_file=password_file,
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

    _apply_privacy_mode(summary)
    write_json(out_path, summary)

    # Privacy-safe stdout summary: parts + counts only (no raw XML/text).
    print(f"{summary['display_name']} sha256={workbook_id} diff critical_parts={len(summary['critical_parts'])}")
    for part in summary["critical_parts"]:
        counts = summary["part_counts"].get(part, {})
        group = (summary.get("part_groups") or {}).get(part) if isinstance(summary.get("part_groups"), dict) else None
        part_meta = (summary.get("critical_part_hashes") or {}).get(part) or {}
        sha = part_meta.get("sha256")
        sha_short = sha[:16] if isinstance(sha, str) else None
        size = part_meta.get("size_bytes")
        suffix = ""
        if sha_short:
            suffix += f" sha256={sha_short}"
        if isinstance(size, int):
            suffix += f" size={size}"
        group_prefix = f"[{group}] " if isinstance(group, str) and group else ""
        print(f"  {group_prefix}{part}: {_format_counts(counts)}{suffix}")
        if part.endswith(".rels"):
            ids = summary["rels_critical_ids"].get(part) or []
            if ids:
                print(f"    rel_ids: {', '.join(ids)}")

    if summary.get("diff_entries_truncated") is True:
        emitted = summary.get("diff_entries_emitted")
        total = summary.get("diff_entries_total")
        print(f"NOTE: diff entries truncated (emitted={emitted} total={total}).")
        if summary.get("rels_critical_ids_complete") is False:
            print(
                f"NOTE: rels_critical_ids may be incomplete (rerun is capped at {MAX_DIFF_ENTRIES_AUTO_RERUN} diffs)."
            )

    wrote_summary = str(out_path)
    if args.privacy_mode == triage_mod._PRIVACY_PRIVATE:  # noqa: SLF001
        wrote_summary = _basename_for_privacy(str(out_path))
    print(f"Wrote summary: {wrote_summary}")
    if out_xlsx is not None:
        minimized = summary.get("minimized")
        if isinstance(minimized, dict) and "path" in minimized:
            wrote_min = str(out_xlsx)
            if args.privacy_mode == triage_mod._PRIVACY_PRIVATE:  # noqa: SLF001
                wrote_min = _basename_for_privacy(str(out_xlsx))
            print(f"Wrote minimized workbook: {wrote_min}")
        elif isinstance(minimized, dict) and "error" in minimized:
            print(f"Minimized workbook not written: {minimized['error']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
