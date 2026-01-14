#!/usr/bin/env python3

from __future__ import annotations

import argparse
import concurrent.futures
import hashlib
import io
import inspect
import json
import os
import sys
import shutil
import subprocess
import tempfile
import time
import zipfile
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from .sanitize import scan_xlsx_bytes_for_leaks
from .util import (
    WorkbookInput,
    ensure_dir,
    github_commit_sha,
    github_run_url,
    iter_workbook_paths,
    read_workbook_input,
    sha256_hex,
    utc_now_iso,
    write_json,
)


DEFAULT_DIFF_IGNORE = {
    # These tend to change across round-trips in most writers due to timestamps and app metadata.
    "docProps/core.xml",
    "docProps/app.xml",
    # NOTE: calcChain (`xl/calcChain.xml` / `xl/calcChain.bin`) is intentionally *not* ignored by
    # default. `xlsx-diff` downgrades calcChain-related churn to WARNING so it shows up in
    # metrics/dashboards without failing CI gates (which key off critical diffs).
}

# Glob patterns to ignore during diffing. Keep conservative; this is primarily intended for local
# corpus minimization workflows where you want to suppress clearly non-semantic churn (e.g. media
# assets) without removing the files from the workbook.
DEFAULT_DIFF_IGNORE_GLOBS: set[str] = set()


def _normalize_diff_ignore_part(part: str) -> str:
    """Normalize a diff-ignore part path to match the Rust helper's canonicalization."""

    part = (part or "").strip().replace("\\", "/").lstrip("/")
    return part


def _compute_diff_ignore(*, diff_ignore: list[str], use_default: bool) -> set[str]:
    ignore: set[str] = set()
    if use_default:
        ignore |= DEFAULT_DIFF_IGNORE
    for part in diff_ignore:
        normalized = _normalize_diff_ignore_part(part)
        if normalized:
            ignore.add(normalized)
    return ignore


def _normalize_part_name(part_name: str) -> str:
    # OPC part names should use forward slashes and never start with `/`, but be tolerant since
    # we ingest reports from multiple sources (including synthetic/unit-test fixtures).
    return part_name.replace("\\", "/").lstrip("/")


def _round_trip_fail_on(report: dict[str, Any]) -> str:
    res = report.get("result")
    if isinstance(res, dict):
        fail_on = res.get("round_trip_fail_on")
        if isinstance(fail_on, str):
            fail_on = fail_on.casefold().strip()
            if fail_on in ("critical", "warning", "info"):
                return fail_on
    return "critical"


def _extract_failure_diff_parts(report: dict[str, Any]) -> set[str]:
    """Return normalized part names that contain diffs that contribute to round_trip_ok=False.

    Uses `result.round_trip_fail_on` to decide which severities count as a failure.
    Prefers the Rust helper's per-part summaries (`parts_with_diffs`) when present and falls back to
    scanning the (truncated) `top_differences` list.
    """

    fail_on = _round_trip_fail_on(report)
    parts: set[str] = set()

    def _add_part(value: Any) -> None:
        if not isinstance(value, str):
            return
        normalized = _normalize_part_name(value)
        if normalized:
            parts.add(normalized)

    def _row_has_failure_diff(row: dict[str, Any]) -> bool:
        def _int_field(key: str) -> int:
            v = row.get(key)
            if isinstance(v, bool):
                return 0
            if isinstance(v, int):
                return v
            if isinstance(v, str) and v.isdigit():
                return int(v)
            return 0

        critical = _int_field("critical")
        warning = _int_field("warning")
        info = _int_field("info")
        total = _int_field("total") or (critical + warning + info)

        if fail_on == "critical":
            return critical > 0
        if fail_on == "warning":
            return critical > 0 or warning > 0
        return total > 0

    def _severity_counts_as_failure(sev: str) -> bool:
        sev = (sev or "").upper()
        if fail_on == "critical":
            return sev == "CRITICAL"
        if fail_on == "warning":
            return sev in ("CRITICAL", "WARN", "WARNING")
        return sev in ("CRITICAL", "WARN", "WARNING", "INFO")

    steps = report.get("steps")
    diff_step = steps.get("diff") if isinstance(steps, dict) else None
    diff_details = diff_step.get("details") if isinstance(diff_step, dict) else {}
    res = report.get("result") or {}

    for container in (diff_details, res):
        if not isinstance(container, dict):
            continue
        pwd = container.get("parts_with_diffs")
        if pwd is None:
            continue

        # Newer schema: list of `{part, critical, warning, info, total, ...}`.
        if isinstance(pwd, list):
            for row in pwd:
                if not isinstance(row, dict):
                    continue
                if not _row_has_failure_diff(row):
                    continue
                _add_part(row.get("part"))
            if parts:
                return parts

        # Older schema: dict like {"critical": [...], "warning": [...], ...}.
        if isinstance(pwd, dict):
            wanted_keys: tuple[str, ...]
            if fail_on == "critical":
                wanted_keys = ("critical", "CRITICAL")
            elif fail_on == "warning":
                wanted_keys = ("critical", "CRITICAL", "warning", "WARN", "WARNING")
            else:
                wanted_keys = ("critical", "CRITICAL", "warning", "WARN", "WARNING", "info", "INFO")

            for key in wanted_keys:
                seq = pwd.get(key)
                if isinstance(seq, list):
                    for item in seq:
                        _add_part(item)
            if parts:
                return parts

    # Fallback: scan truncated top differences.
    top = diff_details.get("top_differences") if isinstance(diff_details, dict) else None
    if isinstance(top, list):
        for entry in top:
            if not isinstance(entry, dict):
                continue
            if not _severity_counts_as_failure(entry.get("severity") or ""):
                continue
            _add_part(entry.get("part"))

    return parts


def _extract_failure_diff_part_groups(report: dict[str, Any]) -> set[str]:
    """Return diff part-group names that contain diffs contributing to round_trip_ok=False.

    Uses the Rust helper's per-part summaries when present:
    - `parts_with_diffs`: list of `{part, group, critical, ...}`
    - `critical_parts` + `part_groups`: fallback mapping

    Returns normalized, lowercase group names (e.g. `rels`, `content_types`).
    """

    fail_on = _round_trip_fail_on(report)
    groups: set[str] = set()

    steps = report.get("steps")
    diff_step = steps.get("diff") if isinstance(steps, dict) else None
    diff_details = diff_step.get("details") if isinstance(diff_step, dict) else {}
    res = report.get("result") or {}

    def _add_group(value: Any) -> None:
        if not isinstance(value, str):
            return
        value = value.strip().casefold()
        if value:
            groups.add(value)

    for container in (diff_details, res):
        if not isinstance(container, dict):
            continue

        # Preferred schema: list of per-part summaries that includes a `group` field.
        pwd = container.get("parts_with_diffs")
        if isinstance(pwd, list):
            part_groups_map = container.get("part_groups") or {}
            if not isinstance(part_groups_map, dict):
                part_groups_map = {}
            for row in pwd:
                if not isinstance(row, dict):
                    continue
                critical = row.get("critical")
                warning = row.get("warning")
                info = row.get("info")
                total = row.get("total")
                if isinstance(critical, bool) or not isinstance(critical, int):
                    critical = 0
                if isinstance(warning, bool) or not isinstance(warning, int):
                    warning = 0
                if isinstance(info, bool) or not isinstance(info, int):
                    info = 0
                if isinstance(total, bool) or not isinstance(total, int):
                    total = critical + warning + info

                if fail_on == "critical":
                    ok = critical > 0
                elif fail_on == "warning":
                    ok = critical > 0 or warning > 0
                else:
                    ok = total > 0
                if not ok:
                    continue
                group = row.get("group")
                if not isinstance(group, str) or not group.strip():
                    part = row.get("part")
                    if isinstance(part, str) and part:
                        group = part_groups_map.get(part)
                _add_group(group)
            if groups:
                return groups

        # Fallback schema: `critical_parts` list plus a `part_groups` mapping.
        critical_parts = container.get("critical_parts")
        part_groups = container.get("part_groups")
        if isinstance(critical_parts, list) and isinstance(part_groups, dict):
            for part in critical_parts:
                if not isinstance(part, str):
                    continue
                group = part_groups.get(part)
                _add_group(group)
            if groups:
                return groups

    return groups


def infer_round_trip_failure_kind(report: dict[str, Any]) -> str | None:
    """Return a high-signal category for round-trip diffs.

    This is intentionally privacy-safe: it only looks at OPC part names and part-group labels
    (no XML paths/values).
    """

    res = report.get("result", {})
    if not isinstance(res, dict) or res.get("round_trip_ok") is not False:
        return None

    # Prefer Rust helper part-group labels when available. This is more stable across formats
    # (xlsx vs xlsb) and avoids fragile path-prefix heuristics.
    failure_groups = _extract_failure_diff_part_groups(report)
    if failure_groups:
        if failure_groups == {"rels"}:
            return "round_trip_rels"
        if "content_types" in failure_groups:
            return "round_trip_content_types"
        if "styles" in failure_groups:
            return "round_trip_styles"
        if any(g.startswith("worksheet") for g in failure_groups):
            return "round_trip_worksheets"
        if "shared_strings" in failure_groups:
            return "round_trip_shared_strings"
        if "media" in failure_groups:
            return "round_trip_media"
        if "doc_props" in failure_groups:
            return "round_trip_doc_props"
        if "workbook" in failure_groups:
            return "round_trip_workbook"
        if "theme" in failure_groups:
            return "round_trip_theme"
        if "pivots" in failure_groups:
            return "round_trip_pivots"
        if "charts" in failure_groups:
            return "round_trip_charts"
        if "drawings" in failure_groups:
            return "round_trip_drawings"
        if "tables" in failure_groups:
            return "round_trip_tables"
        if "external_links" in failure_groups:
            return "round_trip_external_links"

    # Either the Rust helper did not provide part-group summaries or the groups are too broad
    # (e.g. everything classified as `other`). Fall back to part-name heuristics.
    failure_parts = {p.casefold() for p in _extract_failure_diff_parts(report)}
    if not failure_parts:
        return "round_trip_other"

    if all(p.endswith(".rels") for p in failure_parts):
        return "round_trip_rels"

    if "[content_types].xml" in failure_parts:
        return "round_trip_content_types"

    if (
        "xl/styles.xml" in failure_parts
        or "xl/styles.bin" in failure_parts
        or any(p.startswith("xl/styles/") for p in failure_parts)
        or any(p.endswith("/styles.xml") or p.endswith("/styles.bin") for p in failure_parts)
    ):
        return "round_trip_styles"

    if any(p.startswith("xl/worksheets/") for p in failure_parts):
        return "round_trip_worksheets"

    if (
        "xl/sharedstrings.xml" in failure_parts
        or "xl/sharedstrings.bin" in failure_parts
        or any(
            p.endswith("/sharedstrings.xml") or p.endswith("/sharedstrings.bin")
            for p in failure_parts
        )
    ):
        return "round_trip_shared_strings"

    if any(p.startswith("xl/media/") for p in failure_parts):
        return "round_trip_media"

    if any(p.startswith("docprops/") for p in failure_parts):
        return "round_trip_doc_props"

    # Additional high-signal buckets for common "other" churn.
    if "xl/workbook.xml" in failure_parts or "xl/workbook.bin" in failure_parts:
        return "round_trip_workbook"

    if any(p.startswith("xl/theme/") for p in failure_parts):
        return "round_trip_theme"

    if any(p.startswith("xl/pivottables/") or p.startswith("xl/pivotcache/") for p in failure_parts):
        return "round_trip_pivots"

    if any(p.startswith("xl/charts/") for p in failure_parts):
        return "round_trip_charts"

    if any(p.startswith("xl/drawings/") for p in failure_parts):
        return "round_trip_drawings"

    if any(p.startswith("xl/tables/") for p in failure_parts):
        return "round_trip_tables"

    if any(p.startswith("xl/externallinks/") for p in failure_parts):
        return "round_trip_external_links"

    return "round_trip_other"
def _normalize_diff_ignore_glob(pattern: str) -> str:
    """Normalize a diff-ignore glob pattern to match the Rust helper's canonicalization."""

    return _normalize_diff_ignore_part(pattern)


def _compute_diff_ignore_globs(
    *, diff_ignore_globs: list[str], use_default: bool
) -> set[str]:
    ignore: set[str] = set()
    if use_default:
        ignore |= DEFAULT_DIFF_IGNORE_GLOBS
    for pattern in diff_ignore_globs:
        normalized = _normalize_diff_ignore_glob(pattern)
        if normalized:
            ignore.add(normalized)
    return ignore


@dataclass(frozen=True)
class StepResult:
    status: str  # ok | failed | skipped
    duration_ms: int | None = None
    error: str | None = None
    details: dict[str, Any] | None = None


@dataclass(frozen=True)
class LeakScanFailure:
    """Returned by triage workers when --leak-scan detects suspicious plaintext.

    We intentionally do not include the matched plaintext, only safe metadata (sha256 of match).
    """

    display_name: str
    findings: list[dict[str, str]]

_PRIVACY_PUBLIC = "public"
_PRIVACY_PRIVATE = "private"

# Relationship types and namespaces often use these domains. We treat them as safe to emit even in
# privacy-mode=private since they're standard schema URLs and don't embed customer-controlled data.
_SAFE_SCHEMA_HOST_SUFFIXES = {
    "schemas.openxmlformats.org",
    "schemas.microsoft.com",
}

# `run_url` can point at a GitHub Enterprise Server domain (e.g. `github.corp.example.com`). When
# triage output is uploaded as an artifact to a less-trusted environment, this can leak corporate
# domains. Allowlist github.com (public) and hash everything else in privacy mode.
_SAFE_RUN_URL_HOST_SUFFIXES = {
    "github.com",
}

# Function fingerprints can include custom add-in / UDF names (e.g. `CORP.ADDIN.FOO`). In private
# corpus artifacts, hash any non-standard function names to avoid leaking internal product/company
# identifiers while keeping built-in Excel functions readable.
_KNOWN_FUNCTION_NAMES: set[str] | None = None


def _load_known_function_names() -> set[str]:
    global _KNOWN_FUNCTION_NAMES
    if _KNOWN_FUNCTION_NAMES is not None:
        return _KNOWN_FUNCTION_NAMES

    names: set[str] = set()
    try:
        catalog_path = _repo_root() / "shared" / "functionCatalog.json"
        data = json.loads(catalog_path.read_text(encoding="utf-8"))
        funcs = data.get("functions") if isinstance(data, dict) else None
        if isinstance(funcs, list):
            for entry in funcs:
                if not isinstance(entry, dict):
                    continue
                name = entry.get("name")
                if isinstance(name, str) and name:
                    names.add(name.upper())
    except Exception:  # noqa: BLE001 (best-effort; privacy mode still works without allowlist)
        names = set()

    _KNOWN_FUNCTION_NAMES = names
    return names


def _redact_function_counts(counts: Counter[str], *, privacy_mode: str) -> dict[str, int]:
    if privacy_mode != _PRIVACY_PRIVATE:
        return dict(counts)

    known = _load_known_function_names()
    out: Counter[str] = Counter()
    for fn, cnt in counts.items():
        # `counts` should already be normalized (uppercase) by `_extract_function_counts`.
        if fn in known:
            out[fn] += cnt
        else:
            out[f"sha256={_sha256_text(fn)}"] += cnt
    return dict(out)

# OpenXML content types are not URIs, but custom producers sometimes embed corporate domains or
# internal product names (e.g. `application/vnd.corp.example+xml`). In privacy mode we keep only
# well-known standard prefixes and hash everything else.
_SAFE_CONTENT_TYPE_PREFIXES = (
    "application/xml",
    "text/xml",
    "application/octet-stream",
    "image/",
    "text/",
    # OpenXML + OPC package content types.
    "application/vnd.openxmlformats-officedocument.",
    "application/vnd.openxmlformats-package.",
    # Microsoft/Office ecosystem content types.
    "application/vnd.ms-",
)


def _anonymized_display_name(*, sha256: str, original_name: str) -> str:
    # Keep the extension as a lightweight hint (xlsx vs xlsm) while avoiding leaking the original
    # filename. Fall back to `.xlsx` for unknown extensions.
    suffix = Path(original_name).suffix.lower()
    if suffix not in (".xlsx", ".xlsm", ".xlsb"):
        suffix = ".xlsx"
    return f"workbook-{sha256[:16]}{suffix}"


def _is_safe_schema_host(host: str | None) -> bool:
    if not host:
        return False
    host = host.casefold()
    return any(
        host == allowed or host.endswith(f".{allowed}") for allowed in _SAFE_SCHEMA_HOST_SUFFIXES
    )


def _redact_uri_like(text: str) -> str:
    """Redact URI-like strings to avoid leaking custom domains in private corpus artifacts."""

    import re
    import urllib.parse

    parsed = urllib.parse.urlparse(text)
    if parsed.scheme in ("http", "https") and parsed.netloc:
        if _is_safe_schema_host(parsed.hostname):
            return text
        return f"sha256={_sha256_text(text)}"

    if text.startswith("~/"):
        return f"sha256={_sha256_text(text)}"

    # Absolute filesystem paths can appear in external relationship targets (e.g. linked workbooks)
    # and can leak usernames/mount points. Hash common OS-level path prefixes.
    if text.startswith("/") and re.search(
        r"^/(Users|home|mnt|Volumes|private|var|opt)/", text, flags=re.IGNORECASE
    ):
        return f"sha256={_sha256_text(text)}"

    # Network-path reference / UNC-like URLs (`//host/path` or `\\\\host\\share`).
    # These are often used for internal file shares and can leak corporate hostnames.
    if text.startswith("//") and parsed.netloc:
        if _is_safe_schema_host(parsed.hostname):
            return text
        return f"sha256={_sha256_text(text)}"
    if text.startswith("\\\\"):
        # Normalize to avoid treating `\\` differently from `//` in downstream tooling.
        normalized = "//" + text.lstrip("\\").replace("\\", "/")
        parsed2 = urllib.parse.urlparse(normalized)
        if _is_safe_schema_host(parsed2.hostname):
            return text
        return f"sha256={_sha256_text(text)}"

    # Non-http schemes (e.g. urn:) are still URI-like; hash them to be safe.
    if parsed.scheme and ":" in text:
        return f"sha256={_sha256_text(text)}"

    # Content types are not URIs, but some custom producers embed domains (e.g. vnd.company.com.*).
    # If the string contains something that *looks* like a hostname with a common TLD, hash it.
    lowered = text.casefold()
    if any(allowed in lowered for allowed in _SAFE_SCHEMA_HOST_SUFFIXES):
        return text

    # Linked workbook filenames can appear in external relationship targets. Hash common spreadsheet
    # file extensions even when the value is a relative path (no scheme/hostname).
    if re.search(r"\.(xlsx|xlsm|xlsb|xltx|xltm|xls|csv|tsv)\b", lowered):
        return f"sha256={_sha256_text(text)}"

    # Redact IPv4 address-like tokens (e.g. `10.0.0.1` or `10.0.0.1/share`). These often appear in
    # internal URLs and can leak corporate network topology.
    if re.search(r"\b\d{1,3}(?:\.\d{1,3}){3}\b", text):
        return f"sha256={_sha256_text(text)}"

    # Keep this intentionally conservative: only hash if we see a likely TLD boundary.
    if re.search(r"\.(com|net|org|io|ai|dev|edu|gov|local|internal|corp)\b", lowered):
        return f"sha256={_sha256_text(text)}"

    return text


def _redact_uri_like_in_text(text: str) -> str:
    """Redact URI-like substrings embedded in a larger string.

    `xlsx-diff` paths often embed expanded XML namespaces in the form `{uri}localName`. In
    privacy-mode=private we want to avoid leaking any non-standard/custom domains, while keeping the
    overall diff path structure readable.
    """

    import re

    # First, redact any `{uri}` namespace expansions (QName rendering).
    def _replace_braced(match: re.Match[str]) -> str:
        inner = match.group(1)
        return "{" + _redact_uri_like(inner) + "}"

    out = re.sub(r"\{([^{}]+)\}", _replace_braced, text)

    # Also redact any raw http(s) URL tokens that appear outside of `{}`.
    def _replace_url(match: re.Match[str]) -> str:
        url = match.group(0)
        return _redact_uri_like(url)

    out = re.sub(r"https?://[^\s\"'<>]+", _replace_url, out)

    # Some OOXML relationship types/targets can use non-http schemes (e.g. `urn:` or `file:`). Only
    # match a conservative allowlist of schemes to avoid accidentally hashing non-URI strings like
    # `A1:B2`.
    out = re.sub(
        r"(?:urn|mailto|file|ftp|ftps|tel|smb):[^\s\"'<>]+", _replace_url, out, flags=re.IGNORECASE
    )

    # Network-path reference / UNC-like URLs (`//host/path`). This can show up for external
    # relationships where TargetMode=External points at a file share.
    out = re.sub(r"(?<!:)//[^\s\"'<>]+", _replace_url, out)

    # As a last resort, redact bare domain-like tokens that appear without a scheme.
    out = re.sub(
        r"(?:[A-Za-z0-9-]+\.)+(?:com|net|org|io|ai|dev|edu|gov|local|internal|corp)\b[^\s\"'<>]*",
        _replace_url,
        out,
        flags=re.IGNORECASE,
    )
    out = re.sub(
        r"\b\d{1,3}(?:\.\d{1,3}){3}(?::\d+)?(?:/[^\s\"'<>]*)?",
        _replace_url,
        out,
    )

    # Common absolute filesystem path prefixes that can appear in external relationship targets.
    out = re.sub(
        r"/(?:Users|home|mnt|Volumes|private|var|opt)/[^\s\"'<>]+",
        _replace_url,
        out,
        flags=re.IGNORECASE,
    )
    out = re.sub(r"\b[A-Za-z]:/[^\s\"'<>]+", _replace_url, out)
    out = re.sub(r"\b[A-Za-z]:\\[^\s\"'<>]+", _replace_url, out)
    out = re.sub(r"\\\\[^\s\"'<>]+", _replace_url, out)
    out = re.sub(r"~/(?:[^\s\"'<>]+)", _replace_url, out)

    # Linked workbook filenames (relative paths) without a domain/scheme.
    out = re.sub(
        r"[^\s\"'<>]+\.(?:xlsx|xlsm|xlsb|xltx|xltm|xls|csv|tsv)\b",
        _replace_url,
        out,
        flags=re.IGNORECASE,
    )
    return out


def _is_safe_run_url_host(host: str | None) -> bool:
    if not host:
        return False
    host = host.casefold()
    return any(
        host == allowed or host.endswith(f".{allowed}") for allowed in _SAFE_RUN_URL_HOST_SUFFIXES
    )


def _redact_run_url(url: str | None, *, privacy_mode: str) -> str | None:
    if not url or privacy_mode != _PRIVACY_PRIVATE:
        return url
    if isinstance(url, str) and url.startswith("sha256="):
        # Already redacted (defense in depth / idempotence).
        return url

    import urllib.parse

    parsed = urllib.parse.urlparse(url)
    if parsed.scheme in ("http", "https") and parsed.netloc and _is_safe_run_url_host(parsed.hostname):
        return url
    return f"sha256={_sha256_text(url)}"


def _redact_content_type(value: str | None, *, privacy_mode: str) -> str | None:
    if not value or privacy_mode != _PRIVACY_PRIVATE:
        return value
    if isinstance(value, str) and value.startswith("sha256="):
        # Already redacted (defense in depth / idempotence).
        return value
    lowered = value.strip().casefold()
    if any(lowered.startswith(prefix) for prefix in _SAFE_CONTENT_TYPE_PREFIXES):
        return value
    return f"sha256={_sha256_text(value)}"


def _now_ms() -> float:
    return time.perf_counter() * 1000.0


def _step_ok(start_ms: float, *, details: dict[str, Any] | None = None) -> StepResult:
    return StepResult(status="ok", duration_ms=int(_now_ms() - start_ms), details=details)


def _sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _step_failed(start_ms: float, err: Exception) -> StepResult:
    # Triage reports are uploaded as artifacts for both public and private corpora. Avoid leaking
    # workbook content (sheet names, defined names, file paths, etc.) through exception strings by
    # hashing the message and emitting only the digest (mirrors the Rust helper behavior).
    msg = str(err)
    return StepResult(
        status="failed",
        duration_ms=int(_now_ms() - start_ms),
        error=f"sha256={_sha256_text(msg)}",
    )


def _step_skipped(reason: str) -> StepResult:
    return StepResult(status="skipped", details={"reason": reason})


def _normalize_opc_path(path: str) -> str:
    """Normalize an OPC part path.

    Mirrors `xlsx_diff::normalize_opc_path` behavior:
    - Convert `\\` to `/`
    - Strip any leading `/`
    - Collapse empty / `.` segments
    - Resolve `..` segments without ever escaping the package root (extra `..` are ignored)
    """

    normalized = (path or "").replace("\\", "/").lstrip("/")
    out: list[str] = []
    for segment in normalized.split("/"):
        match segment:
            case "" | ".":
                continue
            case "..":
                if out:
                    out.pop()
                continue
            case _:
                out.append(segment)
    return "/".join(out)


def _normalize_zip_entry_name(name: str) -> str:
    # ZIP entry names in valid XLSX/XLSM packages should not start with `/`, but some producers do.
    # Normalize for comparisons while preserving the original entry name for `ZipFile.read`.
    return _normalize_opc_path(name)


def _scan_features(zip_names: list[str]) -> dict[str, Any]:
    import posixpath
    import re

    normalized_names = [_normalize_zip_entry_name(n) for n in zip_names]
    normalized_casefold = [name.casefold() for name in normalized_names]

    prefixes = {
        "charts": "xl/charts/",
        "drawings": "xl/drawings/",
        "tables": "xl/tables/",
        "pivot_tables": "xl/pivotTables/",
        "pivot_cache": "xl/pivotCache/",
        "external_links": "xl/externalLinks/",
        "query_tables": "xl/queryTables/",
        "printer_settings": "xl/printerSettings/",
        "custom_xml_root": "customXml/",
        "custom_xml_xl": "xl/customXml/",
    }

    features: dict[str, Any] = {}
    for key, prefix in prefixes.items():
        prefix_casefold = prefix.casefold()
        features[f"has_{key}"] = any(n.startswith(prefix_casefold) for n in normalized_casefold)

    normalized_casefold_set = set(normalized_casefold)
    features["has_vba"] = "xl/vbaproject.bin" in normalized_casefold_set
    features["has_connections"] = "xl/connections.xml" in normalized_casefold_set
    features["has_shared_strings"] = "xl/sharedstrings.xml" in normalized_casefold_set
    # Excel's "Images in Cell" feature (aka `cellImages`).
    cellimages_part_re = re.compile(r"(?i)^cellimages\d*\.xml$")
    features["has_cell_images"] = any(
        _normalize_zip_entry_name(n).casefold().startswith("xl/")
        and cellimages_part_re.match(posixpath.basename(normalized))
        for (n, normalized) in zip(zip_names, normalized_names)
    )
    features["sheet_xml_count"] = sum(
        1 for n in normalized_casefold if n.startswith("xl/worksheets/sheet")
    )
    return features


def _find_zip_entry_case_insensitive(zip_names: list[str], wanted: str) -> str | None:
    wanted_lower = _normalize_zip_entry_name(wanted).casefold()
    for n in zip_names:
        if _normalize_zip_entry_name(n).casefold() == wanted_lower:
            return n
    return None


def _split_etree_tag(tag: str) -> tuple[str, str | None]:
    if tag.startswith("{") and "}" in tag:
        ns, local = tag[1:].split("}", 1)
        return local, ns
    return tag, None


def _extract_cell_images(z: zipfile.ZipFile, zip_names: list[str]) -> dict[str, Any] | None:
    import posixpath
    import re

    cellimages_part_re = re.compile(r"(?i)^cellimages(\d*)\.xml$")
    zip_names_casefold_set = {_normalize_zip_entry_name(n).casefold() for n in zip_names}

    def _is_cellimages_part(n: str) -> bool:
        normalized = _normalize_zip_entry_name(n)
        if not normalized.casefold().startswith("xl/"):
            return False
        return bool(cellimages_part_re.match(posixpath.basename(normalized)))

    candidates = [n for n in zip_names if _is_cellimages_part(n)]
    if not candidates:
        return None

    part_name: str | None = None
    for n in candidates:
        if _normalize_zip_entry_name(n).casefold() == "xl/cellimages.xml":
            part_name = n
            break

    if part_name is None:
        # Prefer the smallest numeric suffix (treat empty suffix as 0), with a stable tie-breaker.
        scored: list[tuple[int, int, int, str, str]] = []
        for n in candidates:
            normalized = _normalize_zip_entry_name(n)
            m = cellimages_part_re.match(posixpath.basename(normalized))
            if not m:
                continue
            suffix_str = m.group(1) or ""
            suffix_num = int(suffix_str) if suffix_str else 0
            # Prefer parts directly under `xl/` (mirrors the most common Excel layout) before
            # deeper `xl/**/` variants if multiple parts share the same numeric suffix.
            direct_under_xl = 0 if posixpath.dirname(normalized).casefold() == "xl" else 1
            path_depth = normalized.count("/")
            scored.append((suffix_num, direct_under_xl, path_depth, n.casefold(), n))
        if not scored:
            return None
        scored.sort()
        part_name = scored[0][4]

    content_type: str | None = None
    content_types_name = _find_zip_entry_case_insensitive(zip_names, "[Content_Types].xml")
    if content_types_name:
        try:
            from xml.etree import ElementTree as ET

            ct_root = ET.fromstring(z.read(content_types_name))
            part_name_normalized = _normalize_zip_entry_name(part_name)
            for el in ct_root.iter():
                if el.tag.split("}")[-1] != "Override":
                    continue
                part = el.attrib.get("PartName") or ""
                normalized_part = _normalize_opc_path(part)
                if normalized_part.casefold() == part_name_normalized.casefold():
                    content_type = el.attrib.get("ContentType")
                    break
        except Exception:
            content_type = None

    workbook_rel_type: str | None = None
    workbook_rels_name = _find_zip_entry_case_insensitive(zip_names, "xl/_rels/workbook.xml.rels")
    if workbook_rels_name:
        try:
            from xml.etree import ElementTree as ET

            def _resolve_workbook_target(target: str) -> str:
                target = (target or "").strip()
                # Relationship targets are URIs; internal targets may include a fragment
                # (e.g. `foo.xml#bar`). OPC part names do not include fragments.
                target = target.split("#", 1)[0].replace("\\", "/")
                if not target:
                    return ""
                if target.startswith("/"):
                    return _normalize_opc_path(target)
                # Some producers incorrectly include the `xl/` prefix without a leading `/`,
                # even though relationship targets are supposed to be relative to the source
                # part (`xl/workbook.xml`). Treat this as a package-root-relative path.
                if target.casefold().startswith("xl/"):
                    return _normalize_opc_path(target)
                return _normalize_opc_path(posixpath.join("xl", target))

            rels_root = ET.fromstring(z.read(workbook_rels_name))
            for el in rels_root.iter():
                if el.tag.split("}")[-1] != "Relationship":
                    continue
                target = el.attrib.get("Target") or ""
                resolved = _resolve_workbook_target(target)
                if not resolved:
                    continue
                part_name_normalized = _normalize_zip_entry_name(part_name)
                if resolved.casefold() == part_name_normalized.casefold():
                    workbook_rel_type = el.attrib.get("Type")
                    break
                # Be tolerant of producers that use an incorrect base path but a correct basename.
                # This helps us still fingerprint the workbook relationship type even when the
                # package layout is non-standard.
                if (
                    resolved.casefold() not in zip_names_casefold_set
                    and posixpath.basename(resolved).casefold()
                    == posixpath.basename(part_name_normalized).casefold()
                ):
                    workbook_rel_type = el.attrib.get("Type")
                    break
        except Exception:
            workbook_rel_type = None

    root_local_name: str | None = None
    root_namespace: str | None = None
    embed_rids_count: int | None = None
    try:
        from xml.etree import ElementTree as ET

        REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        rid_re = re.compile(r"^rId\d+$")

        cellimages_root = ET.fromstring(z.read(part_name))
        root_local_name, root_namespace = _split_etree_tag(cellimages_root.tag)
        embed_rids_count = 0
        for el in cellimages_root.iter():
            el_local_name, _ = _split_etree_tag(el.tag)
            for attr_name, attr_value in el.attrib.items():
                attr_local_name, attr_ns = _split_etree_tag(attr_name)

                # Count relationship-id references in a schema-agnostic way. Some cellImages
                # variants put the relationship ID on a top-level <cellImage r:embed="..."/> or
                # <cellImage r:id="..."/> rather than a nested <a:blip r:embed="..."/>.
                #
                # Keep the legacy heuristic as a subset: count any `embed` attribute on a `blip`
                # element, even if it's not namespace-qualified (some producers are sloppy).
                if not (
                    (attr_ns == REL_NS and attr_local_name in ("embed", "id"))
                    or (el_local_name == "blip" and attr_local_name == "embed")
                ):
                    continue

                if not rid_re.match((attr_value or "").strip()):
                    continue

                embed_rids_count += 1
    except Exception:
        root_local_name = None
        root_namespace = None
        embed_rids_count = None

    rels_types: list[str] = []
    part_name_normalized = _normalize_zip_entry_name(part_name)
    part_dir = posixpath.dirname(part_name_normalized)
    part_base = posixpath.basename(part_name_normalized)
    cellimages_rels_path = posixpath.join(part_dir, "_rels", f"{part_base}.rels")
    cellimages_rels_name = _find_zip_entry_case_insensitive(zip_names, cellimages_rels_path)
    if cellimages_rels_name:
        try:
            from xml.etree import ElementTree as ET

            rels_root = ET.fromstring(z.read(cellimages_rels_name))
            types: set[str] = set()
            for el in rels_root.iter():
                if el.tag.split("}")[-1] != "Relationship":
                    continue
                t = el.attrib.get("Type")
                if t:
                    types.add(t)
            rels_types = sorted(types)
        except Exception:
            rels_types = []

    return {
        "part_name": part_name,
        "content_type": content_type,
        "workbook_rel_type": workbook_rel_type,
        "root_local_name": root_local_name,
        "root_namespace": root_namespace,
        "embed_rids_count": embed_rids_count,
        "rels_types": rels_types,
    }


def _extract_function_counts(z: zipfile.ZipFile) -> Counter[str]:
    """Return function name -> usage count (no formulas or values are emitted)."""

    import re

    fn_re = re.compile(r"(?i)\b([A-Z_][A-Z0-9_.]*)\s*\(")
    string_lit_re = re.compile(r'"(?:[^"]|"")*"')

    counts: Counter[str] = Counter()
    for name in z.namelist():
        # Use a case-insensitive match for worksheet part paths. Some malformed packages store
        # entries like `XL/Worksheets/Sheet1.XML`, which should still be treated as worksheets for
        # privacy-safe formula function fingerprinting.
        normalized_casefold = _normalize_zip_entry_name(name).casefold()
        if not (
            normalized_casefold.startswith("xl/worksheets/")
            and normalized_casefold.endswith(".xml")
        ):
            continue
        root = z.read(name)

        try:
            from xml.etree import ElementTree as ET

            tree = ET.fromstring(root)
        except Exception:
            continue

        for el in tree.iter():
            if el.tag.split("}")[-1] != "f":
                continue
            if not el.text:
                continue
            # Remove string literals before function matching.
            text = string_lit_re.sub('""', el.text)
            for match in fn_re.finditer(text):
                fn = match.group(1).upper()
                for prefix in ("_XLFN.", "_XLWS.", "_XLUDF."):
                    if fn.startswith(prefix):
                        fn = fn[len(prefix) :]
                counts[fn] += 1
    return counts


def _extract_style_stats(
    z: zipfile.ZipFile, zip_names: list[str]
) -> tuple[dict[str, int] | None, str | None]:
    """Extract privacy-safe styling complexity metrics from `xl/styles.xml`.

    Returns (stats, error). If the part is missing, returns (None, None). If parsing fails,
    returns (None, "<error>") without raising.
    """

    styles_name = _find_zip_entry_case_insensitive(zip_names, "xl/styles.xml")
    if not styles_name:
        return None, None

    try:
        styles_bytes = z.read(styles_name)
    except Exception as e:  # noqa: BLE001
        return None, f"Failed to read {styles_name}: {e}"

    try:
        from xml.etree import ElementTree as ET

        root = ET.fromstring(styles_bytes)
    except Exception as e:  # noqa: BLE001
        return None, f"Failed to parse {styles_name}: {e}"

    def _local(tag: str) -> str:
        return tag.split("}")[-1]

    def _parse_int(value: str | None) -> int | None:
        if value is None:
            return None
        try:
            return int(value)
        except Exception:  # noqa: BLE001
            return None

    def _find_child(local_name: str) -> Any | None:
        for child in list(root):
            if _local(child.tag) == local_name:
                return child
        return None

    def _count_container(container_local: str, item_local: str | None) -> int:
        container = _find_child(container_local)
        if container is None:
            return 0

        declared = _parse_int(container.attrib.get("count"))

        if item_local is None:
            if declared is not None:
                return declared
            return len(list(container))

        children = [c for c in list(container) if _local(c.tag) == item_local]
        if children:
            return len(children)
        if declared is not None:
            return declared
        return 0

    stats: dict[str, int] = {
        # Core OOXML style collections.
        "numFmts": _count_container("numFmts", "numFmt"),
        "fonts": _count_container("fonts", "font"),
        "fills": _count_container("fills", "fill"),
        "borders": _count_container("borders", "border"),
        "cellStyleXfs": _count_container("cellStyleXfs", "xf"),
        "cellXfs": _count_container("cellXfs", "xf"),
        "cellStyles": _count_container("cellStyles", "cellStyle"),
        # Differential formats (used by conditional formatting, tables, etc).
        "dxfs": _count_container("dxfs", "dxf"),
        # Optional: custom table styles.
        "tableStyles": _count_container("tableStyles", None),
        # Extension list - often present in modern Excel output.
        "extLst": _count_container("extLst", "ext"),
    }

    return stats, None


def _repo_root() -> Path:
    # tools/corpus/triage.py -> tools/corpus -> tools -> repo root
    return Path(__file__).resolve().parents[2]


def _rust_exe_name() -> str:
    return "formula-corpus-triage.exe" if os.name == "nt" else "formula-corpus-triage"


def _build_rust_helper() -> Path:
    """Build (or reuse) the Rust triage helper binary."""

    root = _repo_root()
    env = os.environ.copy()
    ci_raw = env.get("CI")
    is_ci = bool(ci_raw and ci_raw.strip().casefold() not in ("0", "false", "no"))
    # `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml`. Some environments set it
    # globally (often to `stable`), which would bypass the pinned toolchain and reintroduce
    # "whatever stable is today" drift when building the helper.
    if env.get("RUSTUP_TOOLCHAIN") and (root / "rust-toolchain.toml").is_file():
        env.pop("RUSTUP_TOOLCHAIN", None)
    default_global_cargo_home = Path.home() / ".cargo"
    cargo_home = env.get("CARGO_HOME")
    cargo_home_path = Path(cargo_home).expanduser() if cargo_home else None
    if not cargo_home or (
        not is_ci
        and not env.get("FORMULA_ALLOW_GLOBAL_CARGO_HOME")
        and cargo_home_path == default_global_cargo_home
    ):
        env["CARGO_HOME"] = str(root / "target" / "cargo-home")
    Path(env["CARGO_HOME"]).mkdir(parents=True, exist_ok=True)

    # Some environments configure Cargo to use `sccache` via a global config file. Prefer
    # compiling locally for determinism/reliability unless the user explicitly opted in.
    #
    # Cargo respects both `RUSTC_WRAPPER` and the config/env-var equivalent
    # `CARGO_BUILD_RUSTC_WRAPPER`. When these are unset (or explicitly set to the empty string),
    # override any global config by forcing a benign wrapper (`env`) that simply executes rustc.
    rustc_wrapper = env.get("RUSTC_WRAPPER")
    if rustc_wrapper is None:
        rustc_wrapper = env.get("CARGO_BUILD_RUSTC_WRAPPER")
    if not rustc_wrapper:
        rustc_wrapper = (shutil.which("env") or "env") if os.name != "nt" else ""

    rustc_workspace_wrapper = env.get("RUSTC_WORKSPACE_WRAPPER")
    if rustc_workspace_wrapper is None:
        rustc_workspace_wrapper = env.get("CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER")
    if not rustc_workspace_wrapper:
        rustc_workspace_wrapper = (shutil.which("env") or "env") if os.name != "nt" else ""

    env["RUSTC_WRAPPER"] = rustc_wrapper
    env["RUSTC_WORKSPACE_WRAPPER"] = rustc_workspace_wrapper
    env["CARGO_BUILD_RUSTC_WRAPPER"] = rustc_wrapper
    env["CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER"] = rustc_workspace_wrapper

    # Concurrency defaults: keep Rust builds stable on high-core-count multi-agent hosts.
    #
    # Prefer explicit overrides, but default to a conservative job count when unset. On very
    # high core-count hosts, linking (lld) can spawn many threads per link step; combining that
    # with Cargo-level parallelism can exceed sandbox process/thread limits and cause flaky
    # "Resource temporarily unavailable" failures.
    cpu_count = os.cpu_count() or 0
    default_jobs = 2 if cpu_count >= 64 else 4
    jobs_raw = env.get("FORMULA_CARGO_JOBS") or env.get("CARGO_BUILD_JOBS") or str(default_jobs)
    try:
        jobs_int = int(jobs_raw)
    except ValueError:
        jobs_int = default_jobs
    if jobs_int < 1:
        jobs_int = default_jobs
    jobs = str(jobs_int)

    env["CARGO_BUILD_JOBS"] = jobs
    env.setdefault("MAKEFLAGS", f"-j{jobs}")
    env.setdefault("CARGO_PROFILE_DEV_CODEGEN_UNITS", jobs)
    env.setdefault("CARGO_PROFILE_TEST_CODEGEN_UNITS", jobs)
    env.setdefault("CARGO_PROFILE_RELEASE_CODEGEN_UNITS", jobs)
    env.setdefault("CARGO_PROFILE_BENCH_CODEGEN_UNITS", jobs)
    env.setdefault("RAYON_NUM_THREADS", env.get("FORMULA_RAYON_NUM_THREADS") or jobs)

    target_dir_env = os.environ.get("CARGO_TARGET_DIR")
    if target_dir_env:
        target_dir = Path(target_dir_env)
        # `cargo` interprets relative `CARGO_TARGET_DIR` paths relative to its working directory
        # (we run it from `root`). Make sure we resolve the same way so `exe.exists()` works even
        # when the caller runs this script from another CWD.
        if not target_dir.is_absolute():
            target_dir = root / target_dir
    else:
        target_dir = root / "target"
    exe = target_dir / "debug" / _rust_exe_name()

    try:
        use_cargo_agent = (
            os.name != "nt"
            and shutil.which("bash") is not None
            and (root / "scripts" / "cargo_agent.sh").is_file()
        )
        subprocess.run(
            ["bash", "scripts/cargo_agent.sh", "build", "-p", "formula-corpus-triage"]
            if use_cargo_agent
            else ["cargo", "build", "-p", "formula-corpus-triage"],
            cwd=root,
            env=env,
            check=True,
        )
    except subprocess.CalledProcessError as e:
        # In high-churn dev environments the Rust workspace can temporarily fail to build due to
        # unrelated merge conflicts/regressions. If we already have a previously-built helper
        # binary available, prefer using it for local tooling so users can continue minimizing
        # corpora without waiting on the workspace to compile again.
        #
        # In CI, always treat build failures as fatal to avoid masking real regressions.
        if exe.exists() and not is_ci:
            print(
                f"warning: failed to build Rust triage helper (exit {e.returncode}); using existing binary: {exe}\n"
                "note: this is a local-only fallback; CI treats Rust build failures as fatal",
                file=sys.stderr,
            )
            return exe
        raise
    except FileNotFoundError as e:  # noqa: PERF203 (CI signal)
        if exe.exists() and not is_ci:
            missing = e.filename or "cargo"
            print(
                f"warning: {missing} not found; using existing Rust triage helper: {exe}",
                file=sys.stderr,
            )
            return exe
        raise RuntimeError("cargo not found; Rust toolchain is required for corpus triage") from e

    if not exe.exists():
        raise RuntimeError(f"Rust triage helper was built but executable is missing: {exe}")
    return exe


def _run_rust_triage(
    exe: Path,
    workbook_bytes: bytes,
    *,
    workbook_name: str,
    password: str | None = None,
    password_file: Path | None = None,
    diff_ignore: set[str],
    diff_ignore_globs: set[str] | None = None,
    diff_ignore_path: tuple[str, ...] | list[str] = (),
    diff_ignore_path_in: tuple[str, ...] | list[str] = (),
    diff_ignore_path_kind: tuple[str, ...] | list[str] = (),
    diff_ignore_path_kind_in: tuple[str, ...] | list[str] = (),
    diff_ignore_presets: tuple[str, ...] = tuple(),
    diff_limit: int,
    round_trip_fail_on: str = "critical",
    recalc: bool,
    render_smoke: bool,
    strict_calc_chain: bool = False,
) -> dict[str, Any]:
    """Invoke the Rust helper to run load/save/diff (+ optional recalc/render) on a workbook blob."""

    with tempfile.TemporaryDirectory(prefix="corpus-triage-") as tmpdir:
        # The Rust helper auto-detects based on file extension, so preserve `.xlsb` vs `.xlsx`
        # here (we strip `.b64`/`.enc` earlier in `read_workbook_input`).
        lower_name = (workbook_name or "").lower()
        if lower_name.endswith(".xlsb"):
            suffix = ".xlsb"
            fmt = "xlsb"
        elif lower_name.endswith(".xlsm"):
            suffix = ".xlsm"
            fmt = "xlsx"
        else:
            suffix = ".xlsx"
            fmt = "xlsx"

        input_path = Path(tmpdir) / f"input{suffix}"
        input_path.write_bytes(workbook_bytes)

        cmd = [
            str(exe),
            "--input",
            str(input_path),
            "--format",
            fmt,
            "--diff-limit",
            str(diff_limit),
            "--fail-on",
            round_trip_fail_on,
        ]
        if password_file is not None:
            cmd.extend(["--password-file", str(password_file)])
        elif password is not None:
            cmd.extend(["--password", password])
        for part in sorted(diff_ignore):
            cmd.extend(["--ignore-part", part])
        if diff_ignore_globs is None:
            diff_ignore_globs = set()
        for pattern in sorted({p.strip() for p in diff_ignore_globs if p and p.strip()}):
            cmd.extend(["--ignore-glob", pattern])
        for pattern in sorted({p.strip() for p in diff_ignore_path if p and p.strip()}):
            cmd.extend(["--ignore-path", pattern])
        for scoped in sorted({p.strip() for p in diff_ignore_path_in if p and p.strip()}):
            cmd.extend(["--ignore-path-in", scoped])
        for spec in sorted({p.strip() for p in diff_ignore_path_kind if p and p.strip()}):
            cmd.extend(["--ignore-path-kind", spec])
        for spec in sorted({p.strip() for p in diff_ignore_path_kind_in if p and p.strip()}):
            cmd.extend(["--ignore-path-kind-in", spec])
        for preset in diff_ignore_presets:
            if preset:
                cmd.extend(["--ignore-preset", preset])
        if strict_calc_chain:
            cmd.append("--strict-calc-chain")
        if recalc:
            cmd.append("--recalc")
        if render_smoke:
            cmd.append("--render-smoke")

        proc = subprocess.run(
            cmd,
            cwd=_repo_root(),
            capture_output=True,
            text=True,
        )

        if proc.returncode != 0:
            raise RuntimeError(
                f"Rust triage helper failed (exit {proc.returncode}): {proc.stderr.strip()}"
            )

        out = proc.stdout.strip()
        if not out:
            raise RuntimeError("Rust triage helper returned empty stdout")
        return json.loads(out)


def triage_workbook(
    workbook: WorkbookInput,
    *,
    rust_exe: Path,
    password: str | None = None,
    password_file: Path | None = None,
    diff_ignore: set[str],
    diff_ignore_globs: set[str] | None = None,
    diff_ignore_path: tuple[str, ...] | list[str] = (),
    diff_ignore_path_in: tuple[str, ...] | list[str] = (),
    diff_ignore_path_kind: tuple[str, ...] | list[str] = (),
    diff_ignore_path_kind_in: tuple[str, ...] | list[str] = (),
    diff_ignore_presets: tuple[str, ...] = tuple(),
    diff_limit: int,
    round_trip_fail_on: str = "critical",
    recalc: bool,
    render_smoke: bool,
    strict_calc_chain: bool = False,
    privacy_mode: str = _PRIVACY_PUBLIC,
) -> dict[str, Any]:
    if diff_ignore_globs is None:
        diff_ignore_globs = set(DEFAULT_DIFF_IGNORE_GLOBS)

    sha = sha256_hex(workbook.data)
    display_name = workbook.display_name
    if privacy_mode == _PRIVACY_PRIVATE:
        display_name = _anonymized_display_name(sha256=sha, original_name=workbook.display_name)

    report: dict[str, Any] = {
        "display_name": display_name,
        "sha256": sha,
        "size_bytes": len(workbook.data),
        "timestamp": utc_now_iso(),
        "commit": github_commit_sha(),
        "run_url": _redact_run_url(github_run_url(), privacy_mode=privacy_mode),
    }

    # Best-effort scan for feature/function fingerprints (privacy-safe).
    try:
        with zipfile.ZipFile(io.BytesIO(workbook.data), "r") as z:
            zip_names = [info.filename for info in z.infolist() if not info.is_dir()]
            features = _scan_features(zip_names)
            report["features"] = features
            if features.get("has_cell_images") is True:
                cell_images = _extract_cell_images(z, zip_names)
                if cell_images is not None:
                    if privacy_mode == _PRIVACY_PRIVATE:
                        # Redact any non-standard/custom URIs to avoid leaking corporate domains in
                        # CI artifacts. Standard OpenXML/Microsoft schema URLs are allowlisted.
                        for k in ("content_type", "workbook_rel_type", "root_namespace"):
                            v = cell_images.get(k)
                            if isinstance(v, str) and v:
                                if k == "content_type":
                                    cell_images[k] = _redact_content_type(
                                        v, privacy_mode=privacy_mode
                                    )
                                else:
                                    cell_images[k] = _redact_uri_like(v)
                        rels_types = cell_images.get("rels_types")
                        if isinstance(rels_types, list):
                            cell_images["rels_types"] = [
                                _redact_uri_like(v) if isinstance(v, str) and v else v
                                for v in rels_types
                            ]
                    report["cell_images"] = cell_images
            report["functions"] = _redact_function_counts(
                _extract_function_counts(z), privacy_mode=privacy_mode
            )
            style_stats, style_err = _extract_style_stats(z, zip_names)
            if style_stats is not None:
                report["style_stats"] = style_stats
            elif style_err:
                if privacy_mode == _PRIVACY_PRIVATE:
                    report["style_stats_error"] = f"sha256={_sha256_text(style_err)}"
                else:
                    report["style_stats_error"] = style_err
    except Exception as e:  # noqa: BLE001 (triage tool)
        # Feature scanning failures should not leak exception strings into JSON artifacts.
        msg = str(e)
        report["features_error"] = f"sha256={hashlib.sha256(msg.encode('utf-8')).hexdigest()}"

    # Core triage (Rust): load → optional recalc/render → round-trip save → structural diff.
    try:
        rust_kwargs: dict[str, Any] = {
            "diff_ignore": diff_ignore,
            "diff_limit": diff_limit,
            "recalc": recalc,
            "render_smoke": render_smoke,
        }

        # Some unit tests patch `_run_rust_triage` with a legacy signature. Only forward optional
        # params when the target callable can accept them.
        optional_kwargs = {
            # Use the privacy-mode-adjusted display name so the Rust helper never sees raw
            # filenames in private mode. The helper only uses this for extension/format inference.
            "workbook_name": display_name,
            "password": password,
            "password_file": password_file,
            "diff_ignore_globs": diff_ignore_globs,
            "diff_ignore_presets": diff_ignore_presets,
            "round_trip_fail_on": round_trip_fail_on,
            "diff_ignore_path": diff_ignore_path,
            "diff_ignore_path_in": diff_ignore_path_in,
            "diff_ignore_path_kind": diff_ignore_path_kind,
            "diff_ignore_path_kind_in": diff_ignore_path_kind_in,
            "strict_calc_chain": strict_calc_chain,
        }
        try:
            sig = inspect.signature(_run_rust_triage)
            supports_kwargs = any(
                p.kind == inspect.Parameter.VAR_KEYWORD for p in sig.parameters.values()
            )
            for k, v in optional_kwargs.items():
                if k in sig.parameters or supports_kwargs:
                    rust_kwargs[k] = v
        except (TypeError, ValueError):
            rust_kwargs.update(optional_kwargs)

        rust_out = _run_rust_triage(rust_exe, workbook.data, **rust_kwargs)
        report["steps"] = rust_out.get("steps") or {}
        report["result"] = rust_out.get("result") or {}

        if privacy_mode == _PRIVACY_PRIVATE:
            # Defense in depth: diff paths may include expanded XML namespaces like
            # `{http://corp.example.com/ns}attr`. Redact any non-allowlisted URI-like strings.
            steps = report.get("steps")
            if isinstance(steps, dict):
                diff_step = steps.get("diff")
                if isinstance(diff_step, dict):
                    details = diff_step.get("details")
                    if isinstance(details, dict):
                        top = details.get("top_differences")
                        if isinstance(top, list):
                            for entry in top:
                                if not isinstance(entry, dict):
                                    continue
                                path = entry.get("path")
                                if isinstance(path, str) and path:
                                    entry["path"] = _redact_uri_like_in_text(path)
                        ignore_paths = details.get("ignore_paths")
                        if isinstance(ignore_paths, list):
                            # `xlsx-diff` ignore-path rules can embed custom namespaces/domains. Hash
                            # them to avoid leaking potentially sensitive strings into uploaded
                            # private-corpus artifacts.
                            details["ignore_paths"] = [
                                f"sha256={_sha256_text(p)}" if isinstance(p, str) and p else p
                                for p in ignore_paths
                            ]
    except Exception as e:  # noqa: BLE001
        report["steps"] = {"load": asdict(_step_failed(_now_ms(), e))}
        report["result"] = {
            "open_ok": False,
            "round_trip_ok": False,
            "round_trip_fail_on": round_trip_fail_on,
        }
        report["failure_category"] = "triage_error"
        return report

    res = report.get("result", {})
    if res.get("open_ok") is not True:
        report["failure_category"] = "open_error"
    elif res.get("calculate_ok") is False:
        report["failure_category"] = "calc_mismatch"
    elif res.get("render_ok") is False:
        report["failure_category"] = "render_error"
    elif res.get("round_trip_ok") is False:
        report["failure_category"] = "round_trip_diff"
        report["round_trip_failure_kind"] = infer_round_trip_failure_kind(report) or "round_trip_other"

    return report


def _compare_expectations(
    reports: list[dict[str, Any]], expectations: dict[str, Any]
) -> tuple[list[str], list[str]]:
    regressions: list[str] = []
    improvements: list[str] = []
    for r in reports:
        name = r.get("display_name")
        if not name or name not in expectations:
            continue
        exp = expectations[name]
        result = r.get("result", {})
        for key, exp_value in exp.items():
            actual = result.get(key)

            # Booleans: exact match required (treat skips as regressions).
            if isinstance(exp_value, bool):
                if actual is not exp_value:
                    regressions.append(
                        f"{name}: expected {key}={str(exp_value).lower()}, got {actual}"
                    )
                continue

            # Numbers: treat larger-than-expected as regression, smaller as improvement.
            if isinstance(exp_value, (int, float)):
                if not isinstance(actual, (int, float)):
                    regressions.append(
                        f"{name}: expected {key}={exp_value}, got {actual}"
                    )
                    continue
                if actual > exp_value:
                    regressions.append(
                        f"{name}: expected {key}<={exp_value}, got {actual}"
                    )
                elif actual < exp_value:
                    improvements.append(
                        f"{name}: expected {key}={exp_value}, got {actual}"
                    )
                continue

            # Fallback: strict equality.
            if actual != exp_value:
                regressions.append(f"{name}: expected {key}={exp_value}, got {actual}")
    return regressions, improvements


def _resolve_triage_input_dir(corpus_dir: Path, input_scope: str) -> Path:
    """Resolve which directory should be searched for workbook inputs.

    This supports the recommended private corpus layout described in
    `docs/compatibility-corpus.md`:

      tools/corpus/private/
        originals/   # encrypted originals (*.enc)
        sanitized/   # sanitized workbooks (plaintext)

    When `--input-scope auto` (default), we prefer `sanitized/` if present to
    avoid double-processing (and accidentally parsing unsanitized originals).
    """

    sanitized_dir = corpus_dir / "sanitized"
    originals_dir = corpus_dir / "originals"

    if input_scope == "all":
        return corpus_dir

    if input_scope == "auto":
        return sanitized_dir if sanitized_dir.is_dir() else corpus_dir

    if input_scope == "sanitized":
        if not sanitized_dir.is_dir():
            raise ValueError(
                f"--input-scope sanitized requested, but directory does not exist: {sanitized_dir}"
            )
        return sanitized_dir

    if input_scope == "originals":
        if not originals_dir.is_dir():
            raise ValueError(
                f"--input-scope originals requested, but directory does not exist: {originals_dir}"
            )
        return originals_dir

    raise ValueError(f"Unknown --input-scope value: {input_scope!r}")


def _triage_one_path(
    path_str: str,
    *,
    rust_exe: str,
    password: str | None = None,
    password_file: str | None = None,
    diff_ignore: tuple[str, ...],
    diff_ignore_globs: tuple[str, ...] = (),
    diff_ignore_path: tuple[str, ...] = (),
    diff_ignore_path_in: tuple[str, ...] = (),
    diff_ignore_path_kind: tuple[str, ...] = (),
    diff_ignore_path_kind_in: tuple[str, ...] = (),
    diff_ignore_presets: tuple[str, ...] = (),
    diff_limit: int,
    round_trip_fail_on: str = "critical",
    recalc: bool,
    render_smoke: bool,
    strict_calc_chain: bool = False,
    leak_scan: bool,
    fernet_key: str | None,
    privacy_mode: str = _PRIVACY_PUBLIC,
) -> dict[str, Any] | LeakScanFailure:
    """Worker entrypoint for triaging a single workbook path.

    Note: Keep this function top-level and pickleable; it may run in a ProcessPoolExecutor.
    """

    path = Path(path_str)
    try:
        wb = read_workbook_input(path, fernet_key=fernet_key)
        if leak_scan:
            scan = scan_xlsx_bytes_for_leaks(wb.data)
            if not scan.ok:
                display_name = wb.display_name
                if privacy_mode == _PRIVACY_PRIVATE:
                    display_name = _anonymized_display_name(
                        sha256=sha256_hex(wb.data), original_name=wb.display_name
                    )
                return LeakScanFailure(
                    display_name=display_name,
                    findings=[
                        {
                            "kind": f.kind,
                            "part_name": f.part_name,
                            "match_sha256": f.match_sha256,
                        }
                        for f in scan.findings[:25]
                    ],
                )

        return triage_workbook(
            wb,
            rust_exe=Path(rust_exe),
            password=password,
            password_file=Path(password_file) if password_file else None,
            diff_ignore=set(diff_ignore),
            diff_ignore_globs=set(diff_ignore_globs),
            diff_ignore_path=diff_ignore_path,
            diff_ignore_path_in=diff_ignore_path_in,
            diff_ignore_path_kind=diff_ignore_path_kind,
            diff_ignore_path_kind_in=diff_ignore_path_kind_in,
            diff_ignore_presets=diff_ignore_presets,
            diff_limit=diff_limit,
            round_trip_fail_on=round_trip_fail_on,
            recalc=recalc,
            render_smoke=render_smoke,
            strict_calc_chain=strict_calc_chain,
            privacy_mode=privacy_mode,
        )
    except Exception as e:  # noqa: BLE001
        sha: str | None
        try:
            sha = sha256_hex(path.read_bytes())
        except Exception:  # noqa: BLE001
            # Fall back to hashing the path itself so report file naming doesn't need to re-read
            # the (possibly unreadable) input.
            sha = _sha256_text(str(path))

        display_name = path.name
        if privacy_mode == _PRIVACY_PRIVATE:
            display_name = _anonymized_display_name(sha256=sha, original_name=path.name)
        return {
            "display_name": display_name,
            "sha256": sha,
            "timestamp": utc_now_iso(),
            "commit": github_commit_sha(),
            "run_url": _redact_run_url(github_run_url(), privacy_mode=privacy_mode),
            "steps": {"load": asdict(_step_failed(_now_ms(), e))},
            "result": {"open_ok": False},
            "failure_category": "triage_error",
        }


def _report_id_for_report(report: dict[str, Any], *, path: Path) -> str:
    """Stable report ID (used for indexing and file naming).

    Prefer the workbook content SHA from the report itself, but fall back to hashing the on-disk
    bytes (e.g. for unreadable/encrypted inputs that failed before hashing).
    """

    sha = report.get("sha256")
    if isinstance(sha, str) and sha:
        return sha[:16]
    return sha256_hex(path.read_bytes())[:16]


def _display_name_for_report(report: dict[str, Any], *, path: Path, privacy_mode: str) -> str:
    name = report.get("display_name")
    if isinstance(name, str) and name:
        return name
    if privacy_mode == _PRIVACY_PRIVATE:
        sha = report.get("sha256")
        if not isinstance(sha, str) or not sha:
            try:
                sha = sha256_hex(path.read_bytes())
            except Exception:  # noqa: BLE001
                sha = _sha256_text(str(path))
        return _anonymized_display_name(sha256=sha, original_name=path.name)
    return path.name


def _report_filename_for_path(
    report: dict[str, Any], *, path: Path, corpus_dir: Path
) -> str:
    """Deterministic, non-colliding report filename for a workbook path."""

    report_id = _report_id_for_report(report, path=path)
    try:
        rel = path.relative_to(corpus_dir).as_posix()
    except Exception:  # noqa: BLE001
        rel = path.name
    # Include a hash of the workbook's relative path to avoid filename collisions when the same
    # workbook content appears multiple times in the corpus. Use a 64-bit prefix (16 hex chars)
    # to make collisions vanishingly unlikely even in very large corpora.
    path_hash = sha256_hex(rel.encode("utf-8"))[:16]
    return f"{report_id}-{path_hash}.json"


def _triage_paths(
    paths: list[Path],
    *,
    rust_exe: str,
    password: str | None = None,
    password_file: str | None = None,
    diff_ignore: set[str],
    diff_ignore_globs: set[str] | None = None,
    diff_ignore_path: tuple[str, ...] = (),
    diff_ignore_path_in: tuple[str, ...] = (),
    diff_ignore_path_kind: tuple[str, ...] = (),
    diff_ignore_path_kind_in: tuple[str, ...] = (),
    diff_ignore_presets: tuple[str, ...] = (),
    diff_limit: int,
    round_trip_fail_on: str = "critical",
    recalc: bool,
    render_smoke: bool,
    strict_calc_chain: bool = False,
    leak_scan: bool,
    fernet_key: str | None,
    jobs: int,
    privacy_mode: str = _PRIVACY_PUBLIC,
    executor_cls: type[concurrent.futures.Executor] | None = None,
) -> list[dict[str, Any]] | LeakScanFailure:
    """Run triage over a list of workbook paths, possibly in parallel.

    Returns either:
    - ordered list of reports (matching `paths` ordering), or
    - a LeakScanFailure (fail-fast signal for --leak-scan).
    """

    if jobs < 1:
        jobs = 1

    # Clamp to the number of workbooks so we don't spin up unnecessary worker processes for small
    # corpora (and so downstream CPU tuning can use the effective worker count).
    worker_count = min(jobs, len(paths)) if paths else 1

    reports_by_index: list[dict[str, Any] | None] = [None] * len(paths)
    diff_ignore_tuple = tuple(sorted({p for p in diff_ignore if p}))
    if diff_ignore_globs is None:
        diff_ignore_globs = set()
    diff_ignore_globs_tuple = tuple(sorted({p for p in diff_ignore_globs if p}))
    diff_ignore_path_tuple = tuple(sorted({p for p in diff_ignore_path if p and p.strip()}))
    diff_ignore_path_in_tuple = tuple(sorted({p for p in diff_ignore_path_in if p and p.strip()}))
    diff_ignore_path_kind_tuple = tuple(
        sorted({p for p in diff_ignore_path_kind if p and p.strip()})
    )
    diff_ignore_path_kind_in_tuple = tuple(
        sorted({p for p in diff_ignore_path_kind_in if p and p.strip()})
    )
    diff_ignore_presets_tuple = tuple(sorted({p for p in diff_ignore_presets if p and p.strip()}))

    if worker_count == 1 or len(paths) <= 1:
        for idx, path in enumerate(paths):
            res = _triage_one_path(
                str(path),
                rust_exe=rust_exe,
                password=password,
                password_file=password_file,
                diff_ignore=diff_ignore_tuple,
                diff_ignore_globs=diff_ignore_globs_tuple,
                diff_ignore_path=diff_ignore_path_tuple,
                diff_ignore_path_in=diff_ignore_path_in_tuple,
                diff_ignore_path_kind=diff_ignore_path_kind_tuple,
                diff_ignore_path_kind_in=diff_ignore_path_kind_in_tuple,
                diff_ignore_presets=diff_ignore_presets_tuple,
                diff_limit=diff_limit,
                round_trip_fail_on=round_trip_fail_on,
                recalc=recalc,
                render_smoke=render_smoke,
                strict_calc_chain=strict_calc_chain,
                leak_scan=leak_scan,
                fernet_key=fernet_key,
                privacy_mode=privacy_mode,
            )
            if isinstance(res, LeakScanFailure):
                return res
            reports_by_index[idx] = res
        return [r for r in reports_by_index if r is not None]

    executor_cls = executor_cls or concurrent.futures.ProcessPoolExecutor
    done = 0
    # Large private corpora can contain thousands of workbooks. Submitting one Future per workbook
    # up-front creates unnecessary memory overhead. Instead, keep a bounded number of tasks
    # in-flight and feed the executor as work completes.
    prefetch = max(worker_count * 2, 1)
    with executor_cls(max_workers=worker_count) as executor:
        pending: dict[concurrent.futures.Future[dict[str, Any] | LeakScanFailure], int] = {}
        path_iter = iter(enumerate(paths))

        def _submit_next() -> None:
            try:
                idx, path = next(path_iter)
            except StopIteration:
                return
            fut = executor.submit(
                _triage_one_path,
                str(path),
                rust_exe=rust_exe,
                password=password,
                password_file=password_file,
                diff_ignore=diff_ignore_tuple,
                diff_ignore_globs=diff_ignore_globs_tuple,
                diff_ignore_path=diff_ignore_path_tuple,
                diff_ignore_path_in=diff_ignore_path_in_tuple,
                diff_ignore_path_kind=diff_ignore_path_kind_tuple,
                diff_ignore_path_kind_in=diff_ignore_path_kind_in_tuple,
                diff_ignore_presets=diff_ignore_presets_tuple,
                diff_limit=diff_limit,
                round_trip_fail_on=round_trip_fail_on,
                recalc=recalc,
                render_smoke=render_smoke,
                strict_calc_chain=strict_calc_chain,
                leak_scan=leak_scan,
                fernet_key=fernet_key,
                privacy_mode=privacy_mode,
            )
            pending[fut] = idx

        for _ in range(min(prefetch, len(paths))):
            _submit_next()

        while pending:
            finished, _ = concurrent.futures.wait(
                pending, return_when=concurrent.futures.FIRST_COMPLETED
            )
            for fut in finished:
                idx = pending.pop(fut)
                res = fut.result()
                if isinstance(res, LeakScanFailure):
                    # Best-effort fail-fast: cancel any work we haven't started yet. We intentionally
                    # keep this conservative (no process killing) to avoid leaving behind corrupted
                    # state. With bounded prefetch, the worst-case wait is a small multiple of
                    # `jobs` instead of `len(paths)`.
                    for other in pending:
                        other.cancel()
                    return res
                reports_by_index[idx] = res
                done += 1
                # Keep logs readable: one progress line per workbook (printed by parent only).
                name_for_log = _display_name_for_report(
                    res, path=paths[idx], privacy_mode=privacy_mode
                )
                print(
                    f"[{done}/{len(paths)}] triaged {name_for_log}"
                )
                _submit_next()

    return [r for r in reports_by_index if r is not None]


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run corpus workbook triage.")
    parser.add_argument("--corpus-dir", type=Path, required=True)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument(
        "--jobs",
        type=int,
        default=1,
        help="Number of parallel triage workers to use (default: 1).",
    )
    parser.add_argument(
        "--include-xlsb",
        action="store_true",
        help="Also include `.xlsb` workbooks in corpus scans (off by default).",
    )
    parser.add_argument(
        "--leak-scan",
        action="store_true",
        help="Fail fast if any workbook contains obvious plaintext PII/secrets (emails, URLs, keys).",
    )
    parser.add_argument(
        "--expectations",
        type=Path,
        help="Optional JSON mapping display_name -> expected result fields (for regression gating).",
    )
    parser.add_argument(
        "--privacy-mode",
        choices=[_PRIVACY_PUBLIC, _PRIVACY_PRIVATE],
        default=_PRIVACY_PUBLIC,
        help=(
            "Control redaction of triage reports. "
            "`public` preserves filenames/URIs; `private` anonymizes display_name and hashes "
            "custom URI-like strings (including expanded XML namespaces in diff paths) and "
            "non-standard/custom formula function names. "
            "In private mode, CI metadata like run_url and local paths in index.json are also hashed."
        ),
    )
    parser.add_argument(
        "--input-scope",
        choices=("auto", "sanitized", "originals", "all"),
        default="auto",
        help=(
            "Select which corpus inputs to triage. "
            "auto (default) prefers <corpus-dir>/sanitized if present, otherwise triages <corpus-dir>. "
            "sanitized/originals triage only those subdirectories. "
            "all triages the full <corpus-dir> tree (may double-count in private corpus layouts)."
        ),
    )
    parser.add_argument(
        "--fernet-key-env",
        default="CORPUS_ENCRYPTION_KEY",
        help="Env var containing Fernet key used to decrypt *.enc corpus files.",
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
            "Enables triaging Office-encrypted XLSX/XLSM/XLSB workbooks."
        ),
    )
    parser.add_argument(
        "--recalc",
        action="store_true",
        help="Enable best-effort recalculation correctness check (off by default).",
    )
    parser.add_argument(
        "--render-smoke",
        action="store_true",
        help="Enable lightweight headless render/print smoke test (off by default).",
    )
    parser.add_argument(
        "--diff-ignore",
        action="append",
        default=[],
        help="Additional part path to ignore during diff (can be repeated).",
    )
    parser.add_argument(
        "--diff-ignore-glob",
        action="append",
        default=[],
        help="Additional glob pattern to ignore during diff (can be repeated).",
    )
    parser.add_argument(
        "--diff-ignore-path",
        action="append",
        default=[],
        help=(
            "Substring pattern to ignore within XML diff paths (can be repeated). "
            "Useful for suppressing noisy attributes like dyDescent or xr:uid."
        ),
    )
    parser.add_argument(
        "--diff-ignore-path-in",
        action="append",
        default=[],
        help=(
            "Like --diff-ignore-path, but scoped to parts matched by a glob (can be repeated). "
            "Format: <part_glob>:<path_substring>."
        ),
    )
    parser.add_argument(
        "--diff-ignore-preset",
        action="append",
        default=[],
        help="Pass an xlsx-diff ignore preset through to the diff step (can be repeated).",
    )
    parser.add_argument(
        "--diff-ignore-path-kind",
        action="append",
        default=[],
        help=(
            "Like --diff-ignore-path, but only applies to diffs whose kind matches the provided kind "
            "(can be repeated). Format: <kind>:<path_substring>."
        ),
    )
    parser.add_argument(
        "--diff-ignore-path-kind-in",
        action="append",
        default=[],
        help=(
            "Like --diff-ignore-path-in, but only applies to diffs whose kind matches the provided kind "
            "(can be repeated). Format: <part_glob>:<kind>:<path_substring>."
        ),
    )
    parser.add_argument(
        "--no-default-diff-ignore",
        action="store_true",
        help=(
            "Disable the built-in default ignore list (docProps/*). "
            "When set, only explicit --diff-ignore/--diff-ignore-glob entries are ignored."
        ),
    )
    parser.add_argument(
        "--strict-calc-chain",
        action="store_true",
        help=(
            "Treat calcChain-related diffs as CRITICAL instead of being downgraded to WARN "
            "(useful for strict round-trip preservation scoring)."
        ),
    )
    parser.add_argument(
        "--diff-limit",
        type=int,
        default=25,
        help="Maximum number of diff entries to include in reports (privacy-safe).",
    )
    parser.add_argument(
        "--round-trip-fail-on",
        default="critical",
        choices=["critical", "warning", "info"],
        help="Round-trip diff severity threshold that should fail `round_trip_ok`.",
    )
    return parser


def main() -> int:
    parser = _build_arg_parser()
    args = parser.parse_args()

    if args.jobs < 1:
        parser.error("--jobs must be >= 1")

    password_file: Path | None = None
    if args.password_file:
        password_file = args.password_file.expanduser()
        if not password_file.is_file():
            parser.error(f"--password-file does not exist or is not a file: {password_file}")
        # The Rust helper is invoked with `cwd=_repo_root()`, so resolve to an absolute path to
        # avoid surprising relative-path behavior.
        password_file = password_file.resolve()

    diff_ignore = _compute_diff_ignore(
        diff_ignore=args.diff_ignore, use_default=not args.no_default_diff_ignore
    )
    diff_ignore_globs = _compute_diff_ignore_globs(
        diff_ignore_globs=args.diff_ignore_glob, use_default=not args.no_default_diff_ignore
    )
    diff_ignore_presets = tuple(
        sorted({(p or "").strip() for p in args.diff_ignore_preset if (p or "").strip()})
    )

    try:
        input_dir = _resolve_triage_input_dir(args.corpus_dir, args.input_scope)
    except ValueError as e:
        parser.error(str(e))

    ensure_dir(args.out_dir)
    reports_dir = args.out_dir / "reports"
    ensure_dir(reports_dir)

    fernet_key = os.environ.get(args.fernet_key_env)
    if args.input_scope == "originals" and not fernet_key:
        parser.error(
            f"--input-scope originals requires decryption; set ${args.fernet_key_env} or pass --fernet-key-env"
        )

    rust_exe = _build_rust_helper()

    paths = list(iter_workbook_paths(input_dir, include_xlsb=args.include_xlsb))

    # When running multiple Rust helpers in parallel, cap Rayon parallelism per-process to avoid
    # accidental CPU oversubscription (each helper would otherwise default to all cores). Clamp to
    # the effective worker count so small corpora don't artificially cap Rayon parallelism.
    effective_workers = min(args.jobs, len(paths)) if paths else 1
    if effective_workers > 1 and not os.environ.get("RAYON_NUM_THREADS"):
        cpu = os.cpu_count() or 1
        os.environ["RAYON_NUM_THREADS"] = str(max(1, cpu // effective_workers))

    # Record path-ignore patterns (if any) in the output metadata. In private mode, hash the
    # patterns to avoid leaking custom namespaces/domains in uploaded artifacts.
    diff_ignore_path_values = sorted(
        {p.strip() for p in (args.diff_ignore_path or []) if p and p.strip()}
    )
    diff_ignore_path_in_values = sorted(
        {p.strip() for p in (args.diff_ignore_path_in or []) if p and p.strip()}
    )
    diff_ignore_path_kind_values = sorted(
        {p.strip() for p in (args.diff_ignore_path_kind or []) if p and p.strip()}
    )
    diff_ignore_path_kind_in_values = sorted(
        {p.strip() for p in (args.diff_ignore_path_kind_in or []) if p and p.strip()}
    )
    if args.privacy_mode == _PRIVACY_PRIVATE:
        diff_ignore_path_out = [f"sha256={_sha256_text(p)}" for p in diff_ignore_path_values]
        diff_ignore_path_in_out = [
            f"sha256={_sha256_text(p)}" for p in diff_ignore_path_in_values
        ]
        diff_ignore_path_kind_out = [
            f"sha256={_sha256_text(p)}" for p in diff_ignore_path_kind_values
        ]
        diff_ignore_path_kind_in_out = [
            f"sha256={_sha256_text(p)}" for p in diff_ignore_path_kind_in_values
        ]
    else:
        diff_ignore_path_out = diff_ignore_path_values
        diff_ignore_path_in_out = diff_ignore_path_in_values
        diff_ignore_path_kind_out = diff_ignore_path_kind_values
        diff_ignore_path_kind_in_out = diff_ignore_path_kind_in_values

    rayon_threads_out: int | None = None
    rayon_raw = os.environ.get("RAYON_NUM_THREADS")
    if rayon_raw:
        try:
            rayon_threads_out = int(rayon_raw)
        except ValueError:
            rayon_threads_out = None
    triage_out = _triage_paths(
        paths,
        rust_exe=str(rust_exe),
        password=args.password,
        password_file=str(password_file) if password_file else None,
        diff_ignore=diff_ignore,
        diff_ignore_globs=diff_ignore_globs,
        diff_ignore_path=tuple(diff_ignore_path_values),
        diff_ignore_path_in=tuple(diff_ignore_path_in_values),
        diff_ignore_path_kind=tuple(diff_ignore_path_kind_values),
        diff_ignore_path_kind_in=tuple(diff_ignore_path_kind_in_values),
        diff_ignore_presets=diff_ignore_presets,
        diff_limit=args.diff_limit,
        round_trip_fail_on=args.round_trip_fail_on,
        recalc=args.recalc,
        render_smoke=args.render_smoke,
        strict_calc_chain=bool(args.strict_calc_chain),
        leak_scan=args.leak_scan,
        fernet_key=fernet_key,
        jobs=args.jobs,
        privacy_mode=args.privacy_mode,
    )

    if isinstance(triage_out, LeakScanFailure):
        print(
            f"LEAKS DETECTED in {triage_out.display_name} ({len(triage_out.findings)} findings)"
        )
        for f in triage_out.findings:
            print(f"  {f['kind']} in {f['part_name']} sha256={f['match_sha256'][:16]}")
        return 1

    reports = triage_out
    report_index_entries: list[dict[str, str]] = []
    for path, report in zip(paths, reports, strict=True):
        report_id = _report_id_for_report(report, path=path)
        filename = _report_filename_for_path(report, path=path, corpus_dir=args.corpus_dir)
        write_json(reports_dir / filename, report)
        report_index_entries.append(
            {
                "id": report_id,
                "display_name": _display_name_for_report(
                    report, path=path, privacy_mode=args.privacy_mode
                ),
                "file": filename,
            }
        )

    corpus_dir_str = str(args.corpus_dir)
    input_dir_str = str(input_dir)
    if args.privacy_mode == _PRIVACY_PRIVATE:
        # Avoid leaking local filesystem paths (usernames, mount points) into uploaded artifacts.
        corpus_dir_str = f"sha256={_sha256_text(corpus_dir_str)}"
        input_dir_str = f"sha256={_sha256_text(input_dir_str)}"

    index = {
        "timestamp": utc_now_iso(),
        "commit": github_commit_sha(),
        "run_url": _redact_run_url(github_run_url(), privacy_mode=args.privacy_mode),
        "corpus_dir": corpus_dir_str,
        "input_scope": args.input_scope,
        "input_dir": input_dir_str,
        "jobs": args.jobs,
        "jobs_effective": effective_workers,
        "privacy_mode": args.privacy_mode,
        "include_xlsb": args.include_xlsb,
        "recalc": args.recalc,
        "render_smoke": args.render_smoke,
        "diff_limit": args.diff_limit,
        "round_trip_fail_on": args.round_trip_fail_on,
        "diff_ignore": sorted(diff_ignore),
        "diff_ignore_globs": sorted(diff_ignore_globs),
        "diff_ignore_presets": list(diff_ignore_presets),
        "diff_ignore_path": diff_ignore_path_out,
        "diff_ignore_path_in": diff_ignore_path_in_out,
        "diff_ignore_path_kind": diff_ignore_path_kind_out,
        "diff_ignore_path_kind_in": diff_ignore_path_kind_in_out,
        "no_default_diff_ignore": args.no_default_diff_ignore,
        "strict_calc_chain": bool(args.strict_calc_chain),
        "rayon_num_threads": rayon_threads_out,
        "report_count": len(reports),
        "reports": report_index_entries,
    }
    write_json(args.out_dir / "index.json", index)

    regressions: list[str] = []
    improvements: list[str] = []
    # Expectations gating is intended for the public corpus. When triaging a scoped
    # subdirectory (e.g. private corpus `sanitized/`), skip expectations to avoid
    # accidental regression gating on private/sensitive datasets.
    if (
        args.expectations
        and args.expectations.exists()
        and input_dir.resolve() == args.corpus_dir.resolve()
    ):
        expectations = json.loads(args.expectations.read_text(encoding="utf-8"))
        regressions, improvements = _compare_expectations(reports, expectations)
        write_json(
            args.out_dir / "expectations-result.json",
            {"regressions": regressions, "improvements": improvements},
        )

    # Fail fast if a regression is detected; that is the primary CI signal.
    if regressions:
        for r in regressions:
            print(f"REGRESSION: {r}")
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
