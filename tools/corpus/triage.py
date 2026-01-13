#!/usr/bin/env python3

from __future__ import annotations

import argparse
import concurrent.futures
import hashlib
import io
import json
import os
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
    # NOTE: calcChain (`xl/calcChain.xml`) is intentionally *not* ignored by default.
    # `xlsx-diff` downgrades calcChain-related churn to WARNING so it shows up in metrics/dashboards
    # without failing CI gates (which key off critical diffs).
}


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


def _now_ms() -> float:
    return time.perf_counter() * 1000.0


def _step_ok(start_ms: float, *, details: dict[str, Any] | None = None) -> StepResult:
    return StepResult(status="ok", duration_ms=int(_now_ms() - start_ms), details=details)


def _step_failed(start_ms: float, err: Exception) -> StepResult:
    # Triage reports are uploaded as artifacts for both public and private corpora. Avoid leaking
    # workbook content (sheet names, defined names, file paths, etc.) through exception strings by
    # hashing the message and emitting only the digest (mirrors the Rust helper behavior).
    msg = str(err)
    digest = hashlib.sha256(msg.encode("utf-8")).hexdigest()
    return StepResult(
        status="failed", duration_ms=int(_now_ms() - start_ms), error=f"sha256={digest}"
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
    default_global_cargo_home = Path.home() / ".cargo"
    cargo_home = env.get("CARGO_HOME")
    cargo_home_path = Path(cargo_home).expanduser() if cargo_home else None
    if not cargo_home or (
        not env.get("CI")
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
    except FileNotFoundError as e:  # noqa: PERF203 (CI signal)
        raise RuntimeError("cargo not found; Rust toolchain is required for corpus triage") from e

    if not exe.exists():
        raise RuntimeError(f"Rust triage helper was built but executable is missing: {exe}")
    return exe


def _run_rust_triage(
    exe: Path,
    workbook_bytes: bytes,
    *,
    diff_ignore: set[str],
    diff_limit: int,
    recalc: bool,
    render_smoke: bool,
) -> dict[str, Any]:
    """Invoke the Rust helper to run load/save/diff (+ optional recalc/render) on a workbook blob."""

    with tempfile.TemporaryDirectory(prefix="corpus-triage-") as tmpdir:
        input_path = Path(tmpdir) / "input.xlsx"
        input_path.write_bytes(workbook_bytes)

        cmd = [
            str(exe),
            "--input",
            str(input_path),
            "--diff-limit",
            str(diff_limit),
        ]
        for part in sorted(diff_ignore):
            cmd.extend(["--ignore-part", part])
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
    diff_ignore: set[str],
    diff_limit: int,
    recalc: bool,
    render_smoke: bool,
) -> dict[str, Any]:
    report: dict[str, Any] = {
        "display_name": workbook.display_name,
        "sha256": sha256_hex(workbook.data),
        "size_bytes": len(workbook.data),
        "timestamp": utc_now_iso(),
        "commit": github_commit_sha(),
        "run_url": github_run_url(),
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
                    report["cell_images"] = cell_images
            report["functions"] = dict(_extract_function_counts(z))
            style_stats, style_err = _extract_style_stats(z, zip_names)
            if style_stats is not None:
                report["style_stats"] = style_stats
            elif style_err:
                report["style_stats_error"] = style_err
    except Exception as e:  # noqa: BLE001 (triage tool)
        # Feature scanning failures should not leak exception strings into JSON artifacts.
        msg = str(e)
        report["features_error"] = f"sha256={hashlib.sha256(msg.encode('utf-8')).hexdigest()}"

    # Core triage (Rust): load → optional recalc/render → round-trip save → structural diff.
    try:
        rust_out = _run_rust_triage(
            rust_exe,
            workbook.data,
            diff_ignore=diff_ignore,
            diff_limit=diff_limit,
            recalc=recalc,
            render_smoke=render_smoke,
        )
        report["steps"] = rust_out.get("steps") or {}
        report["result"] = rust_out.get("result") or {}
    except Exception as e:  # noqa: BLE001
        report["steps"] = {"load": asdict(_step_failed(_now_ms(), e))}
        report["result"] = {"open_ok": False, "round_trip_ok": False}
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


def _triage_one_path(
    path_str: str,
    *,
    rust_exe: str,
    diff_ignore: tuple[str, ...],
    diff_limit: int,
    recalc: bool,
    render_smoke: bool,
    leak_scan: bool,
    fernet_key: str | None,
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
                return LeakScanFailure(
                    display_name=wb.display_name,
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
            diff_ignore=set(diff_ignore),
            diff_limit=diff_limit,
            recalc=recalc,
            render_smoke=render_smoke,
        )
    except Exception as e:  # noqa: BLE001
        return {
            "display_name": path.name,
            "timestamp": utc_now_iso(),
            "commit": github_commit_sha(),
            "run_url": github_run_url(),
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


def _report_filename_for_path(
    report: dict[str, Any], *, path: Path, corpus_dir: Path
) -> str:
    """Deterministic, non-colliding report filename for a workbook path."""

    report_id = _report_id_for_report(report, path=path)
    try:
        rel = path.relative_to(corpus_dir).as_posix()
    except Exception:  # noqa: BLE001
        rel = path.name
    path_hash = sha256_hex(rel.encode("utf-8"))[:8]
    return f"{report_id}-{path_hash}.json"


def _triage_paths(
    paths: list[Path],
    *,
    rust_exe: str,
    diff_ignore: set[str],
    diff_limit: int,
    recalc: bool,
    render_smoke: bool,
    leak_scan: bool,
    fernet_key: str | None,
    jobs: int,
    executor_cls: type[concurrent.futures.Executor] | None = None,
) -> list[dict[str, Any]] | LeakScanFailure:
    """Run triage over a list of workbook paths, possibly in parallel.

    Returns either:
    - ordered list of reports (matching `paths` ordering), or
    - a LeakScanFailure (fail-fast signal for --leak-scan).
    """

    if jobs < 1:
        jobs = 1

    reports_by_index: list[dict[str, Any] | None] = [None] * len(paths)
    diff_ignore_tuple = tuple(sorted({p for p in diff_ignore if p}))

    if jobs == 1 or len(paths) <= 1:
        for idx, path in enumerate(paths):
            res = _triage_one_path(
                str(path),
                rust_exe=rust_exe,
                diff_ignore=diff_ignore_tuple,
                diff_limit=diff_limit,
                recalc=recalc,
                render_smoke=render_smoke,
                leak_scan=leak_scan,
                fernet_key=fernet_key,
            )
            if isinstance(res, LeakScanFailure):
                return res
            reports_by_index[idx] = res
        return [r for r in reports_by_index if r is not None]

    executor_cls = executor_cls or concurrent.futures.ProcessPoolExecutor
    done = 0
    with executor_cls(max_workers=jobs) as executor:
        future_to_index: dict[concurrent.futures.Future[dict[str, Any] | LeakScanFailure], int] = {}
        for idx, path in enumerate(paths):
            fut = executor.submit(
                _triage_one_path,
                str(path),
                rust_exe=rust_exe,
                diff_ignore=diff_ignore_tuple,
                diff_limit=diff_limit,
                recalc=recalc,
                render_smoke=render_smoke,
                leak_scan=leak_scan,
                fernet_key=fernet_key,
            )
            future_to_index[fut] = idx

        for fut in concurrent.futures.as_completed(future_to_index):
            idx = future_to_index[fut]
            res = fut.result()
            if isinstance(res, LeakScanFailure):
                return res
            reports_by_index[idx] = res
            done += 1
            # Keep logs readable: one progress line per workbook (printed by parent only).
            print(f"[{done}/{len(paths)}] triaged {res.get('display_name', paths[idx].name)}")

    return [r for r in reports_by_index if r is not None]


def main() -> int:
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
        "--fernet-key-env",
        default="CORPUS_ENCRYPTION_KEY",
        help="Env var containing Fernet key used to decrypt *.enc corpus files.",
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
        "--diff-limit",
        type=int,
        default=25,
        help="Maximum number of diff entries to include in reports (privacy-safe).",
    )
    args = parser.parse_args()

    if args.jobs < 1:
        parser.error("--jobs must be >= 1")

    rust_exe = _build_rust_helper()
    diff_ignore = set(DEFAULT_DIFF_IGNORE) | {p for p in args.diff_ignore if p}

    ensure_dir(args.out_dir)
    reports_dir = args.out_dir / "reports"
    ensure_dir(reports_dir)

    fernet_key = os.environ.get(args.fernet_key_env)
    paths = list(iter_workbook_paths(args.corpus_dir))
    triage_out = _triage_paths(
        paths,
        rust_exe=str(rust_exe),
        diff_ignore=diff_ignore,
        diff_limit=args.diff_limit,
        recalc=args.recalc,
        render_smoke=args.render_smoke,
        leak_scan=args.leak_scan,
        fernet_key=fernet_key,
        jobs=args.jobs,
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
            {"id": report_id, "display_name": report.get("display_name", path.name), "file": filename}
        )

    index = {
        "timestamp": utc_now_iso(),
        "commit": github_commit_sha(),
        "run_url": github_run_url(),
        "corpus_dir": str(args.corpus_dir),
        "report_count": len(reports),
        "reports": report_index_entries,
    }
    write_json(args.out_dir / "index.json", index)

    regressions: list[str] = []
    improvements: list[str] = []
    if args.expectations and args.expectations.exists():
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
