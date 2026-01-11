#!/usr/bin/env python3

from __future__ import annotations

import argparse
import io
import json
import os
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
    # Many apps rewrite calcChain as part of recalculation.
    "xl/calcChain.xml",
}


@dataclass(frozen=True)
class StepResult:
    status: str  # ok | failed | skipped
    duration_ms: int | None = None
    error: str | None = None
    details: dict[str, Any] | None = None


def _now_ms() -> float:
    return time.perf_counter() * 1000.0


def _step_ok(start_ms: float, *, details: dict[str, Any] | None = None) -> StepResult:
    return StepResult(status="ok", duration_ms=int(_now_ms() - start_ms), details=details)


def _step_failed(start_ms: float, err: Exception) -> StepResult:
    return StepResult(
        status="failed", duration_ms=int(_now_ms() - start_ms), error=str(err)
    )


def _step_skipped(reason: str) -> StepResult:
    return StepResult(status="skipped", details={"reason": reason})


def _scan_features(zip_names: list[str]) -> dict[str, Any]:
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
        features[f"has_{key}"] = any(n.startswith(prefix) for n in zip_names)

    features["has_vba"] = "xl/vbaProject.bin" in zip_names
    features["has_connections"] = "xl/connections.xml" in zip_names
    features["has_shared_strings"] = "xl/sharedStrings.xml" in zip_names
    features["sheet_xml_count"] = len([n for n in zip_names if n.startswith("xl/worksheets/sheet")])
    return features


def _extract_function_counts(z: zipfile.ZipFile) -> Counter[str]:
    """Return function name -> usage count (no formulas or values are emitted)."""

    import re

    fn_re = re.compile(r"(?i)\b([A-Z_][A-Z0-9_.]*)\s*\(")
    string_lit_re = re.compile(r'"(?:[^"]|"")*"')

    counts: Counter[str] = Counter()
    for name in z.namelist():
        if not (name.startswith("xl/worksheets/") and name.endswith(".xml")):
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


def _repo_root() -> Path:
    # tools/corpus/triage.py -> tools/corpus -> tools -> repo root
    return Path(__file__).resolve().parents[2]


def _rust_exe_name() -> str:
    return "formula-corpus-triage.exe" if os.name == "nt" else "formula-corpus-triage"


def _build_rust_helper() -> Path:
    """Build (or reuse) the Rust triage helper binary."""

    root = _repo_root()
    env = os.environ.copy()
    env.setdefault("CARGO_HOME", str(root / "target" / "cargo-home"))
    Path(env["CARGO_HOME"]).mkdir(parents=True, exist_ok=True)
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
        subprocess.run(
            ["cargo", "build", "-p", "formula-corpus-triage"],
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
            report["features"] = _scan_features(zip_names)
            report["functions"] = dict(_extract_function_counts(z))
    except Exception as e:  # noqa: BLE001 (triage tool)
        report["features_error"] = str(e)

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


def main() -> int:
    parser = argparse.ArgumentParser(description="Run corpus workbook triage.")
    parser.add_argument("--corpus-dir", type=Path, required=True)
    parser.add_argument("--out-dir", type=Path, required=True)
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

    rust_exe = _build_rust_helper()
    diff_ignore = set(DEFAULT_DIFF_IGNORE) | {p for p in args.diff_ignore if p}

    ensure_dir(args.out_dir)
    reports_dir = args.out_dir / "reports"
    ensure_dir(reports_dir)

    fernet_key = os.environ.get(args.fernet_key_env)
    reports: list[dict[str, Any]] = []
    for path in iter_workbook_paths(args.corpus_dir):
        try:
            wb = read_workbook_input(path, fernet_key=fernet_key)
            if args.leak_scan:
                scan = scan_xlsx_bytes_for_leaks(wb.data)
                if not scan.ok:
                    print(f"LEAKS DETECTED in {path.name} ({len(scan.findings)} findings)")
                    for f in scan.findings[:25]:
                        print(f"  {f.kind} in {f.part_name} sha256={f.match_sha256[:16]}")
                    return 1
            report = triage_workbook(
                wb,
                rust_exe=rust_exe,
                diff_ignore=diff_ignore,
                diff_limit=args.diff_limit,
                recalc=args.recalc,
                render_smoke=args.render_smoke,
            )
        except Exception as e:  # noqa: BLE001
            report = {
                "display_name": path.name,
                "timestamp": utc_now_iso(),
                "commit": github_commit_sha(),
                "run_url": github_run_url(),
                "steps": {"load": asdict(_step_failed(_now_ms(), e))},
                "result": {"open_ok": False},
                "failure_category": "triage_error",
            }

        reports.append(report)
        report_id = (report.get("sha256") or sha256_hex(path.read_bytes()))[:16]
        write_json(reports_dir / f"{report_id}.json", report)

    index = {
        "timestamp": utc_now_iso(),
        "commit": github_commit_sha(),
        "run_url": github_run_url(),
        "corpus_dir": str(args.corpus_dir),
        "report_count": len(reports),
        "reports": [
            {"id": r.get("sha256", "")[:16], "display_name": r.get("display_name")}
            for r in reports
        ],
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
