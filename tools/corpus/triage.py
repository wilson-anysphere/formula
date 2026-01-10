#!/usr/bin/env python3

from __future__ import annotations

import argparse
import io
import json
import os
import time
import zipfile
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

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


def _diff_workbooks(a: bytes, b: bytes, *, ignore: set[str] | None = None) -> dict[str, Any]:
    if ignore is None:
        ignore = set(DEFAULT_DIFF_IGNORE)

    def parts(blob: bytes) -> dict[str, str]:
        out: dict[str, str] = {}
        with zipfile.ZipFile(io.BytesIO(blob), "r") as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                if info.filename in ignore:
                    continue
                out[info.filename] = sha256_hex(z.read(info.filename))
        return out

    a_parts = parts(a)
    b_parts = parts(b)

    a_names = set(a_parts.keys())
    b_names = set(b_parts.keys())

    added = sorted(b_names - a_names)
    removed = sorted(a_names - b_names)
    modified = sorted([n for n in (a_names & b_names) if a_parts[n] != b_parts[n]])

    return {
        "ignore": sorted(ignore),
        "added": added,
        "removed": removed,
        "modified": modified,
        "counts": {"added": len(added), "removed": len(removed), "modified": len(modified)},
        "equal": not added and not removed and not modified,
    }


def triage_workbook(workbook: WorkbookInput) -> dict[str, Any]:
    report: dict[str, Any] = {
        "display_name": workbook.display_name,
        "sha256": sha256_hex(workbook.data),
        "size_bytes": len(workbook.data),
        "timestamp": utc_now_iso(),
        "commit": github_commit_sha(),
        "run_url": github_run_url(),
    }

    # Step: load (zip + parse workbook.xml)
    start = _now_ms()
    try:
        with zipfile.ZipFile(io.BytesIO(workbook.data), "r") as z:
            zip_names = [info.filename for info in z.infolist() if not info.is_dir()]
            report["features"] = _scan_features(zip_names)
            report["functions"] = dict(_extract_function_counts(z))

            # Minimal "opens" check: workbook.xml must be present and parseable.
            wb_xml = z.read("xl/workbook.xml")
            from xml.etree import ElementTree as ET

            ET.fromstring(wb_xml)

        report["steps"] = {"load": asdict(_step_ok(start, details={"parts": len(zip_names)}))}
        report["result"] = {"open_ok": True}
    except Exception as e:  # noqa: BLE001 (reporting tool)
        report["steps"] = {"load": asdict(_step_failed(start, e))}
        report["result"] = {"open_ok": False}
        report["failure_category"] = "parse_error"
        return report

    # Step: recalc (placeholder until engine integration exists)
    report["steps"]["recalc"] = asdict(_step_skipped("no_calc_engine_configured"))
    report["result"]["calculate_ok"] = None

    # Step: render smoke (placeholder; opt-in to avoid heavy deps)
    report["steps"]["render"] = asdict(_step_skipped("no_headless_renderer_configured"))
    report["result"]["render_ok"] = None

    # Step: round-trip save (placeholder: byte-for-byte copy)
    start = _now_ms()
    round_tripped = workbook.data
    report["steps"]["round_trip"] = asdict(
        _step_ok(start, details={"engine": "copy"})
    )

    # Step: diff
    start = _now_ms()
    try:
        diff = _diff_workbooks(workbook.data, round_tripped)
        report["steps"]["diff"] = asdict(_step_ok(start, details=diff))
        report["result"]["round_trip_ok"] = diff["equal"]
        if not diff["equal"]:
            report["failure_category"] = "round_trip_diff"
    except Exception as e:  # noqa: BLE001
        report["steps"]["diff"] = asdict(_step_failed(start, e))
        report["result"]["round_trip_ok"] = False
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
            if exp_value is True and actual is False:
                regressions.append(f"{name}: expected {key}=true, got false")
            if exp_value is False and actual is True:
                improvements.append(f"{name}: expected {key}=false, got true")
    return regressions, improvements


def main() -> int:
    parser = argparse.ArgumentParser(description="Run corpus workbook triage.")
    parser.add_argument("--corpus-dir", type=Path, required=True)
    parser.add_argument("--out-dir", type=Path, required=True)
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
    args = parser.parse_args()

    ensure_dir(args.out_dir)
    reports_dir = args.out_dir / "reports"
    ensure_dir(reports_dir)

    fernet_key = os.environ.get(args.fernet_key_env)
    reports: list[dict[str, Any]] = []
    for path in iter_workbook_paths(args.corpus_dir):
        try:
            wb = read_workbook_input(path, fernet_key=fernet_key)
            report = triage_workbook(wb)
        except Exception as e:  # noqa: BLE001
            report = {
                "display_name": path.name,
                "timestamp": utc_now_iso(),
                "commit": github_commit_sha(),
                "run_url": github_run_url(),
                "steps": {"load": asdict(_step_failed(_now_ms(), e))},
                "result": {"open_ok": False},
                "failure_category": "parse_error",
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
