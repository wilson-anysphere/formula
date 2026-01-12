#!/usr/bin/env python3
"""
Regenerate the Excel-oracle artifacts from the in-repo Rust formula engine.

This script is intended as an "integration safety net" for changes that add new
deterministic built-in functions (e.g. STAT distributions / moments / frequency
functions) and therefore need coordinated updates across:

- shared/functionCatalog.json (+ .mjs)
- tools/excel-oracle/generate_cases.py (must include coverage cases)
- tests/compatibility/excel-oracle/cases.json (generated)
- tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json (synthetic baseline)

Unlike tools/excel-oracle/run-excel-oracle.ps1, this does NOT require Microsoft
Excel. It uses `crates/formula-excel-oracle` to evaluate the corpus with the
formula-engine, then pins the results as a synthetic CI baseline via
tools/excel-oracle/pin_dataset.py.

Typical usage (from repo root):

  python tools/excel-oracle/regenerate_synthetic_baseline.py

Then commit the resulting diffs.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Sequence


def _run(*, cmd: Sequence[str], cwd: Path) -> None:
    rendered = " ".join(cmd)
    print(f"+ {rendered}")
    subprocess.run(list(cmd), cwd=str(cwd), check=True)


def _have_command(name: str) -> bool:
    return shutil.which(name) is not None


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument(
        "--skip-function-catalog",
        action="store_true",
        help="Skip regenerating shared/functionCatalog.json via scripts/generate-function-catalog.js",
    )
    p.add_argument(
        "--skip-cases",
        action="store_true",
        help="Skip regenerating tests/compatibility/excel-oracle/cases.json via tools/excel-oracle/generate_cases.py",
    )
    p.add_argument(
        "--skip-pinned",
        action="store_true",
        help="Skip regenerating tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json via crates/formula-excel-oracle + pin_dataset.py",
    )
    p.add_argument(
        "--run-tests",
        action="store_true",
        help="Run validation tests after regeneration (formula-engine + tools/excel-oracle python tests).",
    )
    args = p.parse_args()

    repo_root = Path(__file__).resolve().parents[2]
    cases_path = repo_root / "tests/compatibility/excel-oracle/cases.json"
    pinned_path = repo_root / "tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json"

    if not args.skip_function_catalog:
        if not _have_command("node"):
            raise SystemExit(
                "node was not found on PATH. Install node, or re-run with --skip-function-catalog."
            )
        _run(cmd=("node", "scripts/generate-function-catalog.js"), cwd=repo_root)

    if not args.skip_cases:
        _run(
            cmd=(
                sys.executable,
                "tools/excel-oracle/generate_cases.py",
                "--out",
                str(cases_path),
            ),
            cwd=repo_root,
        )

    if not args.skip_pinned:
        with tempfile.TemporaryDirectory(prefix="excel-oracle-") as tmp:
            engine_results_path = Path(tmp) / "engine-results.json"
            _run(
                cmd=(
                    "cargo",
                    "run",
                    "--quiet",
                    "-p",
                    "formula-excel-oracle",
                    "--",
                    "--cases",
                    str(cases_path),
                    "--out",
                    str(engine_results_path),
                ),
                cwd=repo_root,
            )
            _run(
                cmd=(
                    sys.executable,
                    "tools/excel-oracle/pin_dataset.py",
                    "--dataset",
                    str(engine_results_path),
                    "--pinned",
                    str(pinned_path),
                ),
                cwd=repo_root,
            )

    if args.run_tests:
        # Prefer `scripts/cargo_agent.sh` when bash is available (it sets conservative defaults for
        # high-core-count environments), but fall back to plain `cargo test` on platforms without bash.
        if _have_command("bash") and (repo_root / "scripts/cargo_agent.sh").is_file():
            _run(cmd=("bash", "scripts/cargo_agent.sh", "test", "-p", "formula-engine"), cwd=repo_root)
        else:
            _run(cmd=("cargo", "test", "-p", "formula-engine"), cwd=repo_root)
        _run(cmd=(sys.executable, "-m", "unittest", "discover", "-s", "tools/excel-oracle/tests"), cwd=repo_root)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

