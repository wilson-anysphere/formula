#!/usr/bin/env python3
"""
Regenerate the Excel-oracle artifacts from the in-repo Rust formula engine.

This script is intended as an "integration safety net" for changes that add new
deterministic built-in functions (e.g. STAT distributions / moments / frequency
functions) and therefore need coordinated updates across:

- shared/functionCatalog.json (+ .mjs + .d.mts)
- tools/excel-oracle/generate_cases.py (must include coverage cases)
- tests/compatibility/excel-oracle/cases.json (generated)
- tools/excel-oracle/odd_coupon_*_cases.json (derived subset corpora for quick Excel runs)
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
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Sequence


def _run(*, cmd: Sequence[str], cwd: Path, env: dict[str, str], dry_run: bool) -> None:
    rendered = " ".join(cmd)
    print(f"+ {rendered}")
    if dry_run:
        return
    subprocess.run(list(cmd), cwd=str(cwd), env=env, check=True)


def _have_command(name: str) -> bool:
    return shutil.which(name) is not None


def _tool_env(repo_root: Path) -> dict[str, str]:
    """
    Build a conservative environment for running Cargo/Node/Python tools.

    In agent/CI environments we often want to avoid:
    - global Cargo home lock contention across concurrent processes
    - user/global Cargo config (which can set `build.rustc-wrapper = "sccache"` and be flaky)
    """

    env = dict(os.environ)
    # `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml`. Some environments set it
    # globally (often to `stable`), which would bypass the pinned toolchain and reintroduce drift
    # when running `cargo` directly.
    if env.get("RUSTUP_TOOLCHAIN") and (repo_root / "rust-toolchain.toml").is_file():
        env.pop("RUSTUP_TOOLCHAIN", None)

    default_global_cargo_home = Path.home() / ".cargo"
    cargo_home = env.get("CARGO_HOME")
    cargo_home_path = Path(cargo_home).expanduser() if cargo_home else None

    if not cargo_home or (
        not env.get("CI")
        and not env.get("FORMULA_ALLOW_GLOBAL_CARGO_HOME")
        and cargo_home_path == default_global_cargo_home
    ):
        env["CARGO_HOME"] = str(repo_root / "target" / "cargo-home")

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

    return env


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
        help="Run validation tests after regeneration (formula-engine + node function-catalog + tools/excel-oracle python tests).",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Print the commands that would run without executing them or writing any files.",
    )
    args = p.parse_args()

    repo_root = Path(__file__).resolve().parents[2]
    # Use a *relative* cases path when invoking tools so the pinned dataset metadata stays stable
    # (and doesn't capture developer/CI absolute paths).
    cases_relpath = Path("tests/compatibility/excel-oracle/cases.json")
    pinned_relpath = Path(
        "tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json"
    )
    versioned_relpath = Path("tests/compatibility/excel-oracle/datasets/versioned")
    cases_path = repo_root / cases_relpath
    pinned_path = repo_root / pinned_relpath

    env = _tool_env(repo_root)

    if not args.skip_function_catalog:
        if not _have_command("node"):
            raise SystemExit(
                "node was not found on PATH. Install node, or re-run with --skip-function-catalog."
            )
        _run(
            cmd=("node", "scripts/generate-function-catalog.js"),
            cwd=repo_root,
            env=env,
            dry_run=args.dry_run,
        )

    if not args.skip_cases:
        _run(
            cmd=(
                sys.executable,
                "tools/excel-oracle/generate_cases.py",
                "--out",
                str(cases_relpath),
            ),
            cwd=repo_root,
            env=env,
            dry_run=args.dry_run,
        )
        # Keep convenience subset corpora (under tools/excel-oracle/) aligned with the canonical
        # cases.json corpus so Windows+Excel runs can target a small case set and still merge
        # results back into the pinned dataset by canonical caseId.
        _run(
            cmd=(sys.executable, "tools/excel-oracle/regenerate_subset_corpora.py"),
            cwd=repo_root,
            env=env,
            dry_run=args.dry_run,
        )

    if not args.skip_pinned:
        with tempfile.TemporaryDirectory(prefix="excel-oracle-") as tmp:
            engine_results_path = Path(tmp) / "engine-results.json"
            if _have_command("bash") and (repo_root / "scripts/cargo_agent.sh").is_file():
                _run(
                    cmd=(
                        "bash",
                        "scripts/cargo_agent.sh",
                        "run",
                        "--quiet",
                        "-p",
                        "formula-excel-oracle",
                        "--locked",
                        "--",
                        "--cases",
                        str(cases_relpath),
                        "--out",
                        str(engine_results_path),
                    ),
                    cwd=repo_root,
                    env=env,
                    dry_run=args.dry_run,
                )
            else:
                _run(
                    cmd=(
                        "cargo",
                        "run",
                        "--quiet",
                        "-p",
                        "formula-excel-oracle",
                        "--locked",
                        "--",
                        "--cases",
                        str(cases_relpath),
                        "--out",
                        str(engine_results_path),
                    ),
                    cwd=repo_root,
                    env=env,
                    dry_run=args.dry_run,
                )
            _run(
                cmd=(
                    sys.executable,
                    "tools/excel-oracle/pin_dataset.py",
                    "--dataset",
                    str(engine_results_path),
                    "--pinned",
                    str(pinned_relpath),
                    "--versioned-dir",
                    str(versioned_relpath),
                ),
                cwd=repo_root,
                env=env,
                dry_run=args.dry_run,
            )

    if args.run_tests:
        # Prefer `scripts/cargo_agent.sh` when bash is available (it sets conservative defaults for
        # high-core-count environments), but fall back to plain `cargo test` on platforms without bash.
        if _have_command("bash") and (repo_root / "scripts/cargo_agent.sh").is_file():
            _run(
                cmd=("bash", "scripts/cargo_agent.sh", "test", "-p", "formula-engine"),
                cwd=repo_root,
                env=env,
                dry_run=args.dry_run,
            )
        else:
            _run(
                cmd=("cargo", "test", "-p", "formula-engine"),
                cwd=repo_root,
                env=env,
                dry_run=args.dry_run,
            )
        # Validate that the committed JS/TS function catalog artifacts are in sync.
        if _have_command("node"):
            _run(
                cmd=(
                    "node",
                    "--test",
                    "packages/ai-completion/test/functionCatalogArtifact.test.js",
                ),
                cwd=repo_root,
                env=env,
                dry_run=args.dry_run,
            )
        else:
            print("Skipping function catalog node:test suite (node not found on PATH).")
        _run(
            cmd=(sys.executable, "-m", "unittest", "discover", "-s", "tools/excel-oracle/tests"),
            cwd=repo_root,
            env=env,
            dry_run=args.dry_run,
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
