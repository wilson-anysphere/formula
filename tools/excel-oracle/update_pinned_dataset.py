#!/usr/bin/env python3
"""
Incrementally update the pinned Excel-oracle dataset (`excel-oracle.pinned.json`).

Why this exists
---------------

When adding new deterministic cases to `tests/compatibility/excel-oracle/cases.json`, the pinned
dataset used by CI must be updated to include results for the new case IDs.

Regenerating the *entire* pinned dataset via `tools/excel-oracle/regenerate_synthetic_baseline.py`
is correct but produces a very large diff (and often conflicts during rebases/parallel work).

This script keeps the existing pinned results and only fills in missing cases by:

1) Updating `caseSet.sha256`/`caseSet.count` to match the current cases.json
2) Removing results for case IDs that no longer exist in cases.json
3) Appending results for any missing case IDs, sourced from either:
   - one or more `--merge-results` JSON files (e.g. from a tag-filtered engine run), and/or
   - a targeted `formula-excel-oracle` run on a temporary corpus containing only the missing cases

Typical usage
-------------

Update pinned dataset after adding new cases:

  python tools/excel-oracle/update_pinned_dataset.py

If you already generated results for a subset (e.g. only the new Thai cases) and want to avoid a
full engine run:

  target/debug/formula-excel-oracle --cases tests/compatibility/excel-oracle/cases.json \\
    --out /tmp/new-results.json --include-tag thai
  python tools/excel-oracle/update_pinned_dataset.py --merge-results /tmp/new-results.json --no-engine

If you have **real Excel** results for a subset of cases and want to overwrite the synthetic
baseline values in the pinned dataset (keeping the rest of the corpus intact):

  powershell -ExecutionPolicy Bypass -File tools/excel-oracle/run-excel-oracle.ps1 `
    -CasesPath tools/excel-oracle/odd_coupon_long_stub_cases.json `
    -OutPath /tmp/excel-odd-coupon.json
  python tools/excel-oracle/update_pinned_dataset.py \
    --merge-results /tmp/excel-odd-coupon.json \
    --overwrite-existing \
    --no-engine
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Iterable


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, sort_keys=False)
        f.write("\n")


def _sanitize_fragment(text: str) -> str:
    # Keep filenames portable and reasonably readable.
    safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", text.strip())
    safe = re.sub(r"_+", "_", safe).strip("_")
    return safe or "unknown"


def _stable_path_string_from_payload(*, repo_root: Path, raw: object) -> str | None:
    """
    Normalize a case-set path string from a results payload so it is stable/portable when committed.

    `run-excel-oracle.ps1` is often invoked with an absolute `-CasesPath` (especially from wrapper
    scripts). If we record patch metadata into the pinned dataset, prefer a repo-relative path when
    possible so it is not machine-specific.
    """

    if not isinstance(raw, str):
        return None
    raw = raw.strip()
    if not raw:
        return None

    try:
        p = Path(raw)
        # Note: on non-Windows platforms, a Windows path like `C:\repo\cases.json` is not treated as
        # absolute by pathlib. In that situation we just keep the raw string.
        if p.is_absolute():
            try:
                rel = p.resolve().relative_to(repo_root.resolve())
                return rel.as_posix()
            except Exception:
                return p.as_posix()
        return p.as_posix()
    except Exception:
        return raw


def _is_real_excel_source(source: dict[str, Any]) -> bool:
    """
    Heuristic: real Excel results come from `run-excel-oracle.ps1` and have `source.kind == "excel"`
    with no `source.syntheticSource`.
    """

    return (
        source.get("kind") == "excel"
        and not isinstance(source.get("syntheticSource"), dict)
    )


def _record_excel_merge_patch(
    *,
    repo_root: Path,
    pinned_payload: dict[str, Any],
    merge_payload: dict[str, Any],
    added: int,
    overwritten: int,
) -> None:
    """
    Record (optional) metadata about a real-Excel patch into the pinned dataset.

    This is intentionally *best effort* and never blocks an update: it should not prevent CI/dev
    workflows if the merge payload is missing metadata.
    """

    if added <= 0 and overwritten <= 0:
        return

    pinned_source = pinned_payload.get("source")
    if not isinstance(pinned_source, dict):
        return

    merge_source = merge_payload.get("source")
    merge_case_set = merge_payload.get("caseSet")
    if not isinstance(merge_source, dict) or not isinstance(merge_case_set, dict):
        return

    if not _is_real_excel_source(merge_source):
        return

    raw_case_set_path = merge_case_set.get("path")
    stable_case_set_path = _stable_path_string_from_payload(repo_root=repo_root, raw=raw_case_set_path)

    patch_entry = {
        "kind": "excel",
        "version": str(merge_source.get("version", "unknown")),
        "build": str(merge_source.get("build", "unknown")),
        "operatingSystem": str(merge_source.get("operatingSystem", "unknown")),
        "caseSet": {
            "path": stable_case_set_path or str(raw_case_set_path or ""),
            "sha256": str(merge_case_set.get("sha256", "")),
            "count": int(merge_case_set.get("count", 0) or 0),
        },
        "applied": {"added": int(added), "overwritten": int(overwritten)},
    }

    patches = pinned_source.get("patches")
    if patches is None:
        pinned_source["patches"] = [patch_entry]
    elif isinstance(patches, list):
        # De-dupe: replace an existing entry for the same (Excel build, caseSet) instead of
        # appending repeatedly (merge-friendly).
        def key(entry: object) -> tuple[str, str, str, str, str] | None:
            if not isinstance(entry, dict):
                return None
            cs = entry.get("caseSet")
            if not isinstance(cs, dict):
                return None
            return (
                str(entry.get("version", "")),
                str(entry.get("build", "")),
                str(entry.get("operatingSystem", "")),
                str(cs.get("sha256", "")),
                str(cs.get("path", "")),
            )

        patch_key = key(patch_entry)
        if patch_key is None:
            patches.append(patch_entry)
        else:
            for i, existing in enumerate(list(patches)):
                if key(existing) == patch_key:
                    patches[i] = patch_entry
                    break
            else:
                patches.append(patch_entry)
    else:
        # Unexpected type; don't clobber it.
        return

    # If the pinned dataset is still tagged as a synthetic baseline, make the note accurate once
    # real Excel patches have been applied.
    note = pinned_source.get("note")
    if isinstance(note, str) and "Synthetic CI baseline" in note and "real Excel patches" not in note:
        pinned_source["note"] = note.rstrip() + " (Includes real Excel patches; see source.patches.)"


def write_versioned_copy(*, pinned_path: Path, versioned_dir: Path) -> Path:
    """
    Write a version-tagged copy of `pinned_path` into `versioned_dir`.

    This mirrors the naming scheme used by `tools/excel-oracle/pin_dataset.py` so other tooling
    (like `compat_gate.py`) can auto-select the expected dataset for the current `cases.json` hash.
    """

    payload = _load_json(pinned_path)
    source = payload.get("source", {})
    case_set = payload.get("caseSet", {})

    if not isinstance(source, dict):
        raise SystemExit(f"{pinned_path}: expected source object")
    if not isinstance(case_set, dict):
        raise SystemExit(f"{pinned_path}: expected caseSet object")

    excel_version = _sanitize_fragment(str(source.get("version", "unknown")))
    excel_build = _sanitize_fragment(str(source.get("build", "unknown")))
    cases_sha = _sanitize_fragment(str(case_set.get("sha256", "unknown")))

    versioned_name = f"excel-{excel_version}-build-{excel_build}-cases-{cases_sha[:8]}.json"
    versioned_dir.mkdir(parents=True, exist_ok=True)
    out_path = versioned_dir / versioned_name
    shutil.copyfile(pinned_path, out_path)
    return out_path


def _tool_env(repo_root: Path) -> dict[str, str]:
    """
    Build a conservative environment for running Cargo tools.

    In agent/CI environments we often want to avoid:
    - global Cargo home lock contention across concurrent processes
    - user/global Cargo config (which can set `build.rustc-wrapper = "sccache"` and be flaky)
    - extreme default parallelism on high-core-count hosts
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

    return env


def _stable_case_path_string(*, repo_root: Path, cases_path: Path) -> str:
    try:
        rel = cases_path.resolve().relative_to(repo_root.resolve())
        return rel.as_posix()
    except Exception:
        return cases_path.as_posix()


def _iter_result_entries(payload: Any) -> Iterable[dict[str, Any]]:
    if not isinstance(payload, dict):
        return []
    results = payload.get("results", [])
    if not isinstance(results, list):
        return []
    for r in results:
        if isinstance(r, dict):
            yield r


def _run_formula_excel_oracle(
    *,
    repo_root: Path,
    engine_bin: Path | None,
    cases_path: Path,
    out_path: Path,
    env: dict[str, str] | None,
) -> None:
    if engine_bin is not None:
        cmd = [str(engine_bin), "--cases", str(cases_path), "--out", str(out_path)]
        subprocess.run(cmd, cwd=str(repo_root), env=env, check=True)
        return

    use_cargo_agent = (
        os.name != "nt"
        and shutil.which("bash") is not None
        and (repo_root / "scripts" / "cargo_agent.sh").is_file()
    )

    if use_cargo_agent:
        cmd = [
            "bash",
            "scripts/cargo_agent.sh",
            "run",
            "-p",
            "formula-excel-oracle",
            "--quiet",
            "--locked",
            "--",
            "--cases",
            str(cases_path),
            "--out",
            str(out_path),
        ]
    else:
        cmd = [
            "cargo",
            "run",
            "-p",
            "formula-excel-oracle",
            "--quiet",
            "--locked",
            "--",
            "--cases",
            str(cases_path),
            "--out",
            str(out_path),
        ]
    subprocess.run(cmd, cwd=str(repo_root), env=env, check=True)


def update_pinned_dataset(
    *,
    cases_path: Path,
    pinned_path: Path,
    merge_results_paths: list[Path],
    engine_bin: Path | None,
    run_engine_for_missing: bool,
    env: dict[str, str] | None = None,
    force_engine: bool = False,
    overwrite_existing: bool = False,
) -> tuple[int, int]:
    """
    Update `pinned_path` in-place.

    Returns: (missing_before, missing_after)
    """

    repo_root = Path(__file__).resolve().parents[2]
    cases_payload = _load_json(cases_path)
    pinned_payload = _load_json(pinned_path)
    source = pinned_payload.get("source", {})
    # Pinned datasets are typically either:
    # - Real Excel results (source.kind == "excel", no syntheticSource)
    # - Synthetic baseline re-tagged as Excel (source.kind == "excel", syntheticSource present)
    #
    # Filling missing results by running the engine is only safe for the synthetic baseline.
    is_synthetic_baseline = isinstance(source, dict) and isinstance(source.get("syntheticSource"), dict)

    cases_sha = _sha256_file(cases_path)
    cases_list = cases_payload.get("cases", [])
    if not isinstance(cases_list, list):
        raise SystemExit(f"{cases_path}: expected top-level 'cases' array")

    case_ids: set[str] = set()
    for c in cases_list:
        if isinstance(c, dict) and isinstance(c.get("id"), str):
            case_ids.add(c["id"])

    if not case_ids:
        raise SystemExit(f"{cases_path}: no case IDs found")

    # Normalize pinned metadata.
    pinned_payload.setdefault("caseSet", {})
    if not isinstance(pinned_payload.get("caseSet"), dict):
        raise SystemExit(f"{pinned_path}: expected caseSet object")

    case_set = pinned_payload["caseSet"]
    assert isinstance(case_set, dict)
    case_set["path"] = _stable_case_path_string(repo_root=repo_root, cases_path=cases_path)
    case_set["sha256"] = cases_sha

    # Filter existing pinned results: drop duplicates + drop results for removed cases.
    existing_results = pinned_payload.get("results", [])
    if not isinstance(existing_results, list):
        raise SystemExit(f"{pinned_path}: expected top-level 'results' array")

    filtered_results: list[dict[str, Any]] = []
    seen: set[str] = set()
    index_by_case_id: dict[str, int] = {}
    for r in _iter_result_entries(pinned_payload):
        cid = r.get("caseId")
        if not isinstance(cid, str):
            continue
        if cid not in case_ids:
            continue
        if cid in seen:
            continue
        seen.add(cid)
        index_by_case_id[cid] = len(filtered_results)
        filtered_results.append(r)

    missing = set(case_ids.difference(seen))
    missing_before = len(missing)

    # Merge any provided results files before running the engine.
    for path in merge_results_paths:
        payload = _load_json(path)
        merged_added = 0
        merged_overwritten = 0
        for r in _iter_result_entries(payload):
            cid = r.get("caseId")
            if not isinstance(cid, str):
                continue
            if cid not in case_ids:
                continue
            if cid in missing:
                seen.add(cid)
                missing.remove(cid)
                index_by_case_id[cid] = len(filtered_results)
                filtered_results.append(r)
                merged_added += 1
                continue

            # Optional: allow merge-results to overwrite existing pinned results. This is useful
            # when gradually replacing a synthetic baseline with real Excel results for a subset
            # of cases (e.g. financial edge cases).
            if overwrite_existing and cid in seen:
                idx = index_by_case_id.get(cid)
                if idx is not None and 0 <= idx < len(filtered_results):
                    filtered_results[idx] = r
                    merged_overwritten += 1
                continue

        if isinstance(payload, dict):
            _record_excel_merge_patch(
                repo_root=repo_root,
                pinned_payload=pinned_payload,
                merge_payload=payload,
                added=merged_added,
                overwritten=merged_overwritten,
            )

    # If still missing, optionally run the engine on a temp corpus containing only missing cases.
    if missing and run_engine_for_missing:
        if not is_synthetic_baseline and not force_engine:
            raise SystemExit(
                "Refusing to fill missing oracle results by running formula-engine because the pinned dataset "
                "appears to be produced by real Excel (source.syntheticSource is missing). "
                "Generate additional Excel results and pass them via --merge-results, or re-run with "
                "--no-engine to require merge-only updates. If you intentionally want to produce a synthetic "
                "baseline dataset, pass --force-engine."
            )
        missing_cases = [c for c in cases_list if isinstance(c, dict) and c.get("id") in missing]
        temp_corpus = {
            "schemaVersion": cases_payload.get("schemaVersion"),
            "caseSet": cases_payload.get("caseSet"),
            "defaultSheet": cases_payload.get("defaultSheet"),
            "cases": missing_cases,
        }

        with tempfile.TemporaryDirectory(prefix="excel-oracle-missing-") as tmp:
            tmp_dir = Path(tmp)
            tmp_cases = tmp_dir / "missing-cases.json"
            tmp_results = tmp_dir / "missing-results.json"
            _write_json(tmp_cases, temp_corpus)
            _run_formula_excel_oracle(
                repo_root=repo_root,
                engine_bin=engine_bin,
                cases_path=tmp_cases,
                out_path=tmp_results,
                env=env,
            )
            payload = _load_json(tmp_results)
            for r in _iter_result_entries(payload):
                cid = r.get("caseId")
                if not isinstance(cid, str):
                    continue
                if cid not in missing:
                    continue
                if cid in seen:
                    continue
                seen.add(cid)
                missing.remove(cid)
                filtered_results.append(r)

    missing_after = len(missing)
    if missing_after:
        missing_preview = ", ".join(sorted(list(missing))[:25])
        suffix = "" if missing_after <= 25 else f" (+{missing_after - 25} more)"
        raise SystemExit(
            "Pinned dataset is still missing results for some case IDs. "
            "Provide additional --merge-results or re-run without --no-engine. "
            f"Missing ({missing_after}): {missing_preview}{suffix}"
        )

    pinned_payload["results"] = filtered_results
    case_set["count"] = len(filtered_results)

    _write_json(pinned_path, pinned_payload)
    return (missing_before, missing_after)


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument(
        "--cases",
        default="tests/compatibility/excel-oracle/cases.json",
        help="Path to cases.json (default: %(default)s)",
    )
    p.add_argument(
        "--pinned",
        default="tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json",
        help="Path to pinned dataset to update (default: %(default)s)",
    )
    p.add_argument(
        "--versioned-dir",
        default="tests/compatibility/excel-oracle/datasets/versioned",
        help=(
            "If set, also write/update a version-tagged copy of the pinned dataset in this directory "
            "(default: %(default)s). Use --no-versioned to disable."
        ),
    )
    p.add_argument(
        "--no-versioned",
        action="store_true",
        help="Do not write/update the versioned dataset copy (only update the pinned dataset).",
    )
    p.add_argument(
        "--merge-results",
        action="append",
        default=[],
        help=(
            "Path to a results JSON file (engine output schema) to merge into the pinned dataset "
            "before running the engine (can be repeated)."
        ),
    )
    p.add_argument(
        "--overwrite-existing",
        action="store_true",
        help=(
            "When merging --merge-results, overwrite existing case results in the pinned dataset "
            "(default: only fill missing cases). Useful for patching a synthetic baseline with "
            "real Excel results for specific case IDs."
        ),
    )
    p.add_argument(
        "--no-engine",
        action="store_true",
        help="Do not run formula-excel-oracle. Require --merge-results to cover all missing cases.",
    )
    p.add_argument(
        "--force-engine",
        action="store_true",
        help=(
            "Allow filling missing case results by running formula-engine even if the pinned dataset appears "
            "to be produced by real Excel. This will produce a mixed dataset and is not recommended."
        ),
    )
    p.add_argument(
        "--engine-bin",
        default="",
        help=(
            "Optional path to a prebuilt formula-excel-oracle binary. If omitted, the script will "
            "use target/debug/formula-excel-oracle when present, else fall back to `cargo run`."
        ),
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Print what would change (missing cases, whether the engine would run, etc) without writing files or invoking cargo.",
    )
    args = p.parse_args()

    repo_root = Path(__file__).resolve().parents[2]
    env = _tool_env(repo_root)
    cases_path = (repo_root / args.cases).resolve() if not os.path.isabs(args.cases) else Path(args.cases)
    pinned_path = (repo_root / args.pinned).resolve() if not os.path.isabs(args.pinned) else Path(args.pinned)

    merge_results_paths = [Path(p).resolve() for p in args.merge_results]

    engine_bin: Path | None = None
    if args.engine_bin:
        engine_bin = Path(args.engine_bin).resolve()
    else:
        candidate = repo_root / "target" / "debug" / "formula-excel-oracle"
        if candidate.is_file() and os.access(candidate, os.X_OK):
            engine_bin = candidate

    if args.dry_run:
        cases_payload = _load_json(cases_path)
        pinned_payload = _load_json(pinned_path)

        cases_list = cases_payload.get("cases", [])
        if not isinstance(cases_list, list):
            raise SystemExit(f"{cases_path}: expected top-level 'cases' array")

        case_ids: set[str] = set()
        for c in cases_list:
            if isinstance(c, dict) and isinstance(c.get("id"), str):
                case_ids.add(c["id"])
        if not case_ids:
            raise SystemExit(f"{cases_path}: no case IDs found")

        pinned_source = pinned_payload.get("source", {})
        is_synthetic_baseline = isinstance(pinned_source, dict) and isinstance(
            pinned_source.get("syntheticSource"), dict
        )

        # Mirror the updater's duplicate/removal filtering when counting coverage.
        seen: set[str] = set()
        for r in _iter_result_entries(pinned_payload):
            cid = r.get("caseId")
            if not isinstance(cid, str):
                continue
            if cid not in case_ids:
                continue
            if cid in seen:
                continue
            seen.add(cid)

        missing = set(case_ids.difference(seen))
        missing_before = len(missing)

        merged = 0
        overwritten = 0
        for path in merge_results_paths:
            payload = _load_json(path)
            for r in _iter_result_entries(payload):
                cid = r.get("caseId")
                if not isinstance(cid, str):
                    continue
                if cid not in case_ids:
                    continue
                if cid in missing:
                    missing.remove(cid)
                    merged += 1
                    continue
                if args.overwrite_existing and cid in seen:
                    overwritten += 1

        missing_after_merges = len(missing)
        would_run_engine = missing_after_merges > 0 and not args.no_engine

        cases_sha8 = _sha256_file(cases_path)[:8]
        print("Dry run: update_pinned_dataset")
        print(f"cases:  {cases_path} ({len(case_ids)} cases, sha256={cases_sha8}...)")
        print(f"pinned: {pinned_path}")
        print(f"missing before merges: {missing_before}")
        if merge_results_paths:
            print(f"merge-results: {len(merge_results_paths)} file(s) (filled {merged} missing)")
            if args.overwrite_existing:
                print(f"overwrite-existing: true (would overwrite {overwritten} existing case(s))")
        else:
            print("merge-results: <none>")
        print(f"missing after merges: {missing_after_merges}")
        if would_run_engine:
            if not is_synthetic_baseline and not args.force_engine:
                print(
                    "engine: would refuse to fill missing cases because the pinned dataset appears to be real Excel "
                    "(source.syntheticSource missing). Pass --force-engine or provide --merge-results."
                )
                return 1
            engine_desc = str(engine_bin) if engine_bin is not None else "<cargo run -p formula-excel-oracle>"
            print(f"engine: would run for {missing_after_merges} missing case(s) via {engine_desc}")
        else:
            print("engine: <skipped>")
            if missing_after_merges:
                print("note: update would fail because cases are still missing and --no-engine was set")
                return 1

        if not args.no_versioned:
            raw = str(args.versioned_dir or "").strip()
            if raw:
                versioned_dir = (repo_root / raw).resolve() if not os.path.isabs(raw) else Path(raw).resolve()
                print(f"versioned copy: would write/update under {versioned_dir.as_posix()}")
        else:
            print("versioned copy: <skipped>")

        print("pinned dataset: would be updated in-place")
        return 0

    missing_before, missing_after = update_pinned_dataset(
        cases_path=cases_path,
        pinned_path=pinned_path,
        merge_results_paths=merge_results_paths,
        engine_bin=engine_bin,
        run_engine_for_missing=not args.no_engine,
        env=env,
        force_engine=args.force_engine,
        overwrite_existing=args.overwrite_existing,
    )

    if missing_before == 0:
        print("Pinned dataset already covered all cases (updated metadata only).")
    else:
        print(f"Filled {missing_before - missing_after}/{missing_before} missing case results.")

    if not args.no_versioned:
        raw = str(args.versioned_dir or "").strip()
        if raw:
            versioned_dir = (repo_root / raw).resolve() if not os.path.isabs(raw) else Path(raw).resolve()
            out = write_versioned_copy(pinned_path=pinned_path, versioned_dir=versioned_dir)
            print(f"Versioned copy -> {out.as_posix()}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
