#!/usr/bin/env python3
"""
End-to-end Excel-oracle compatibility gate for the Rust formula engine.

This is intentionally lightweight and CI-friendly:

  1) Run `crates/formula-excel-oracle` to produce engine-results.json
  2) Compare against a pinned Excel dataset via `tools/excel-oracle/compare.py`

The compare step emits `tests/compatibility/excel-oracle/reports/mismatch-report.json`
and exits non-zero if the mismatch rate exceeds the configured threshold.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shlex
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any, NamedTuple


SMOKE_INCLUDE_TAGS = [
    # Keep CI bounded to a small, high-signal slice of the corpus.
    "add",
    "sub",
    "mul",
    "div",
    "cmp",
    "SUM",
    "IF",
    "IFERROR",
    "error",
    # Minimal spill coverage (range reference + a couple of array functions).
    "range",
    "TRANSPOSE",
    "SEQUENCE",
    "FREQUENCY",
    # Representative deterministic functions beyond the arithmetic baseline.
    "COUNT",
    "COUNTIF",
    "TEXT",
    "TEXTSPLIT",
    "VALUE",
    "DATEVALUE",
    "WORKDAY",
    "NETWORKDAYS",
    "FVSCHEDULE",
    "XLOOKUP",
    "XMATCH",
    "FILTER",
    "SORT",
    "UNIQUE",
    "ISERROR",
    # Small, high-signal slice of statistical regression functions.
    "LINEST",
    "LOGEST",
    "TREND",
    "GROWTH",
    # Thai deterministic localization functions (BAHTTEXT/THAI*/ROUNDBAHT*).
    "thai",
    # Exercise small-but-important date validation boundary conditions (e.g. odd-coupon bonds where
    # settlement == coupon date) without pulling in the full financial corpus.
    "boundary",
    # Keep a small slice of the odd-coupon bond corpus in CI so ODDF*/ODDL* regressions are caught
    # without enabling the full `financial` tag set.
    "odd_coupon",
    # Keep COUP* schedule/date-count behavior in the smoke gate (basis=4 February/EOM quirks are
    # easy to regress, and this tag set is small).
    "coupon_schedule",
    # Explicitly include value coercion cases so CI exercises the conversion rules
    # (text -> number/date/time) we implement and diff against Excel later.
    "coercion",
]

P0_INCLUDE_TAGS = [
    # A broader but still bounded set intended to cover common "P0" Excel behavior.
    #
    # Note: include-tag filtering uses OR semantics (a case is included if it contains
    # ANY include tag). These tags map to the curated corpus tags produced by
    # tools/excel-oracle/generate_cases.py.
    "arith",
    "cmp",
    "math",
    "agg",
    "logical",
    "text",
    "date",
    "lookup",
    # Dynamic arrays / spilling.
    "spill",
    "dynarr",
    # Explicit error cases (and any cases tagged as error).
    "error",
    # Common info/conversion semantics (ensure p0 is a strict superset of smoke).
    "info",
    "coercion",
    # Thai deterministic localization functions (BAHTTEXT/THAI*/ROUNDBAHT*).
    "thai",
]

_TIER_TO_INCLUDE_TAGS: dict[str, list[str]] = {
    "smoke": list(SMOKE_INCLUDE_TAGS),
    "p0": list(P0_INCLUDE_TAGS),
    # Full corpus run: no include-tag filtering unless the user passes --include-tag.
    "full": [],
}


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


class _ExpectedDatasetInfo(NamedTuple):
    is_synthetic: bool
    synthetic_source: dict[str, Any] | None
    patches: list[Any]


def _inspect_expected_dataset(expected_path: Path) -> _ExpectedDatasetInfo | None:
    """
    Best-effort parse of the pinned expected dataset.

    Returns None if the dataset cannot be parsed. This is a CI/developer ergonomics helper; the
    correctness gate is enforced by the downstream compare step.
    """

    try:
        payload = json.loads(expected_path.read_text(encoding="utf-8"))
    except Exception:
        return None
    if not isinstance(payload, dict):
        return None

    source = payload.get("source")
    if not isinstance(source, dict):
        return None

    synthetic_source = source.get("syntheticSource")
    is_synthetic = isinstance(synthetic_source, dict)

    patches = source.get("patches")
    if not isinstance(patches, list):
        patches = []

    return _ExpectedDatasetInfo(
        is_synthetic=is_synthetic,
        synthetic_source=synthetic_source if isinstance(synthetic_source, dict) else None,
        patches=patches,
    )


def _print_expected_dataset_synthetic_warning(*, expected_path: Path, info: _ExpectedDatasetInfo) -> None:
    """
    Print a loud, explicit warning when the expected dataset is synthetic.

    The pinned dataset in this repo is often a "synthetic CI baseline" (engine-generated), tagged
    as source.kind="excel" but annotated with source.syntheticSource. It is critical that CI logs
    make this distinction unambiguous so mismatches aren't mistaken for true Excel drift.
    """

    if not info.is_synthetic or not info.synthetic_source:
        return

    synth = info.synthetic_source
    kind = str(synth.get("kind", "unknown"))
    version = str(synth.get("version", "unknown"))
    os_name = str(synth.get("os", "unknown"))
    arch = str(synth.get("arch", "unknown"))
    case_set = str(synth.get("caseSet", "unknown"))

    msg_lines = [
        "",
        "=" * 79,
        "WARNING: Expected Excel oracle dataset is SYNTHETIC (not generated by real Excel).",
        f"Path: {expected_path}",
        "",
        "syntheticSource:",
        f"  kind:    {kind}",
        f"  version: {version}",
        f"  os:      {os_name}",
        f"  arch:    {arch}",
        f"  caseSet: {case_set}",
        "=" * 79,
        "",
    ]
    sys.stderr.write("\n".join(msg_lines))
    sys.stderr.flush()


def _print_expected_dataset_patch_summary(*, expected_path: Path) -> None:
    """
    Best-effort reporting for mixed synthetic+Excel datasets.

    The repo typically uses a synthetic baseline pinned dataset (source.kind="excel" with
    source.syntheticSource metadata). Over time we may incrementally overwrite subsets of case IDs
    with real Excel results via `update_pinned_dataset.py --merge-results --overwrite-existing`.

    When those real-Excel merges occur, `update_pinned_dataset.py` records compact provenance entries
    under `source.patches` in the pinned dataset. Surfacing a summary here makes it obvious in CI
    logs whether the expected dataset includes any real Excel patches.
    """

    try:
        payload = json.loads(expected_path.read_text(encoding="utf-8"))
    except Exception:
        return
    if not isinstance(payload, dict):
        return
    source = payload.get("source")
    if not isinstance(source, dict):
        return
    patches = source.get("patches")
    if not isinstance(patches, list) or not patches:
        return

    # Keep logs compact; print a short summary and cap the number of entries we render.
    print(f"Expected dataset includes {len(patches)} real Excel patch entries (source.patches):")
    max_entries = 10
    shown = 0
    for entry in patches:
        if shown >= max_entries:
            break
        if not isinstance(entry, dict):
            continue

        version = str(entry.get("version", "unknown"))
        build = str(entry.get("build", "unknown"))
        os_name = str(entry.get("operatingSystem", "unknown"))
        case_set = entry.get("caseSet")
        if not isinstance(case_set, dict):
            case_set = {}
        case_path = str(case_set.get("path", "")).strip()
        case_sha = str(case_set.get("sha256", "")).strip()
        applied = entry.get("applied")
        if not isinstance(applied, dict):
            applied = {}
        added = int(applied.get("added", 0) or 0)
        overwritten = int(applied.get("overwritten", 0) or 0)

        suffix = []
        if case_sha:
            suffix.append(f"cases_sha={case_sha[:8]}")
        suffix.append(f"added={added}")
        suffix.append(f"overwritten={overwritten}")
        suffix_text = ", ".join(suffix)

        if case_path:
            print(f"  - Excel {version} build {build} ({os_name}): {case_path} ({suffix_text})")
        else:
            print(f"  - Excel {version} build {build} ({os_name}) ({suffix_text})")
        shown += 1

    if len(patches) > shown:
        print(f"  (+{len(patches) - shown} more)")


def _default_expected_dataset(*, cases_path: Path) -> Path:
    versioned_dir = Path("tests/compatibility/excel-oracle/datasets/versioned")
    if versioned_dir.is_dir():
        cases_sha8 = _sha256_file(cases_path)[:8]
        candidates = sorted(p for p in versioned_dir.glob(f"*-cases-{cases_sha8}.json") if p.is_file())
        if candidates:
            # Multiple Excel versions/builds can share the same corpus hash.
            # Keep this deterministic by selecting the lexicographically last filename
            # (version/build are embedded in the name).
            non_unknown = [p for p in candidates if "-unknown-build-unknown-" not in p.name]
            if non_unknown:
                return non_unknown[-1]
            return candidates[-1]

    pinned = Path("tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json")
    if pinned.is_file():
        return pinned

    raise SystemExit(
        "No pinned Excel oracle dataset found. Expected either:\n"
        "  - tests/compatibility/excel-oracle/datasets/versioned/*.json\n"
        "  - tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json\n"
        "\n"
        "See tests/compatibility/excel-oracle/README.md for how to generate/pin datasets."
    )


def _normalize_tags(tags: list[str]) -> list[str]:
    return [t for t in (s.strip() for s in tags) if t]


def _count_selected_cases(
    *, cases_path: Path, include_tags: list[str], exclude_tags: list[str], max_cases: int
) -> tuple[int, int]:
    """
    Return `(matched_after_tag_filters, selected_after_max_cases)`.

    This mirrors the include/exclude tag semantics used by `compare.py` and the engine runner:
    - include tags use OR semantics (a case is included if it has any included tag)
    - exclude tags filter out any case that contains one of the excluded tags
    """

    payload = json.loads(cases_path.read_text(encoding="utf-8"))
    cases = payload.get("cases", [])
    if not isinstance(cases, list):
        return (0, 0)

    include = set(include_tags)
    exclude = set(exclude_tags)

    matched = 0
    for case in cases:
        if not isinstance(case, dict):
            continue
        tags = case.get("tags", [])
        if not isinstance(tags, list):
            tags = []
        tag_set = {t for t in tags if isinstance(t, str)}

        if include and not (include & tag_set):
            continue
        if exclude and (exclude & tag_set):
            continue
        matched += 1

    selected = matched
    if max_cases and max_cases > 0:
        selected = min(selected, max_cases)
    return (matched, selected)


def _effective_include_tags(*, tier: str, user_include_tags: list[str]) -> list[str]:
    normalized = _normalize_tags(user_include_tags)
    if normalized:
        return normalized
    return list(_TIER_TO_INCLUDE_TAGS[tier])


def _build_engine_cmd(
    *,
    cases_path: Path,
    actual_path: Path,
    max_cases: int,
    include_tags: list[str],
    exclude_tags: list[str],
    use_cargo_agent: bool = False,
) -> list[str]:
    if use_cargo_agent:
        cmd = ["bash", "scripts/cargo_agent.sh", "run"]
    else:
        cmd = ["cargo", "run"]

    cmd += [
        "-p",
        "formula-excel-oracle",
        "--quiet",
        "--locked",
        "--",
        "--cases",
        str(cases_path),
        "--out",
        str(actual_path),
    ]
    if max_cases and max_cases > 0:
        cmd += ["--max-cases", str(max_cases)]
    for t in include_tags:
        cmd += ["--include-tag", t]
    for t in exclude_tags:
        cmd += ["--exclude-tag", t]
    return cmd


def _build_compare_cmd(
    *,
    cases_path: Path,
    expected_path: Path,
    actual_path: Path,
    report_path: Path,
    max_cases: int,
    include_tags: list[str],
    exclude_tags: list[str],
    max_mismatch_rate: float,
    abs_tol: float,
    rel_tol: float,
    tag_abs_tol: list[str],
    tag_rel_tol: list[str],
) -> list[str]:
    cmd = [
        sys.executable,
        "tools/excel-oracle/compare.py",
        "--cases",
        str(cases_path),
        "--expected",
        str(expected_path),
        "--actual",
        str(actual_path),
        "--report",
        str(report_path),
        "--max-mismatch-rate",
        str(max_mismatch_rate),
        "--abs-tol",
        str(abs_tol),
        "--rel-tol",
        str(rel_tol),
    ]
    for raw in tag_abs_tol:
        tol = raw.strip()
        if tol:
            cmd += ["--tag-abs-tol", tol]
    for raw in tag_rel_tol:
        tol = raw.strip()
        if tol:
            cmd += ["--tag-rel-tol", tol]
    if max_cases and max_cases > 0:
        cmd += ["--max-cases", str(max_cases)]
    for t in include_tags:
        cmd += ["--include-tag", t]
    for t in exclude_tags:
        cmd += ["--exclude-tag", t]
    return cmd


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument(
        "--cases",
        default="tests/compatibility/excel-oracle/cases.json",
        help="Path to cases.json (default: %(default)s)",
    )
    p.add_argument(
        "--expected",
        default="",
        help=(
            "Path to pinned Excel results JSON. Defaults to the newest matching file in "
            "tests/compatibility/excel-oracle/datasets/versioned/ (by cases.json SHA-256 suffix "
            "'*-cases-<sha8>.json') if present, else "
            "tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json."
        ),
    )
    p.add_argument(
        "--actual",
        default="tests/compatibility/excel-oracle/datasets/engine-results.json",
        help="Where to write engine results JSON (default: %(default)s)",
    )
    p.add_argument(
        "--report",
        default="tests/compatibility/excel-oracle/reports/mismatch-report.json",
        help="Where to write mismatch report JSON (default: %(default)s)",
    )
    p.add_argument(
        "--report-md",
        default="tests/compatibility/excel-oracle/reports/summary.md",
        help="Where to write a human-readable markdown summary (default: %(default)s)",
    )
    p.add_argument(
        "--max-cases",
        type=int,
        default=0,
        help="Optional cap (after tag filtering): evaluate/compare only first N cases (0 = all).",
    )
    p.add_argument(
        "--tier",
        choices=["smoke", "p0", "full"],
        default="smoke",
        help=(
            "Which Excel-compatibility gate to run: "
            "'smoke' (default, fast CI slice), "
            "'p0' (broader common-function slice), "
            "'full' (no include-tag filtering; runs all cases). "
            "If --include-tag is provided, it overrides tier presets."
        ),
    )
    p.add_argument(
        "--include-tag",
        action="append",
        default=[],
        help="Only include cases containing this tag (can be repeated). Overrides --tier presets.",
    )
    p.add_argument(
        "--exclude-tag",
        action="append",
        default=[],
        help="Exclude cases containing this tag (can be repeated).",
    )
    p.add_argument("--abs-tol", type=float, default=1e-9)
    p.add_argument("--rel-tol", type=float, default=1e-9)
    p.add_argument(
        "--tag-abs-tol",
        action="append",
        default=[],
        help=(
            "Override numeric abs tolerance for cases that contain a tag. Format TAG=FLOAT "
            "(example: odd_coupon=1e-6). Can be repeated; the maximum across matching tags wins."
        ),
    )
    p.add_argument(
        "--tag-rel-tol",
        action="append",
        default=[],
        help=(
            "Override numeric rel tolerance for cases that contain a tag. Format TAG=FLOAT "
            "(example: odd_coupon=1e-6). Can be repeated; the maximum across matching tags wins."
        ),
    )
    p.add_argument(
        "--max-mismatch-rate",
        type=float,
        default=0.0,
        help="Fail if mismatches / total exceeds this threshold (default 0).",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Print the engine/compare commands and exit without running them.",
    )
    p.add_argument(
        "--require-real-excel",
        action="store_true",
        help=(
            "Fail if the expected dataset is a synthetic baseline (contains source.syntheticSource). "
            "Useful in CI when you expect a real Excel-generated dataset to be present."
        ),
    )
    args = p.parse_args()

    cases_path = Path(args.cases)
    expected_path = Path(args.expected) if args.expected else _default_expected_dataset(cases_path=cases_path)
    actual_path = Path(args.actual)
    report_path = Path(args.report)

    # Keep the gate debuggable in CI logs: the "expected dataset" selection has precedence rules
    # (versioned corpus-hash match vs pinned fallback) that can otherwise be non-obvious when
    # investigating mismatches.
    print(
        f"Excel-oracle datasets: expected={expected_path} actual={actual_path} report={report_path}",
        flush=True,
    )
    expected_info = _inspect_expected_dataset(expected_path)
    if expected_info is not None:
        if expected_info.is_synthetic:
            _print_expected_dataset_synthetic_warning(expected_path=expected_path, info=expected_info)
            if args.require_real_excel:
                sys.stderr.write(
                    "ERROR: --require-real-excel was set, but the expected dataset contains "
                    "source.syntheticSource (synthetic baseline).\n"
                )
                return 1
        else:
            if expected_info.patches:
                print(
                    "Expected dataset appears to be real Excel with patches applied "
                    f"(source.patches entries: {len(expected_info.patches)}).",
                    flush=True,
                )
            else:
                print(
                    "Expected dataset appears to be real Excel (no source.syntheticSource).",
                    flush=True,
                )
    _print_expected_dataset_patch_summary(expected_path=expected_path)

    include_tags = _effective_include_tags(tier=args.tier, user_include_tags=args.include_tag)
    exclude_tags = _normalize_tags(args.exclude_tag)

    # Some Excel functions are inherently iterative (e.g. yield solvers), and even when the math is
    # correct we can see small (~1e-6) numeric differences vs Excel due to solver stopping criteria
    # and floating-point details. Keep the default global tolerance tight, but allow a slightly
    # looser tag-specific override for known-sensitive areas.
    tag_abs_tol = list(args.tag_abs_tol)
    tag_rel_tol = list(args.tag_rel_tol)
    tag_abs_tol.append("odd_coupon=1e-6")
    tag_rel_tol.append("odd_coupon=1e-6")

    repo_root = Path(__file__).resolve().parents[2]
    use_cargo_agent = (
        os.name != "nt"
        and shutil.which("bash") is not None
        and (repo_root / "scripts" / "cargo_agent.sh").is_file()
    )
    env = os.environ.copy()
    # `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml`. Some environments set it
    # globally (often to `stable`), which would bypass the pinned toolchain and reintroduce
    # drift when running `cargo` fallbacks.
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
        rustc_wrapper = (
            (shutil.which("env") or "env")
            if os.name != "nt"
            else ""
        )

    rustc_workspace_wrapper = env.get("RUSTC_WORKSPACE_WRAPPER")
    if rustc_workspace_wrapper is None:
        rustc_workspace_wrapper = env.get("CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER")
    if not rustc_workspace_wrapper:
        rustc_workspace_wrapper = (
            (shutil.which("env") or "env")
            if os.name != "nt"
            else ""
        )

    env["RUSTC_WRAPPER"] = rustc_wrapper
    env["RUSTC_WORKSPACE_WRAPPER"] = rustc_workspace_wrapper
    env["CARGO_BUILD_RUSTC_WRAPPER"] = rustc_wrapper
    env["CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER"] = rustc_workspace_wrapper

    engine_cmd = _build_engine_cmd(
        cases_path=cases_path,
        actual_path=actual_path,
        max_cases=args.max_cases,
        include_tags=include_tags,
        exclude_tags=exclude_tags,
        use_cargo_agent=use_cargo_agent,
    )

    compare_cmd = _build_compare_cmd(
        cases_path=cases_path,
        expected_path=expected_path,
        actual_path=actual_path,
        report_path=report_path,
        max_cases=args.max_cases,
        include_tags=include_tags,
        exclude_tags=exclude_tags,
        max_mismatch_rate=args.max_mismatch_rate,
        abs_tol=args.abs_tol,
        rel_tol=args.rel_tol,
        tag_abs_tol=tag_abs_tol,
        tag_rel_tol=tag_rel_tol,
    )

    if args.dry_run:
        def fmt(cmd: list[str]) -> str:
            # Print a shell-ready representation (useful for copy/paste).
            return " ".join(shlex.quote(part) for part in cmd)

        matched, selected = _count_selected_cases(
            cases_path=cases_path,
            include_tags=include_tags,
            exclude_tags=exclude_tags,
            max_cases=args.max_cases,
        )

        print(f"cases: {cases_path}")
        print(f"expected: {expected_path}")
        print(f"actual: {actual_path}")
        print(f"report: {report_path}")
        print(f"cases after tag filtering: {matched}")
        print(f"cases selected: {selected}")
        print()
        print("engine_cmd:")
        print(fmt(engine_cmd))
        print()
        print("compare_cmd:")
        print(fmt(compare_cmd))
        return 0

    subprocess.run(engine_cmd, check=True, cwd=repo_root, env=env)

    proc = subprocess.run(compare_cmd)

    # Produce a markdown summary alongside the JSON report for easy viewing in CI.
    try:
        report_payload = json.loads(report_path.read_text(encoding="utf-8"))
        summary = report_payload.get("summary", {}) if isinstance(report_payload, dict) else {}
        md_path = Path(args.report_md)

        lines: list[str] = []
        lines.append("# Excel oracle compatibility report")
        lines.append("")
        lines.append(f"* Cases corpus: `{cases_path}`")
        lines.append(f"* Expected dataset: `{expected_path}`")
        lines.append(f"* Actual dataset: `{actual_path}`")
        lines.append(f"* Report: `{report_path}`")
        lines.append(f"* Total cases: {summary.get('totalCases')}")
        lines.append(f"* Mismatches: {summary.get('mismatches')}")
        lines.append(f"* Mismatch rate: {summary.get('mismatchRate')}")
        lines.append(f"* Max mismatch rate: {summary.get('maxMismatchRate')}")
        cases_sha = summary.get("casesSha256")
        if isinstance(cases_sha, str) and cases_sha:
            lines.append(f"* casesSha256: `{cases_sha}`")
        lines.append(f"* absTol: {summary.get('absTol')}")
        lines.append(f"* relTol: {summary.get('relTol')}")
        tag_abs_tol = summary.get("tagAbsTol")
        if isinstance(tag_abs_tol, dict) and tag_abs_tol:
            parts = [f"{k}={v}" for k, v in sorted(tag_abs_tol.items())]
            lines.append(f"* tagAbsTol: {', '.join(parts)}")
        else:
            lines.append("* tagAbsTol: <none>")
        tag_rel_tol = summary.get("tagRelTol")
        if isinstance(tag_rel_tol, dict) and tag_rel_tol:
            parts = [f"{k}={v}" for k, v in sorted(tag_rel_tol.items())]
            lines.append(f"* tagRelTol: {', '.join(parts)}")
        else:
            lines.append("* tagRelTol: <none>")
        include_tags_summary = summary.get("includeTags")
        if isinstance(include_tags_summary, list) and include_tags_summary:
            lines.append(f"* Include tags: {', '.join(str(t) for t in include_tags_summary)}")
        else:
            lines.append("* Include tags: <all>")

        exclude_tags_summary = summary.get("excludeTags")
        if isinstance(exclude_tags_summary, list) and exclude_tags_summary:
            lines.append(f"* Exclude tags: {', '.join(str(t) for t in exclude_tags_summary)}")
        else:
            lines.append("* Exclude tags: <none>")

        max_cases_summary = summary.get("maxCases")
        if isinstance(max_cases_summary, int) and max_cases_summary > 0:
            lines.append(f"* Max cases: {max_cases_summary}")
        else:
            lines.append("* Max cases: all")
        lines.append("")

        # Expected dataset provenance: explicitly surface whether the dataset is real Excel or a
        # synthetic baseline, and include the embedded syntheticSource metadata so CI logs/artifacts
        # can't be misinterpreted.
        if expected_info is not None:
            lines.append("## Expected dataset provenance")
            lines.append("")
            if expected_info.is_synthetic and expected_info.synthetic_source:
                synth = expected_info.synthetic_source
                lines.append("**WARNING:** Expected dataset is a synthetic baseline (not generated by Excel).")
                lines.append("")
                lines.append("| Field | Value |")
                lines.append("| --- | --- |")
                lines.append(f"| `syntheticSource.kind` | `{synth.get('kind', 'unknown')}` |")
                lines.append(f"| `syntheticSource.version` | `{synth.get('version', 'unknown')}` |")
                lines.append(f"| `syntheticSource.os` | `{synth.get('os', 'unknown')}` |")
                lines.append(f"| `syntheticSource.arch` | `{synth.get('arch', 'unknown')}` |")
                lines.append(f"| `syntheticSource.caseSet` | `{synth.get('caseSet', 'unknown')}` |")
                lines.append("")
            else:
                lines.append("Expected dataset is from real Excel.")
                lines.append("")

            if expected_info.patches:
                lines.append(
                    f"Expected dataset contains {len(expected_info.patches)} patch entries under `source.patches`."
                )
                lines.append("")

        reason_counts = summary.get("reasonCounts")
        if isinstance(reason_counts, dict) and reason_counts:
            lines.append("## Reason counts")
            lines.append("")
            lines.append("| Reason | Count |")
            lines.append("| --- | ---: |")
            for reason, count in sorted(
                ((k, v) for k, v in reason_counts.items() if isinstance(k, str) and isinstance(v, int)),
                key=lambda kv: (-kv[1], kv[0]),
            ):
                lines.append(f"| `{reason}` | {count} |")
            lines.append("")

        tag_summary = summary.get("tagSummary")
        if isinstance(tag_summary, list) and tag_summary:
            lines.append("## Tag summary")
            lines.append("")
            lines.append("| Tag | Passes | Mismatches | Total | Mismatch rate |")
            lines.append("| --- | ---: | ---: | ---: | ---: |")
            for row in tag_summary[:50]:
                if not isinstance(row, dict):
                    continue
                tag = row.get("tag")
                passes = row.get("passes")
                mismatches = row.get("mismatches")
                total = row.get("total")
                rate = row.get("mismatchRate")
                lines.append(f"| {tag} | {passes} | {mismatches} | {total} | {rate:.4%} |")
            lines.append("")

        top_missing = summary.get("topMissingFunctions")
        if isinstance(top_missing, list) and top_missing:
            lines.append("## Top missing functions")
            lines.append("")
            for row in top_missing[:20]:
                if isinstance(row, dict) and "name" in row and "count" in row:
                    lines.append(f"* `{row['name']}`: {row['count']}")
            lines.append("")

        top_errors = summary.get("topActualErrorKinds")
        if isinstance(top_errors, list) and top_errors:
            lines.append("## Top actual error kinds (in mismatches)")
            lines.append("")
            for row in top_errors[:20]:
                if isinstance(row, dict) and "code" in row and "count" in row:
                    lines.append(f"* `{row['code']}`: {row['count']}")
            lines.append("")

        mismatches = report_payload.get("mismatches") if isinstance(report_payload, dict) else None
        if isinstance(mismatches, list) and mismatches:
            lines.append("## Sample mismatches")
            lines.append("")
            for m in mismatches[:10]:
                if not isinstance(m, dict):
                    continue
                tags = m.get("tags")
                tags_str = ""
                if isinstance(tags, list) and tags:
                    tags_str = ",".join(str(t) for t in tags)
                else:
                    tags_str = "<none>"
                expected_display = m.get("expectedDisplayText")
                if not isinstance(expected_display, str) or not expected_display:
                    expected_display = m.get("expected")
                actual_display = m.get("actualDisplayText")
                if not isinstance(actual_display, str) or not actual_display:
                    actual_display = m.get("actual")
                mismatch_detail = m.get("mismatchDetail")
                detail_str = ""
                if mismatch_detail is not None:
                    try:
                        detail_str = json.dumps(mismatch_detail, ensure_ascii=False)
                    except Exception:
                        detail_str = str(mismatch_detail)
                abs_diff = m.get("absDiff")
                rel_diff = m.get("relDiff")
                diff_str = ""
                if isinstance(abs_diff, (int, float)) or isinstance(rel_diff, (int, float)):
                    diff_str = f" absDiff={abs_diff} relDiff={rel_diff}"
                if detail_str:
                    diff_str += f" detail={detail_str}"
                diff_part = f"`{diff_str.strip()}` " if diff_str.strip() else ""
                lines.append(
                    f"* `{m.get('caseId')}` `{m.get('reason')}` "
                    f"`tags={tags_str}` "
                    f"`expected={expected_display}` "
                    f"`actual={actual_display}` "
                    f"{diff_part}"
                    f"`{m.get('formula')}`"
                )
            lines.append("")

        md_path.parent.mkdir(parents=True, exist_ok=True)
        md_path.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")
    except Exception:
        # Don't fail the gate if the summary couldn't be generated (the compare step already
        # enforces correctness via exit code + JSON report).
        pass

    return proc.returncode


if __name__ == "__main__":
    raise SystemExit(main())
