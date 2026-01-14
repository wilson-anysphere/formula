#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable


DEFAULT_LIMIT_MB = 50


@dataclass(frozen=True)
class Artifact:
    path: Path
    size_bytes: int


def _is_truthy_env(val: str | None) -> bool:
    if val is None:
        return False
    return val.strip().lower() in {"1", "true", "yes", "y", "on"}


def _parse_limit_mb(raw: str | None) -> int:
    if raw is None or not raw.strip():
        return DEFAULT_LIMIT_MB
    try:
        mb = int(float(raw))
    except ValueError as exc:  # noqa: PERF203
        raise ValueError(f"Invalid FORMULA_BUNDLE_SIZE_LIMIT_MB={raw!r} (expected a number)") from exc
    if mb <= 0:
        raise ValueError(f"Invalid FORMULA_BUNDLE_SIZE_LIMIT_MB={raw!r} (must be > 0)")
    return mb


def _human_bytes(size_bytes: int) -> str:
    size = float(size_bytes)
    units = ["B", "KB", "MB", "GB", "TB"]
    for unit in units:
        if size < 1000 or unit == units[-1]:
            if unit == "B":
                return f"{int(size)} {unit}"
            return f"{size:.1f} {unit}"
        size /= 1000
    return f"{size_bytes} B"


def _candidate_target_dirs(repo_root: Path) -> list[Path]:
    candidates: list[Path] = []

    # Respect `CARGO_TARGET_DIR` when set (some CI/caching setups override it). Cargo interprets
    # relative paths relative to the working directory used for the build (repo root in CI).
    env_target = os.environ.get("CARGO_TARGET_DIR")
    if env_target:
        p = Path(env_target)
        if not p.is_absolute():
            p = repo_root / p
        if p.is_dir():
            candidates.append(p)

    # Common locations:
    # - standalone Tauri app: apps/desktop/src-tauri/target
    # - builds from apps/desktop: apps/desktop/target
    # - workspace build: target/
    for p in (
        repo_root / "apps" / "desktop" / "src-tauri" / "target",
        repo_root / "apps" / "desktop" / "target",
        repo_root / "target",
    ):
        if p.is_dir():
            candidates.append(p)

    # Fallback: look for any src-tauri/target directories (but avoid scanning node_modules).
    if not candidates:
        for src_tauri in _find_src_tauri_dirs(repo_root):
            target_dir = src_tauri / "target"
            if target_dir.is_dir():
                candidates.append(target_dir)

    # De-dupe while preserving order.
    seen: set[Path] = set()
    uniq: list[Path] = []
    for c in candidates:
        try:
            key = c.resolve()
        except FileNotFoundError:
            key = c
        if key in seen:
            continue
        seen.add(key)
        uniq.append(c)
    return uniq


def _find_src_tauri_dirs(repo_root: Path) -> Iterable[Path]:
    """
    Best-effort discovery for Tauri projects without scanning huge directories.
    """
    skip_dirnames = {
        "node_modules",
        ".git",
        ".cargo",
        ".pnpm-store",
        ".turbo",
        ".cache",
        ".vite",
        "dist",
        "build",
        "coverage",
        "target",
        "security-report",
        "test-results",
        "playwright-report",
    }
    # Avoid traversing arbitrarily deep trees when `repo_root` contains extracted artifacts or
    # build output. A Tauri project should be relatively shallow in the repo (e.g.
    # apps/desktop/src-tauri/tauri.conf.json).
    max_depth = 8
    for root, dirs, files in os.walk(repo_root):
        dirs[:] = [d for d in dirs if d not in skip_dirnames]
        try:
            depth = len(Path(root).relative_to(repo_root).parts)
        except ValueError:
            depth = max_depth
        if depth >= max_depth:
            dirs[:] = []
        if "tauri.conf.json" in files:
            yield Path(root)


def _find_bundle_dirs(target_dir: Path) -> list[Path]:
    # Cover:
    # - target/release/bundle
    # - target/<triple>/release/bundle
    #
    # Avoid recursive `**/release/bundle` globs: Cargo target directories can be large, and
    # recursive globbing can be surprisingly slow in CI. Bundles are always emitted at:
    # - <target_dir>/release/bundle
    # - <target_dir>/<triple>/release/bundle
    bundle_dirs: list[Path] = []
    for pattern in ("release/bundle", "*/release/bundle"):
        for p in target_dir.glob(pattern):
            if p.is_dir():
                bundle_dirs.append(p)
    # De-dupe
    seen: set[Path] = set()
    uniq: list[Path] = []
    for b in bundle_dirs:
        key = b.resolve()
        if key in seen:
            continue
        seen.add(key)
        uniq.append(b)
    return uniq


def _iter_bundle_files(bundle_dir: Path) -> Iterable[Path]:
    # Don't traverse huge app bundles / debug symbol directories; we only want release artifacts.
    skip_suffixes = (".app", ".dSYM", ".framework", ".xcframework")
    for root, dirs, files in os.walk(bundle_dir):
        dirs[:] = [d for d in dirs if not d.endswith(skip_suffixes)]
        for f in files:
            yield Path(root) / f


def _is_release_asset(path: Path) -> bool:
    name = path.name

    # Primary installer/bundle types we care about for size budgets.
    if name.endswith((".dmg", ".msi", ".exe", ".AppImage", ".deb", ".rpm", ".pkg", ".zip")):
        return True
    if name.endswith((".tar.gz", ".tgz")):
        return True

    # Tauri updater signatures (paired with an artifact).
    if name.endswith(".sig"):
        base = name[: -len(".sig")]
        return base.endswith((".dmg", ".msi", ".exe", ".AppImage", ".deb", ".rpm", ".pkg", ".zip", ".tar.gz", ".tgz"))

    # Tauri updater metadata (small, but nice to show).
    if name.endswith(".json") and name in {"latest.json", "update.json"}:
        return True

    return False


def _gather_artifacts(bundle_dirs: list[Path]) -> list[Artifact]:
    artifacts: list[Artifact] = []
    for bundle_dir in bundle_dirs:
        for path in _iter_bundle_files(bundle_dir):
            if not path.is_file():
                continue
            if not _is_release_asset(path):
                continue
            artifacts.append(Artifact(path=path, size_bytes=path.stat().st_size))

    # Sort biggest-first (helps spot offenders quickly).
    artifacts.sort(key=lambda a: a.size_bytes, reverse=True)
    return artifacts


def _render_markdown(artifacts: list[Artifact], limit_mb: int, enforce: bool, repo_root: Path) -> str:
    limit_bytes = limit_mb * 1000 * 1000

    lines: list[str] = []
    runner_os = os.environ.get("RUNNER_OS", "").strip()

    heading = "## Desktop installer artifact sizes"
    if runner_os:
        heading += f" ({runner_os})"
    lines.append(heading)
    lines.append("")
    lines.append(f"- Limit: **{limit_mb} MB** per artifact")
    lines.append(
        f"- Enforcement: **{'enabled' if enforce else 'disabled'}**"
        + (" (set `FORMULA_ENFORCE_BUNDLE_SIZE=1` to fail on oversize)" if not enforce else "")
    )
    lines.append("")

    if not artifacts:
        lines.append("_No installer/bundle artifacts found._")
        lines.append("")
        return "\n".join(lines)

    lines.append("| Artifact | Size | Over limit |")
    lines.append("| --- | ---: | :---: |")

    over_limit: list[Artifact] = []
    for art in artifacts:
        rel = str(art.path)
        try:
            rel = str(art.path.relative_to(repo_root))
        except ValueError:
            pass
        over = art.size_bytes > limit_bytes
        if over:
            over_limit.append(art)
        lines.append(f"| `{rel}` | {_human_bytes(art.size_bytes)} | {'YES' if over else ''} |")

    lines.append("")
    if over_limit:
        lines.append(f"Artifacts over **{limit_mb} MB**: {len(over_limit)}/{len(artifacts)}")
        lines.append("")
        for art in over_limit:
            rel = str(art.path)
            try:
                rel = str(art.path.relative_to(repo_root))
            except ValueError:
                pass
            lines.append(f"- `{rel}` ({_human_bytes(art.size_bytes)})")
        lines.append("")

        if not enforce:
            lines.append("Note: oversize artifacts do not fail the build unless enforcement is enabled.")
            lines.append("")

    return "\n".join(lines)


def _append_step_summary(markdown: str) -> None:
    summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
    if not summary_path:
        return
    with open(summary_path, "a", encoding="utf-8", errors="replace") as f:
        f.write(markdown)
        f.write("\n")


def _report_path_str(path: Path, repo_root: Path) -> str:
    """
    Return a repo-relative path string when possible (for stable CI output).
    """
    try:
        resolved = path.resolve()
    except FileNotFoundError:
        resolved = path
    try:
        return resolved.relative_to(repo_root.resolve()).as_posix()
    except (ValueError, FileNotFoundError):
        try:
            return path.relative_to(repo_root).as_posix()
        except ValueError:
            return path.as_posix()


def _build_json_report(
    artifacts: list[Artifact],
    *,
    bundle_dirs: list[Path] | None = None,
    limit_mb: int,
    enforce: bool,
    repo_root: Path,
) -> dict[str, Any]:
    limit_bytes = limit_mb * 1000 * 1000
    runner_os = os.environ.get("RUNNER_OS", "").strip()

    over_limit_count = 0
    artifact_rows: list[dict[str, Any]] = []
    for art in artifacts:
        over = art.size_bytes > limit_bytes
        if over:
            over_limit_count += 1
        artifact_rows.append(
            {
                "path": _report_path_str(art.path, repo_root=repo_root),
                "size_bytes": art.size_bytes,
                "size_mb": round(art.size_bytes / 1_000_000, 3),
                "over_limit": over,
            }
        )

    report: dict[str, Any] = {
        "limit_mb": limit_mb,
        "enforce": enforce,
        "bundle_dirs": [
            _report_path_str(d, repo_root=repo_root)
            for d in (bundle_dirs or [])
        ],
        "artifacts": artifact_rows,
        "total_artifacts": len(artifacts),
        "over_limit_count": over_limit_count,
    }
    if runner_os:
        report["runner_os"] = runner_os
    return report


def _write_json_report(json_path: Path | None, report: dict[str, Any]) -> bool:
    if json_path is None:
        return True
    try:
        json_path.parent.mkdir(parents=True, exist_ok=True)
        with open(json_path, "w", encoding="utf-8", newline="\n") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
            f.write("\n")
        return True
    except OSError as exc:
        print(f"bundle-size: ERROR Failed to write JSON report to {json_path}: {exc}", file=sys.stderr)
        return False


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Report (and optionally enforce) desktop installer/bundle artifact sizes after a release build."
    )
    parser.add_argument(
        "--bundle-dir",
        type=Path,
        action="append",
        default=None,
        help="Explicit bundle directory to scan (can be specified multiple times). "
        "If omitted, the script will auto-detect bundle dirs under <target>/release/bundle and <target>/*/release/bundle.",
    )
    parser.add_argument(
        "--limit-mb",
        type=int,
        default=None,
        help=f"Override size limit in MB (default: {DEFAULT_LIMIT_MB}, or env FORMULA_BUNDLE_SIZE_LIMIT_MB).",
    )
    parser.add_argument(
        "--enforce",
        action="store_true",
        default=False,
        help="Fail if any artifact exceeds the limit (also enabled by env FORMULA_ENFORCE_BUNDLE_SIZE=1).",
    )
    parser.add_argument(
        "--json",
        nargs="?",
        type=Path,
        const=Path("desktop-bundle-size-report.json"),
        default=None,
        help="Write a machine-readable JSON report to PATH "
        "(default: desktop-bundle-size-report.json; also via env FORMULA_BUNDLE_SIZE_JSON_PATH).",
    )
    args = parser.parse_args()

    repo_root = Path.cwd()
    enforce = args.enforce or _is_truthy_env(os.environ.get("FORMULA_ENFORCE_BUNDLE_SIZE"))
    json_path = args.json
    if json_path is None:
        raw_json_path = os.environ.get("FORMULA_BUNDLE_SIZE_JSON_PATH", "").strip()
        if raw_json_path:
            json_path = Path(raw_json_path)
    try:
        limit_mb = args.limit_mb if args.limit_mb is not None else _parse_limit_mb(os.environ.get("FORMULA_BUNDLE_SIZE_LIMIT_MB"))
    except ValueError as exc:
        print(f"bundle-size: ERROR {exc}", file=sys.stderr)
        return 2

    bundle_dirs: list[Path] = []
    target_dirs: list[Path] = []
    if args.bundle_dir:
        bundle_dirs = [d for d in args.bundle_dir if d.is_dir()]
        missing = [str(d) for d in args.bundle_dir if not d.is_dir()]
        if missing:
            print(f"bundle-size: ERROR --bundle-dir not found: {', '.join(missing)}", file=sys.stderr)
            _write_json_report(
                json_path,
                _build_json_report(
                    [],
                    bundle_dirs=bundle_dirs,
                    limit_mb=limit_mb,
                    enforce=enforce,
                    repo_root=repo_root,
                ),
            )
            return 2
    else:
        target_dirs = _candidate_target_dirs(repo_root)
        for t in target_dirs:
            bundle_dirs.extend(_find_bundle_dirs(t))

    # De-dupe.
    seen: set[Path] = set()
    uniq: list[Path] = []
    for b in bundle_dirs:
        key = b.resolve()
        if key in seen:
            continue
        seen.add(key)
        uniq.append(b)
    bundle_dirs = uniq

    if not bundle_dirs:
        expected_examples = [
            repo_root / "target" / "release" / "bundle",
            repo_root / "apps" / "desktop" / "src-tauri" / "target" / "release" / "bundle",
            repo_root / "apps" / "desktop" / "target" / "release" / "bundle",
        ]
        # If CARGO_TARGET_DIR is configured, include it in the example list.
        env_target = os.environ.get("CARGO_TARGET_DIR", "").strip()
        if env_target:
            p = Path(env_target)
            if not p.is_absolute():
                p = repo_root / p
            expected_examples.insert(0, p / "release" / "bundle")

        expected_joined = "\n".join(f"  - {p}" for p in expected_examples)
        candidates_joined = (
            ", ".join(_report_path_str(p, repo_root=repo_root) for p in target_dirs)
            if target_dirs
            else "(none found)"
        )
        msg = (
            "bundle-size: ERROR No Tauri bundle directories found.\n"
            f"Searched for: <target>/release/bundle and <target>/*/release/bundle (from {repo_root})\n"
            f"Candidate target dirs: {candidates_joined}\n"
            "Expected a directory like one of:\n"
            f"{expected_joined}\n"
            "Hint: build the desktop app with `(cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)` before running this script."
        )
        print(msg, file=sys.stderr)
        # Still write a summary so CI logs show something useful.
        md = _render_markdown([], limit_mb=limit_mb, enforce=enforce, repo_root=repo_root)
        _append_step_summary(md)
        _write_json_report(
            json_path,
            _build_json_report(
                [],
                bundle_dirs=[],
                limit_mb=limit_mb,
                enforce=enforce,
                repo_root=repo_root,
            ),
        )
        return 1

    artifacts = _gather_artifacts(bundle_dirs)

    md = _render_markdown(artifacts, limit_mb=limit_mb, enforce=enforce, repo_root=repo_root)
    print(md)
    sys.stdout.flush()
    _append_step_summary(md)
    if not _write_json_report(
        json_path,
        _build_json_report(
            artifacts,
            bundle_dirs=bundle_dirs,
            limit_mb=limit_mb,
            enforce=enforce,
            repo_root=repo_root,
        ),
    ):
        return 2

    if not enforce:
        return 0

    limit_bytes = limit_mb * 1000 * 1000
    offenders = [a for a in artifacts if a.size_bytes > limit_bytes]
    if offenders:
        print(
            f"bundle-size: ERROR {len(offenders)} artifact(s) exceed {limit_mb} MB "
            f"(set FORMULA_BUNDLE_SIZE_LIMIT_MB to adjust).",
            file=sys.stderr,
        )
        for a in offenders:
            rel = str(a.path)
            try:
                rel = str(a.path.relative_to(repo_root))
            except ValueError:
                pass
            print(f"bundle-size: - {rel} ({_human_bytes(a.size_bytes)})", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
