#!/usr/bin/env python3

from __future__ import annotations

import argparse
import gzip
import json
import os
import subprocess
import sys
import tarfile
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any


MB_BYTES = 1_000_000


@dataclass(frozen=True)
class SizedPath:
    path: Path
    size_bytes: int

    @property
    def size_mb(self) -> float:
        return self.size_bytes / MB_BYTES


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


def _parse_optional_limit_mb(env_key: str) -> float | None:
    raw = os.environ.get(env_key)
    if raw is None or not raw.strip():
        return None
    try:
        limit_mb = float(raw)
    except ValueError as exc:  # noqa: PERF203
        raise ValueError(f"Invalid {env_key}={raw!r} (expected a number)") from exc
    if limit_mb <= 0:
        raise ValueError(f"Invalid {env_key}={raw!r} (must be > 0)")
    return limit_mb


def _dir_size_bytes(path: Path) -> int:
    total = 0
    for root, _dirs, files in os.walk(path):
        for name in files:
            p = Path(root) / name
            try:
                total += p.stat().st_size
            except FileNotFoundError:
                # Racy filesystem changes shouldn't kill the report; just skip.
                continue
    return total


def _tar_gz_dir_size_bytes(path: Path) -> int:
    """
    Approximate compressed download cost with a deterministic tar+gzip of the directory.

    Notes:
    - We normalize tar metadata (mtime/uid/gid) to reduce run-to-run noise.
    - We walk the directory in sorted order to keep gzip output stable.
    """

    def normalize(ti: tarfile.TarInfo) -> tarfile.TarInfo:
        ti.uid = 0
        ti.gid = 0
        ti.uname = ""
        ti.gname = ""
        ti.mtime = 0
        ti.pax_headers = {}
        return ti

    def iter_paths_sorted(root: Path) -> list[Path]:
        paths: list[Path] = []
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames.sort()
            filenames.sort()
            current = Path(dirpath)
            # Include directories so empty dirs are represented.
            for d in dirnames:
                paths.append(current / d)
            for f in filenames:
                paths.append(current / f)
        return paths

    tmp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(prefix="desktop-dist-", suffix=".tar.gz", delete=False) as tmp:
            tmp_path = Path(tmp.name)

        base = path.name
        with open(tmp_path, "wb") as f:
            with gzip.GzipFile(fileobj=f, mode="wb", compresslevel=6, mtime=0) as gz:
                # Streaming mode so we don't need seeking.
                with tarfile.open(fileobj=gz, mode="w|") as tar:
                    # Root directory entry.
                    tar.add(path, arcname=base, recursive=False, filter=normalize)
                    for p in iter_paths_sorted(path):
                        rel = p.relative_to(path).as_posix()
                        tar.add(p, arcname=f"{base}/{rel}", recursive=False, filter=normalize)

        return tmp_path.stat().st_size
    finally:
        if tmp_path is not None:
            try:
                tmp_path.unlink()
            except FileNotFoundError:
                pass


def _append_step_summary(markdown: str) -> None:
    summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
    if not summary_path:
        return
    try:
        with open(summary_path, "a", encoding="utf-8", errors="replace") as f:
            f.write(markdown)
            f.write("\n")
    except OSError:
        # Don't fail the report if GitHub step summary can't be written.
        return


def _append_error_summary(message: str, *, hints: list[str] | None = None) -> None:
    """
    Best-effort: if running in GitHub Actions, surface failures in the step summary.
    """
    lines = [
        "## Desktop size report",
        "",
        f"**ERROR:** {message}",
        "",
    ]
    if hints:
        lines.append("**Hint:**")
        lines.append("")
        for hint in hints:
            lines.append(f"- {hint}")
        lines.append("")
    md = "\n".join(lines)
    _append_step_summary(md)


def _relpath(path: Path, repo_root: Path) -> str:
    try:
        return path.relative_to(repo_root).as_posix()
    except ValueError:
        return path.as_posix()


def _render_markdown(
    *,
    binary: SizedPath,
    dist: SizedPath,
    dist_gzip: SizedPath | None,
    binary_limit_mb: float | None,
    dist_limit_mb: float | None,
    repo_root: Path,
) -> str:
    runner_os = os.environ.get("RUNNER_OS", "").strip()
    heading = "## Desktop size report"
    if runner_os:
        heading += f" ({runner_os})"

    def limit_str(limit: float | None) -> str:
        if limit is None:
            return "_disabled_"
        # Show whole numbers without decimals; otherwise show one decimal.
        if limit.is_integer():
            return f"**{int(limit)} MB**"
        return f"**{limit:.1f} MB**"

    lines: list[str] = []
    lines.append(heading)
    lines.append("")
    lines.append(f"- Binary limit (`FORMULA_DESKTOP_BINARY_SIZE_LIMIT_MB`): {limit_str(binary_limit_mb)}")
    lines.append(f"- Dist limit (`FORMULA_DESKTOP_DIST_SIZE_LIMIT_MB`): {limit_str(dist_limit_mb)}")
    lines.append("")
    lines.append("| Artifact | Size (bytes) | Size (MB) | Over limit |")
    lines.append("| --- | ---: | ---: | :---: |")

    def over(size_bytes: int, limit_mb: float | None) -> str:
        if limit_mb is None:
            return ""
        limit_bytes = int(round(limit_mb * MB_BYTES))
        return "YES" if size_bytes > limit_bytes else ""

    lines.append(
        "| `{}` | {} | {} | {} |".format(
            _relpath(binary.path, repo_root),
            binary.size_bytes,
            f"{binary.size_mb:.1f}",
            over(binary.size_bytes, binary_limit_mb),
        )
    )
    lines.append(
        "| `{}` | {} | {} | {} |".format(
            _relpath(dist.path, repo_root),
            dist.size_bytes,
            f"{dist.size_mb:.1f}",
            over(dist.size_bytes, dist_limit_mb),
        )
    )
    if dist_gzip is not None:
        lines.append(
            "| `{}` (tar.gz) | {} | {} |  |".format(
                _relpath(dist_gzip.path, repo_root),
                dist_gzip.size_bytes,
                f"{dist_gzip.size_mb:.1f}",
            )
        )

    lines.append("")

    # Explicitly call out offenders so failures are obvious in CI summaries.
    offenders: list[str] = []
    if binary_limit_mb is not None and binary.size_bytes > int(round(binary_limit_mb * MB_BYTES)):
        offenders.append(
            f"- Binary `{_relpath(binary.path, repo_root)}`: {_human_bytes(binary.size_bytes)} > {binary_limit_mb} MB"
        )
    if dist_limit_mb is not None and dist.size_bytes > int(round(dist_limit_mb * MB_BYTES)):
        offenders.append(f"- Dist `{_relpath(dist.path, repo_root)}`: {_human_bytes(dist.size_bytes)} > {dist_limit_mb} MB")
    if offenders:
        lines.append("**Size limits exceeded:**")
        lines.append("")
        lines.extend(offenders)
        lines.append("")

    return "\n".join(lines)


def _default_desktop_binary_path() -> Path:
    if os.name == "nt":
        return Path("target/release/formula-desktop.exe")
    return Path("target/release/formula-desktop")


def _cargo_target_directory(repo_root: Path) -> Path | None:
    """
    Resolve Cargo's target directory via `cargo metadata`.

    This respects `CARGO_TARGET_DIR` when set and is more reliable than assuming
    `<repo>/target`.
    """
    try:
        # `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml` pin. Some environments set it
        # globally (often to `stable`), which would bypass the pinned toolchain and reintroduce drift
        # for any cargo subprocess calls in this report.
        env = os.environ.copy()
        if env.get("RUSTUP_TOOLCHAIN") and (repo_root / "rust-toolchain.toml").is_file():
            env.pop("RUSTUP_TOOLCHAIN", None)

        cp = subprocess.run(
            ["cargo", "metadata", "--no-deps", "--format-version=1"],
            cwd=repo_root,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
        )
    except OSError:
        return None
    if cp.returncode != 0:
        return None
    try:
        meta = json.loads(cp.stdout)
    except json.JSONDecodeError:
        return None
    target_dir = meta.get("target_directory")
    if not target_dir:
        return None
    return Path(str(target_dir))


def _candidate_default_binary_paths(repo_root: Path) -> list[Path]:
    """
    Best-effort discovery for the built desktop binary in common locations.

    In this repo, workspace builds typically land in `<repo>/target`, but historical
    or non-workspace builds may place artifacts under `apps/desktop/src-tauri/target`.
    """
    exe = "formula-desktop.exe" if os.name == "nt" else "formula-desktop"

    base_dirs: list[Path] = []

    raw_target_dir = os.environ.get("CARGO_TARGET_DIR", "").strip()
    if raw_target_dir:
        td = Path(raw_target_dir)
        if not td.is_absolute():
            td = repo_root / td
        base_dirs.append(td)

    # Avoid invoking `cargo metadata` when the caller explicitly set `CARGO_TARGET_DIR`.
    # This keeps the default-binary discovery logic dependency-free in lightweight CI
    # guard jobs that do not install Rust.
    if not raw_target_dir:
        cargo_target_dir = _cargo_target_directory(repo_root)
        if cargo_target_dir is not None:
            base_dirs.append(cargo_target_dir)

    base_dirs.extend(
        [
            repo_root / "target",
            repo_root / "apps" / "desktop" / "src-tauri" / "target",
        ]
    )

    # De-dupe base dirs while preserving order.
    seen_base: set[Path] = set()
    uniq_base: list[Path] = []
    for p in base_dirs:
        try:
            key = p.resolve()
        except OSError:
            key = p
        if key in seen_base:
            continue
        seen_base.add(key)
        uniq_base.append(p)
    base_dirs = uniq_base

    candidates: list[Path] = []
    for base in base_dirs:
        candidates.append(base / "release" / exe)

    # Cross-compile / target-triple build outputs.
    for base in base_dirs:
        if not base.is_dir():
            continue
        try:
            entries = sorted(base.iterdir())
        except OSError:
            continue
        for ent in entries:
            if not ent.is_dir():
                continue
            # Avoid noise from common non-target-triple directories in `target/`.
            name = ent.name
            if name in {"debug", "release", "cargo-home", "perf-home", "build", "deps", "incremental", ".fingerprint"}:
                continue
            # Target triples generally contain hyphens (e.g. `x86_64-unknown-linux-gnu`).
            if "-" not in name:
                continue
            candidates.append(ent / "release" / exe)

    # De-dupe while preserving order.
    seen: set[Path] = set()
    uniq: list[Path] = []
    for p in candidates:
        key = p
        try:
            key = p.resolve()
        except OSError:
            key = p
        if key in seen:
            continue
        seen.add(key)
        uniq.append(p)
    return uniq


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Report lightweight desktop sizes (Rust desktop binary + Vite dist dir) without running `tauri build`.",
        epilog=(
            "Environment variables:\n"
            "  FORMULA_DESKTOP_BINARY_SIZE_LIMIT_MB  Fail if the desktop binary exceeds this budget (optional).\n"
            "  FORMULA_DESKTOP_DIST_SIZE_LIMIT_MB    Fail if the dist directory exceeds this budget (optional).\n"
            "  GITHUB_STEP_SUMMARY                   When set (GitHub Actions), append the markdown report to this file.\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--binary",
        type=Path,
        default=_default_desktop_binary_path(),
        help="Path to the built desktop binary (default: target/release/formula-desktop[.exe]).",
    )
    parser.add_argument(
        "--dist",
        type=Path,
        default=Path("apps/desktop/dist"),
        help="Path to the built desktop dist directory (default: apps/desktop/dist).",
    )
    parser.add_argument(
        "--gzip",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Also compute tar+gzip size of the dist directory to approximate download cost.",
    )
    parser.add_argument(
        "--json-out",
        type=Path,
        default=None,
        help="Optional path to write a JSON report.",
    )
    args = parser.parse_args()

    cwd = Path.cwd()
    # Prefer repo-root-relative reporting (stable in CI) even when invoked from a subdirectory.
    repo_root = Path(__file__).resolve().parent.parent

    try:
        binary_limit_mb = _parse_optional_limit_mb("FORMULA_DESKTOP_BINARY_SIZE_LIMIT_MB")
        dist_limit_mb = _parse_optional_limit_mb("FORMULA_DESKTOP_DIST_SIZE_LIMIT_MB")
    except ValueError as exc:
        msg = str(exc)
        print(f"desktop-size: ERROR {msg}", file=sys.stderr)
        _append_error_summary(msg)
        return 2

    binary_path = args.binary
    searched_binary_paths: list[Path] = []
    if not binary_path.is_absolute():
        cwd_candidate = cwd / binary_path
        repo_candidate = repo_root / binary_path
        searched_binary_paths.extend([cwd_candidate, repo_candidate])
        if cwd_candidate.is_file():
            binary_path = cwd_candidate
        elif repo_candidate.is_file():
            binary_path = repo_candidate
    else:
        searched_binary_paths.append(binary_path)

    dist_path = args.dist
    if not dist_path.is_absolute():
        if (cwd / dist_path).is_dir():
            dist_path = cwd / dist_path
        elif (repo_root / dist_path).is_dir():
            dist_path = repo_root / dist_path

    if not binary_path.is_file():
        # If the caller used the default binary path, try a few more common locations
        # (workspace target dir vs per-crate target dir vs target-triple paths).
        if args.binary == _default_desktop_binary_path():
            for candidate in _candidate_default_binary_paths(repo_root):
                searched_binary_paths.append(candidate)
                if candidate.is_file():
                    binary_path = candidate
                    break

    if not binary_path.is_file():
        msg = f"binary not found: {binary_path}"
        print(f"desktop-size: ERROR {msg}", file=sys.stderr)
        uniq_searched: list[str] = []
        seen: set[str] = set()
        for p in searched_binary_paths:
            rel = _relpath(p, repo_root)
            if rel in seen:
                continue
            seen.add(rel)
            uniq_searched.append(rel)
        searched = ", ".join(f"`{p}`" for p in uniq_searched[:6])
        if len(uniq_searched) > 6:
            searched += ", …"
        if searched:
            print(f"desktop-size: searched: {searched}", file=sys.stderr)
        _append_error_summary(
            msg,
            hints=[
                "Build the desktop binary: `cargo build -p formula-desktop-tauri --bin formula-desktop --features desktop --release --locked`",
                "If your checkout uses the historical package name: `cargo build -p desktop --bin formula-desktop --features desktop --release --locked`",
                "Or pass `--binary <path>` (default: `target/release/formula-desktop[.exe]`)",
                f"Searched: {searched}" if searched else "Searched: _(none)_",
            ],
        )
        return 2
    if not dist_path.is_dir():
        msg = f"dist directory not found: {dist_path}"
        print(f"desktop-size: ERROR {msg}", file=sys.stderr)
        _append_error_summary(
            msg,
            hints=[
                "Build desktop renderer assets: `pnpm build:desktop`",
                "Or pass `--dist <path>` (default: `apps/desktop/dist`)",
            ],
        )
        return 2

    binary = SizedPath(path=binary_path, size_bytes=binary_path.stat().st_size)
    dist = SizedPath(path=dist_path, size_bytes=_dir_size_bytes(dist_path))
    dist_gzip: SizedPath | None = None
    if args.gzip:
        dist_gzip = SizedPath(path=dist_path, size_bytes=_tar_gz_dir_size_bytes(dist_path))

    md = _render_markdown(
        binary=binary,
        dist=dist,
        dist_gzip=dist_gzip,
        binary_limit_mb=binary_limit_mb,
        dist_limit_mb=dist_limit_mb,
        repo_root=repo_root,
    )
    print(md)
    sys.stdout.flush()
    _append_step_summary(md)

    report: dict[str, Any] = {
        "binary": {
            "path": _relpath(binary.path, repo_root),
            "size_bytes": binary.size_bytes,
            "size_mb": round(binary.size_mb, 3),
            "over_limit": binary_limit_mb is not None and binary.size_bytes > int(round(binary_limit_mb * MB_BYTES)),
        },
        "dist": {
            "path": _relpath(dist.path, repo_root),
            "size_bytes": dist.size_bytes,
            "size_mb": round(dist.size_mb, 3),
            "over_limit": dist_limit_mb is not None and dist.size_bytes > int(round(dist_limit_mb * MB_BYTES)),
        },
        "dist_tar_gz": None,
        "limits_mb": {"binary": binary_limit_mb, "dist": dist_limit_mb},
    }
    runner_os = os.environ.get("RUNNER_OS", "").strip()
    if runner_os:
        report["runner_os"] = runner_os
    if dist_gzip is not None:
        report["dist_tar_gz"] = {
            "path": _relpath(dist_gzip.path, repo_root),
            "size_bytes": dist_gzip.size_bytes,
            "size_mb": round(dist_gzip.size_mb, 3),
        }

    if args.json_out is not None:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        with open(args.json_out, "w", encoding="utf-8", newline="\n") as f:
            json.dump(report, f, indent=2, sort_keys=True)
            f.write("\n")

    offenders: list[str] = []
    if binary_limit_mb is not None and binary.size_bytes > int(round(binary_limit_mb * MB_BYTES)):
        offenders.append(
            f"binary size {_human_bytes(binary.size_bytes)} exceeds limit {binary_limit_mb} MB "
            "(FORMULA_DESKTOP_BINARY_SIZE_LIMIT_MB)"
        )
    if dist_limit_mb is not None and dist.size_bytes > int(round(dist_limit_mb * MB_BYTES)):
        offenders.append(
            f"dist size {_human_bytes(dist.size_bytes)} exceeds limit {dist_limit_mb} MB "
            "(FORMULA_DESKTOP_DIST_SIZE_LIMIT_MB)"
        )

    if offenders:
        print("desktop-size: ERROR size limits exceeded:", file=sys.stderr)
        for msg in offenders:
            print(f"desktop-size: - {msg}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
