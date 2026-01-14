#!/usr/bin/env python3

from __future__ import annotations

import argparse
import gzip
import json
import os
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
    with open(summary_path, "a", encoding="utf-8", errors="replace") as f:
        f.write(markdown)
        f.write("\n")


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
        return str(path.relative_to(repo_root))
    except ValueError:
        return str(path)


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
    return "\n".join(lines)


def _default_desktop_binary_path() -> Path:
    if os.name == "nt":
        return Path("target/release/formula-desktop.exe")
    return Path("target/release/formula-desktop")


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
    if not binary_path.is_absolute():
        if (cwd / binary_path).is_file():
            binary_path = cwd / binary_path
        elif (repo_root / binary_path).is_file():
            binary_path = repo_root / binary_path

    dist_path = args.dist
    if not dist_path.is_absolute():
        if (cwd / dist_path).is_dir():
            dist_path = cwd / dist_path
        elif (repo_root / dist_path).is_dir():
            dist_path = repo_root / dist_path

    if not binary_path.is_file():
        msg = f"binary not found: {binary_path}"
        print(f"desktop-size: ERROR {msg}", file=sys.stderr)
        _append_error_summary(
            msg,
            hints=[
                "Build the desktop binary: `cargo build -p formula-desktop-tauri --features desktop --release --locked`",
                "Or pass `--binary <path>` (default: `target/release/formula-desktop[.exe]`)",
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
        },
        "dist": {
            "path": _relpath(dist.path, repo_root),
            "size_bytes": dist.size_bytes,
            "size_mb": round(dist.size_mb, 3),
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
