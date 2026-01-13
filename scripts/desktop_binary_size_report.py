#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import os
import platform
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

DEFAULT_TOP = 50
DEFAULT_FEATURES = "desktop"
DEFAULT_BIN_NAME = "formula-desktop"
DEFAULT_DESKTOP_PKG_FALLBACK = "formula-desktop-tauri"


@dataclass(frozen=True)
class CmdResult:
    cmd: list[str]
    returncode: int
    stdout: str
    stderr: str


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


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _desktop_package_name(repo_root: Path) -> str:
    """
    The desktop Tauri shell has historically used both `desktop` and
    `formula-desktop-tauri` as the Cargo package name. Prefer reading the current
    name from `apps/desktop/src-tauri/Cargo.toml` to keep scripts stable.
    """
    cargo_toml = repo_root / "apps" / "desktop" / "src-tauri" / "Cargo.toml"
    if not cargo_toml.is_file():
        return DEFAULT_DESKTOP_PKG_FALLBACK

    in_package = False
    for raw in cargo_toml.read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("[") and line.endswith("]"):
            in_package = line == "[package]"
            continue
        if not in_package:
            continue
        if line.startswith("name"):
            # Accept `name = "..."` with arbitrary whitespace.
            parts = line.split("=", 1)
            if len(parts) != 2:
                continue
            rhs = parts[1].strip()
            if rhs.startswith('"') and '"' in rhs[1:]:
                return rhs.split('"', 2)[1]
    return DEFAULT_DESKTOP_PKG_FALLBACK


def _cargo_target_directory(repo_root: Path) -> Path:
    cp = subprocess.run(
        ["cargo", "metadata", "--no-deps", "--format-version=1"],
        cwd=repo_root,
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    meta = json.loads(cp.stdout)
    target_dir = meta.get("target_directory")
    if not target_dir:
        raise RuntimeError("cargo metadata did not return target_directory")
    return Path(target_dir)


def _binary_path(target_dir: Path, bin_name: str, target: str | None) -> Path:
    rel_dir = target_dir / "release"
    if target:
        rel_dir = target_dir / target / "release"
    exe = bin_name
    if sys.platform == "win32":
        exe += ".exe"
    return rel_dir / exe


def _run_capture(cmd: list[str], *, cwd: Path) -> CmdResult:
    cp = subprocess.run(
        cmd,
        cwd=cwd,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return CmdResult(cmd=cmd, returncode=cp.returncode, stdout=cp.stdout, stderr=cp.stderr)


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


def _render_cmd(cmd: list[str]) -> str:
    # Use shell-ish quoting for readability.
    out: list[str] = []
    for part in cmd:
        if any(c in part for c in (" ", "\t", "\n", '"', "'", "$")):
            out.append(json.dumps(part))
        else:
            out.append(part)
    return " ".join(out)


def _render_markdown(
    *,
    package: str,
    bin_name: str,
    features: str,
    target: str | None,
    target_dir: Path,
    bin_path: Path,
    bin_size_bytes: int | None,
    build_cmd: list[str] | None,
    crates_cmd: list[str] | None,
    crates_out: CmdResult | None,
    symbols_cmd: list[str] | None,
    symbols_out: CmdResult | None,
    llvm_size_cmd: list[str] | None,
    llvm_size_out: CmdResult | None,
    tool_note: str | None,
) -> str:
    lines: list[str] = []
    runner_os = os.environ.get("RUNNER_OS", "").strip() or platform.system()

    heading = "## Desktop Rust binary size breakdown"
    if runner_os:
        heading += f" ({runner_os})"
    lines.append(heading)
    lines.append("")

    lines.append(f"- Package: `{package}`")
    lines.append(f"- Binary: `{bin_name}` (profile: `release`)")
    lines.append(f"- Features: `{features}`")
    if target:
        lines.append(f"- Target: `{target}`")
    lines.append(f"- Target dir: `{target_dir}`")
    lines.append(f"- Binary path: `{bin_path}`")
    if bin_size_bytes is not None:
        lines.append(f"- Binary size: **{_human_bytes(bin_size_bytes)}** ({bin_size_bytes} bytes)")
    lines.append("")

    if tool_note:
        lines.append(tool_note)
        lines.append("")

    if build_cmd:
        lines.append("### Build command")
        lines.append("")
        lines.append("```bash")
        lines.append(_render_cmd(build_cmd))
        lines.append("```")
        lines.append("")

    if crates_cmd and crates_out:
        lines.append("### Top crates (cargo-bloat)")
        lines.append("")
        lines.append("Command:")
        lines.append(f"`{_render_cmd(crates_cmd)}`")
        lines.append("")
        lines.append("```text")
        combined = (crates_out.stdout + ("\n" + crates_out.stderr if crates_out.stderr.strip() else "")).rstrip()
        lines.append(combined or "<no output>")
        lines.append("```")
        lines.append("")

    if symbols_cmd and symbols_out:
        lines.append("### Top symbols (cargo-bloat)")
        lines.append("")
        lines.append("Command:")
        lines.append(f"`{_render_cmd(symbols_cmd)}`")
        lines.append("")
        lines.append("```text")
        combined = (symbols_out.stdout + ("\n" + symbols_out.stderr if symbols_out.stderr.strip() else "")).rstrip()
        lines.append(combined or "<no output>")
        lines.append("```")
        lines.append("")

    if llvm_size_cmd and llvm_size_out:
        lines.append("### Section sizes (llvm-size)")
        lines.append("")
        lines.append("Command:")
        lines.append(f"`{_render_cmd(llvm_size_cmd)}`")
        lines.append("")
        lines.append("```text")
        combined = (llvm_size_out.stdout + ("\n" + llvm_size_out.stderr if llvm_size_out.stderr.strip() else "")).rstrip()
        lines.append(combined or "<no output>")
        lines.append("```")
        lines.append("")

    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate a binary size breakdown report for the desktop Rust/Tauri shell."
    )
    parser.add_argument(
        "--no-build",
        action="store_true",
        default=False,
        help="Assume the release binary is already built and skip `cargo build`.",
    )
    parser.add_argument(
        "--top",
        type=int,
        default=DEFAULT_TOP,
        help=f"How many entries to show from cargo-bloat (default: {DEFAULT_TOP}).",
    )
    parser.add_argument(
        "--target",
        type=str,
        default=None,
        help="Optional Cargo target triple (e.g. x86_64-unknown-linux-gnu).",
    )
    parser.add_argument(
        "--features",
        type=str,
        default=DEFAULT_FEATURES,
        help=f"Cargo feature set to enable (default: {DEFAULT_FEATURES}).",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=None,
        help="Optional markdown output path (in addition to stdout).",
    )
    args = parser.parse_args()

    repo_root = _repo_root()
    package = _desktop_package_name(repo_root)
    bin_name = DEFAULT_BIN_NAME
    features = args.features
    target = args.target

    try:
        target_dir = _cargo_target_directory(repo_root)
    except Exception as exc:  # noqa: BLE001
        md = f"## Desktop Rust binary size breakdown\n\nError: failed to run `cargo metadata`: `{exc}`\n"
        print(md)
        _append_step_summary(md)
        if args.out:
            args.out.parent.mkdir(parents=True, exist_ok=True)
            args.out.write_text(md, encoding="utf-8")
        return 1

    bin_path = _binary_path(target_dir, bin_name, target)

    build_cmd: list[str] | None = [
        "cargo",
        "build",
        "-p",
        package,
        "--bin",
        bin_name,
        "--features",
        features,
        "--release",
        "--locked",
    ]
    if target:
        build_cmd.extend(["--target", target])

    if not args.no_build:
        try:
            subprocess.run(build_cmd, cwd=repo_root, check=True)
        except subprocess.CalledProcessError:
            md = _render_markdown(
                package=package,
                bin_name=bin_name,
                features=features,
                target=target,
                target_dir=target_dir,
                bin_path=bin_path,
                bin_size_bytes=bin_path.stat().st_size if bin_path.exists() else None,
                build_cmd=build_cmd,
                crates_cmd=None,
                crates_out=None,
                symbols_cmd=None,
                symbols_out=None,
                llvm_size_cmd=None,
                llvm_size_out=None,
                tool_note="Build failed (see logs above). The size breakdown report could not be generated.",
            )
            print(md)
            _append_step_summary(md)
            if args.out:
                args.out.parent.mkdir(parents=True, exist_ok=True)
                args.out.write_text(md, encoding="utf-8")
            return 1

    if not bin_path.exists():
        md = _render_markdown(
            package=package,
            bin_name=bin_name,
            features=features,
            target=target,
            target_dir=target_dir,
            bin_path=bin_path,
            bin_size_bytes=None,
            build_cmd=build_cmd,
            crates_cmd=None,
            crates_out=None,
            symbols_cmd=None,
            symbols_out=None,
            llvm_size_cmd=None,
            llvm_size_out=None,
            tool_note="Binary not found. Hint: run the build command in this report (or omit `--no-build`).",
        )
        print(md)
        _append_step_summary(md)
        if args.out:
            args.out.parent.mkdir(parents=True, exist_ok=True)
            args.out.write_text(md, encoding="utf-8")
        return 1

    bin_size_bytes = bin_path.stat().st_size

    # Preferred: cargo-bloat.
    cargo_bloat_probe = _run_capture(["cargo", "bloat", "--version"], cwd=repo_root)
    has_cargo_bloat = cargo_bloat_probe.returncode == 0

    crates_cmd: list[str] | None = None
    crates_out: CmdResult | None = None
    symbols_cmd: list[str] | None = None
    symbols_out: CmdResult | None = None

    tool_note: str | None = None
    if has_cargo_bloat:
        crates_cmd = [
            "cargo",
            "bloat",
            "-p",
            package,
            "--bin",
            bin_name,
            "--features",
            features,
            "--release",
            "--crates",
            "-n",
            str(args.top),
        ]
        symbols_cmd = [
            "cargo",
            "bloat",
            "-p",
            package,
            "--bin",
            bin_name,
            "--features",
            features,
            "--release",
            "-n",
            str(args.top),
        ]
        if target:
            crates_cmd.extend(["--target", target])
            symbols_cmd.extend(["--target", target])

        crates_out = _run_capture(crates_cmd, cwd=repo_root)
        symbols_out = _run_capture(symbols_cmd, cwd=repo_root)

        # If cargo-bloat failed for some reason, keep going to provide a fallback (llvm-size).
        if crates_out.returncode != 0 or symbols_out.returncode != 0:
            tool_note = (
                "Note: `cargo bloat` exited with a non-zero status. See the output below; "
                "falling back to `llvm-size` where available."
            )
    else:
        tool_note = (
            "Note: `cargo-bloat` is not installed (or `cargo bloat` is unavailable). "
            "Install with `cargo install cargo-bloat --locked` for crate/symbol breakdowns."
        )

    # Optional: add llvm-size output (helpful even when cargo-bloat is present).
    llvm_size_cmd: list[str] | None = None
    llvm_size_out: CmdResult | None = None
    llvm_size = shutil.which("llvm-size") or shutil.which("size")
    if llvm_size:
        # Prefer SysV output when supported (section-by-section breakdown).
        llvm_size_cmd = [llvm_size, "-A", str(bin_path)]
        llvm_size_out = _run_capture(llvm_size_cmd, cwd=repo_root)

        # Some `size` implementations don't support `-A`. Retry without it if needed.
        if llvm_size_out.returncode != 0:
            llvm_size_cmd = [llvm_size, str(bin_path)]
            llvm_size_out = _run_capture(llvm_size_cmd, cwd=repo_root)

    md = _render_markdown(
        package=package,
        bin_name=bin_name,
        features=features,
        target=target,
        target_dir=target_dir,
        bin_path=bin_path,
        bin_size_bytes=bin_size_bytes,
        build_cmd=build_cmd,
        crates_cmd=crates_cmd,
        crates_out=crates_out,
        symbols_cmd=symbols_cmd,
        symbols_out=symbols_out,
        llvm_size_cmd=llvm_size_cmd,
        llvm_size_out=llvm_size_out,
        tool_note=tool_note,
    )

    print(md)
    sys.stdout.flush()
    _append_step_summary(md)

    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(md, encoding="utf-8")

    # Informational by default: only fail on catastrophic errors like build/binary not found.
    if has_cargo_bloat and crates_out and symbols_out:
        if crates_out.returncode != 0 or symbols_out.returncode != 0:
            return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
