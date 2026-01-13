#!/usr/bin/env python3
"""
Verify that the Tauri desktop release binary is stripped (no accidental symbols/debug info).

This is intended for use in `.github/workflows/release.yml` after the Tauri build step so the
workflow fails fast if debug/symbol data sneaks into shipped release artifacts.
"""

from __future__ import annotations

import os
import platform
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


DESKTOP_BINARY_NAME = "formula-desktop"
_SKIP_DIR_SUFFIXES = (".app", ".framework", ".xcframework")


@dataclass(frozen=True)
class CandidateBinary:
    path: Path
    mtime: float


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
    # scripts/verify_desktop_binary_stripped.py -> repo root
    return Path(__file__).resolve().parents[1]


def _candidate_target_dirs(repo_root: Path) -> list[Path]:
    """
    The desktop build output location depends on whether Cargo is treating the repo as a workspace,
    and whether callers override `CARGO_TARGET_DIR`.

    - Workspace default: `<repo>/target`
    - Standalone src-tauri build: `apps/desktop/src-tauri/target`
    - Builds from `apps/desktop`: `apps/desktop/target`
    - Custom: `CARGO_TARGET_DIR=...`
    """

    candidates: list[Path] = []

    env_target = os.environ.get("CARGO_TARGET_DIR", "").strip()
    if env_target:
        p = Path(env_target)
        if not p.is_absolute():
            # Cargo interprets relative paths relative to the working directory used for cargo.
            p = repo_root / p
        candidates.append(p)

    for p in (
        repo_root / "target",
        repo_root / "apps" / "desktop" / "src-tauri" / "target",
        repo_root / "apps" / "desktop" / "target",
    ):
        if p.is_dir():
            candidates.append(p)

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


def _iter_release_binary_candidates(target_dir: Path, exe_name: str) -> Iterable[Path]:
    # Typical: target/release/formula-desktop(.exe)
    # Also: target/<triple>/release/formula-desktop(.exe)
    # Avoid scanning huge dirs; glob is fine here.
    for p in target_dir.glob(f"**/release/{exe_name}"):
        if p.is_file():
            yield p


def _find_desktop_binary(repo_root: Path) -> Path:
    exe = f"{DESKTOP_BINARY_NAME}.exe" if sys.platform == "win32" else DESKTOP_BINARY_NAME

    candidates: list[CandidateBinary] = []
    for target_dir in _candidate_target_dirs(repo_root):
        for p in _iter_release_binary_candidates(target_dir, exe):
            try:
                stat = p.stat()
            except FileNotFoundError:
                continue
            candidates.append(CandidateBinary(path=p, mtime=stat.st_mtime))

    if not candidates:
        target_dirs = ", ".join(str(p) for p in _candidate_target_dirs(repo_root))
        raise SystemExit(
            f"[strip-check] ERROR: could not find {exe!r} under any target directory. "
            f"(searched: {target_dirs})"
        )

    # Prefer newest output (handles multiple target triples / stale builds).
    candidates.sort(key=lambda c: c.mtime, reverse=True)
    return candidates[0].path


def _run(cmd: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(cmd, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, check=False)


def _assert(condition: bool, message: str) -> None:
    if not condition:
        raise SystemExit(f"[strip-check] ERROR: {message}")


def _append_step_summary(markdown: str) -> None:
    """
    Best-effort GitHub Actions step summary integration.

    We keep this optional so the script remains usable locally.
    """

    summary_path = os.environ.get("GITHUB_STEP_SUMMARY", "").strip()
    if not summary_path:
        return
    try:
        Path(summary_path).parent.mkdir(parents=True, exist_ok=True)
        with open(summary_path, "a", encoding="utf-8") as f:
            f.write(markdown.rstrip() + "\n")
    except Exception:
        # Don't fail the verification if step-summary writing fails.
        return


def _check_unix_stripped(binary: Path) -> None:
    # `file` output includes "not stripped" when symbols/debug info remain.
    proc = _run(["file", str(binary)])
    _assert(proc.returncode == 0, f"`file` failed for {binary}: {proc.stdout.strip()}")
    out = proc.stdout.strip()
    _assert("not stripped" not in out, f"binary is not stripped: {out}")


def _check_linux_no_debug_sections(binary: Path) -> None:
    # `readelf -S` is a stable way to detect `.debug_*` sections.
    proc = _run(["readelf", "-S", "--wide", str(binary)])
    _assert(proc.returncode == 0, f"`readelf` failed for {binary}: {proc.stdout.strip()}")
    out = proc.stdout
    # If any DWARF sections are present, the binary isn't stripped.
    has_debug_sections = ".debug_" in out
    _assert(not has_debug_sections, "ELF contains .debug_* sections (expected stripped)")


def _check_macos_no_dwarf_segment(binary: Path) -> None:
    # Mach-O debug info is typically stored in a `__DWARF` segment.
    proc = _run(["otool", "-l", str(binary)])
    _assert(proc.returncode == 0, f"`otool` failed for {binary}: {proc.stdout.strip()}")
    out = proc.stdout
    _assert("__DWARF" not in out, "Mach-O contains __DWARF segment (expected stripped/no debug info)")


def _bundle_dirs(repo_root: Path) -> list[Path]:
    bundle_dirs: list[Path] = []
    for target_dir in _candidate_target_dirs(repo_root):
        for p in target_dir.glob("**/release/bundle"):
            if p.is_dir():
                bundle_dirs.append(p)
    # De-dupe
    seen: set[Path] = set()
    uniq: list[Path] = []
    for p in bundle_dirs:
        key = p.resolve()
        if key in seen:
            continue
        seen.add(key)
        uniq.append(p)
    return uniq


def _check_no_symbol_sidecars(repo_root: Path) -> None:
    """
    Even if the main binary is stripped, some toolchains can generate sidecar debug files.
    Ensure we don't ship them from the bundle output directories.
    """

    bundle_dirs = _bundle_dirs(repo_root)
    if not bundle_dirs:
        # Not all builds produce bundles (e.g. local `cargo build`). Don't fail.
        return

    # `cargo`/linkers can emit separate debug symbol artifacts:
    # - Windows: `.pdb` (often not bundled by default, but we guard against it)
    # - macOS: `.dSYM` directory, sometimes archived as `.dSYM.zip`/`.dSYM.tar.gz`
    # - Linux: `.dwp` (DWARF package)
    #
    # We scan bundle output directories rather than `target/` so local debug artifacts
    # don't fail the check unless they would be shipped to users.
    forbidden_file_suffixes = (".pdb", ".dwp")
    offenders: list[Path] = []
    for bundle_dir in bundle_dirs:
        for root, dirs, files in os.walk(bundle_dir):
            # Record (and avoid descending into) dSYM directories.
            for d in list(dirs):
                if d.lower().endswith(".dsym"):
                    offenders.append(Path(root) / d)
                    dirs.remove(d)

            # Prune large nested bundles/framework trees; dSYM directories are handled above.
            dirs[:] = [d for d in dirs if not d.lower().endswith(_SKIP_DIR_SUFFIXES)]

            for f in files:
                lower = f.lower()
                if lower.endswith(forbidden_file_suffixes):
                    offenders.append(Path(root) / f)
                    continue
                # Catch archived debug bundles (e.g. `Formula.app.dSYM.zip`).
                if lower.endswith(".dsym") or ".dsym." in lower:
                    offenders.append(Path(root) / f)
                    continue
                if ".pdb." in lower:
                    offenders.append(Path(root) / f)

    _assert(
        not offenders,
        "found debug/symbol sidecar files in bundle output:\n"
        + "\n".join(f"  - {p.relative_to(repo_root)}" for p in offenders[:50])
        + ("\n  (truncated)" if len(offenders) > 50 else ""),
    )


def main() -> None:
    repo_root = _repo_root()
    binary = _find_desktop_binary(repo_root)
    runner_os = os.environ.get("RUNNER_OS") or platform.system()
    binary_display = binary.relative_to(repo_root) if binary.is_relative_to(repo_root) else binary
    binary_size = _human_bytes(binary.stat().st_size)

    print(f"[strip-check] Runner: {runner_os}")
    print(f"[strip-check] Checking binary: {binary_display}")

    if sys.platform == "win32":
        # On Windows, debug symbols are primarily shipped via `.pdb` sidecar files. We verify that
        # the bundle output does not contain PDBs. (The executable may still reference a PDB path in
        # its debug directory; that is not shipped to users as long as the PDB itself isn't bundled.)
        _check_no_symbol_sidecars(repo_root)
        print("[strip-check] OK (Windows): no PDB/dSYM sidecars in bundle output")
        _append_step_summary(
            "\n".join(
                [
                    "### Desktop binary strip verification",
                    "",
                    f"- Platform: **Windows**",
                    f"- Binary: `{binary_display}`",
                    f"- Binary size: **{binary_size}**",
                    "- Check: no `.pdb`/`.dSYM`/`.dwp` sidecars in `**/release/bundle/**`",
                    "",
                ]
            )
        )
        return

    _check_unix_stripped(binary)
    if sys.platform == "linux":
        _check_linux_no_debug_sections(binary)
        _check_no_symbol_sidecars(repo_root)
        print("[strip-check] OK (Linux): binary stripped, no .debug_* sections, no symbol sidecars")
        _append_step_summary(
            "\n".join(
                [
                    "### Desktop binary strip verification",
                    "",
                    f"- Platform: **Linux**",
                    f"- Binary: `{binary_display}`",
                    f"- Binary size: **{binary_size}**",
                    "- Checks: `file` (stripped), `readelf -S` (no `.debug_*`), no symbol sidecars in `**/release/bundle/**`",
                    "",
                ]
            )
        )
        return

    if sys.platform == "darwin":
        _check_macos_no_dwarf_segment(binary)
        _check_no_symbol_sidecars(repo_root)
        print("[strip-check] OK (macOS): binary stripped, no __DWARF segment, no symbol sidecars")
        _append_step_summary(
            "\n".join(
                [
                    "### Desktop binary strip verification",
                    "",
                    f"- Platform: **macOS**",
                    f"- Binary: `{binary_display}`",
                    f"- Binary size: **{binary_size}**",
                    "- Checks: `file` (stripped), `otool -l` (no `__DWARF`), no symbol sidecars in `**/release/bundle/**`",
                    "",
                ]
            )
        )
        return

    print(f"[strip-check] WARNING: unsupported platform {sys.platform}; skipping additional checks.")


if __name__ == "__main__":
    main()
