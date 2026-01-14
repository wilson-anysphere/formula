#!/usr/bin/env python3

"""
Release guardrail: ensure Windows installers are Authenticode-signed.

Deprecated: the Desktop Release workflow now uses the shared PowerShell validator:
  scripts/validate-windows-bundles.ps1
which also validates installer presence and is the single source of truth for
Windows installer validation in CI.

When Windows code signing secrets are configured (WINDOWS_CERTIFICATE / password),
the release workflow should produce signed NSIS (.exe) and WiX (.msi) installers.

This script locates the produced installers under a bundle directory and runs:
  signtool verify /pa /all /v <installer>

It intentionally does *not* try to validate the certificate subject/issuer (that
can vary by org), only that the installer passes Windows Authenticode policy
validation.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Iterable


def _find_signtool() -> str | None:
    exe = shutil.which("signtool") or shutil.which("signtool.exe")
    if exe:
        return exe

    if os.name != "nt":
        return None

    roots: list[Path] = []
    for env_var in ("ProgramFiles(x86)", "ProgramFiles", "ProgramW6432"):
        root = os.environ.get(env_var)
        if root:
            roots.append(Path(root) / "Windows Kits" / "10" / "bin")
    # Hardcoded fallback.
    roots.append(Path("C:/Program Files (x86)/Windows Kits/10/bin"))

    candidates: list[Path] = []
    existing_roots: list[Path] = []
    for root in roots:
        if not root.is_dir():
            continue
        existing_roots.append(root)

        # Fast path: the Windows SDK layout is predictable:
        #   .../Windows Kits/10/bin/<sdk-version>/<arch>/signtool.exe
        #   .../Windows Kits/10/bin/<arch>/signtool.exe
        #
        # Avoid a recursive `**/signtool.exe` glob when possible; those directories can be
        # large and recursive scanning slows CI.
        try:
            children = [p for p in root.iterdir() if p.is_dir()]
        except OSError:
            children = []

        # Check versioned directories first (prefer highest version, then x64).
        version_dirs: list[Path] = []
        for child in children:
            name = child.name
            if name and name[0].isdigit() and "." in name:
                version_dirs.append(child)
        for version_dir in version_dirs:
            for arch in ("x64", "x86", "arm64"):
                p = version_dir / arch / "signtool.exe"
                if p.is_file():
                    candidates.append(p)

        # Also handle non-versioned layout: bin/<arch>/signtool.exe
        for arch in ("x64", "x86", "arm64"):
            p = root / arch / "signtool.exe"
            if p.is_file():
                candidates.append(p)

    # Fallback: recursive scan only when the expected layout isn't present.
    if not candidates:
        # Keep the fallback bounded: Windows Kits directories can be large, and a fully recursive
        # glob over `bin/` can be slow in CI. The expected layout is shallow (bin/<version>/<arch>).
        max_depth = 4
        for kit_root in existing_roots:
            for dirpath, dirnames, filenames in os.walk(kit_root):
                cur = Path(dirpath)
                try:
                    depth = len(cur.relative_to(kit_root).parts)
                except ValueError:
                    depth = max_depth
                if depth >= max_depth:
                    dirnames[:] = []
                for filename in filenames:
                    if filename.lower() == "signtool.exe":
                        p = cur / filename
                        if p.is_file():
                            candidates.append(p)

    if not candidates:
        return None

    def version_tuple(p: Path) -> tuple[int, ...]:
        # Typical layout: .../Windows Kits/10/bin/<sdk-version>/<arch>/signtool.exe
        arch_dir = p.parent
        parent = arch_dir.parent
        if parent and parent.name and "." in parent.name and parent.name[0].isdigit():
            try:
                return tuple(int(x) for x in parent.name.split("."))
            except Exception:
                return (0,)
        return (0,)

    def sort_key(p: Path) -> tuple[bool, tuple[int, ...], str]:
        parts_lc = [part.lower() for part in p.parts]
        is_x64 = "x64" in parts_lc
        return (is_x64, version_tuple(p), str(p).lower())

    # Prefer x64 + highest SDK version.
    best = max(candidates, key=sort_key)
    return str(best)


def _find_installers(bundle_dir: Path) -> tuple[list[Path], list[Path]]:
    # Tauri installer outputs are emitted directly under bundle/<kind>/ (no deep nesting).
    # Avoid recursive `**/*.msi` globbing even though bundle dirs are usually small.
    msis = sorted([p for p in (bundle_dir / "msi").glob("*.msi") if p.is_file()])

    exes: list[Path] = []
    for sub in ("nsis", "nsis-web"):
        d = bundle_dir / sub
        if d.is_dir():
            exes.extend([p for p in d.glob("*.exe") if p.is_file()])

    # Exclude embedded WebView2 helper installers; we only care about the Formula installer(s).
    filtered_exes: list[Path] = []
    for exe in sorted(exes):
        base = exe.name.lower()
        if base.startswith("microsoftedgewebview2"):
            continue
        filtered_exes.append(exe)

    return (msis, filtered_exes)


def _run_signtool(signtool: str, installer: Path) -> tuple[int, str]:
    cmd = [signtool, "verify", "/pa", "/all", "/v", str(installer)]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=300)
    out = proc.stdout.decode("utf-8", errors="replace")
    return (proc.returncode, out)

def _assert_timestamped(signtool_output: str) -> str | None:
    """
    Return an error string if the signature appears to be missing a timestamp.

    We rely on timestamping so signatures remain valid after the signing cert expires.
    """

    out_lc = signtool_output.lower()
    # Typical signtool output includes either:
    # - "The signature is timestamped."
    # - "The signature is not timestamped."
    if "not timestamped" in out_lc:
        return "signature is not timestamped"
    if "signature is timestamped" in out_lc or "the signature is timestamped" in out_lc:
        return None
    if "timestamp verified by" in out_lc:
        # Some output formats include a timestamp verifier section even if the short sentence is missing.
        return None
    return "unable to determine timestamp status (expected signtool to report a timestamped signature)"


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Verify Authenticode signatures on Windows installers.")
    parser.add_argument(
        "--bundle-dir",
        required=False,
        help="Path to a Tauri bundle output directory (…/release/bundle). Defaults to apps/desktop/src-tauri/target/release/bundle or apps/desktop/src-tauri/target/<triple>/release/bundle discovery.",
    )
    args = parser.parse_args(argv)

    # Resolve the repository root relative to this script location so callers can invoke the
    # verifier from arbitrary working directories.
    repo_root = Path(__file__).resolve().parents[2]

    bundle_dir: Path | None = None
    if args.bundle_dir:
        bundle_dir = Path(args.bundle_dir)
        if not bundle_dir.is_dir():
            print(f"sigcheck: ERROR bundle dir not found: {bundle_dir}", file=sys.stderr)
            return 2
    else:
        # Best-effort fallback: look for a single bundle directory.
        candidates: list[Path] = []
        for root in (
            repo_root / "apps" / "desktop" / "src-tauri" / "target",
            repo_root / "target",
        ):
            if not root.is_dir():
                continue
            # Avoid recursive `**/release/bundle` globbing: Cargo target directories can be large,
            # and we only expect bundles at:
            # - <target>/release/bundle
            # - <target>/<triple>/release/bundle
            for pattern in ("release/bundle", "*/release/bundle"):
                candidates.extend([p for p in root.glob(pattern) if p.is_dir()])
        bundle_dir = candidates[0] if candidates else None
        if bundle_dir is None:
            print(
                "sigcheck: ERROR no bundle dir found. Pass --bundle-dir <path>.\n"
                "Expected something like: apps/desktop/src-tauri/target/<triple>/release/bundle",
                file=sys.stderr,
            )
            return 2

    signtool = _find_signtool()
    if not signtool:
        print(
            "sigcheck: ERROR signtool.exe not found. Ensure Windows SDK is installed on the runner.",
            file=sys.stderr,
        )
        return 2

    msis, exes = _find_installers(bundle_dir)
    if not msis and not exes:
        print(f"sigcheck: ERROR no .msi or installer .exe files found under: {bundle_dir}", file=sys.stderr)
        return 1

    failures: list[str] = []

    def check_paths(paths: Iterable[Path]) -> None:
        nonlocal failures
        for p in paths:
            rel = p
            try:
                rel = p.relative_to(repo_root)
            except ValueError:
                pass
            code, out = _run_signtool(signtool, p)
            if code != 0:
                failures.append(f"{rel} (signtool exit code {code})\n{out}")
                continue
            ts_err = _assert_timestamped(out)
            if ts_err is not None:
                failures.append(f"{rel} ({ts_err})\n{out}")
                continue
            print(f"sigcheck: OK {rel}")

    print(f"sigcheck: using signtool={signtool}")
    check_paths(msis)
    check_paths(exes)

    if failures:
        # Common failure mode: signing succeeds but timestamping fails (network/proxy issues or a bad timestamp URL),
        # producing artifacts that are signed-but-not-timestamped. Timestamping matters because it keeps signatures valid
        # after the signing certificate expires.
        tauri_conf_raw = os.environ.get("FORMULA_TAURI_CONF_PATH", "").strip() or "apps/desktop/src-tauri/tauri.conf.json"
        tauri_conf_display = tauri_conf_raw
        try:
            p = Path(tauri_conf_raw)
            if not p.is_absolute():
                tauri_conf_display = str((repo_root / p).relative_to(repo_root))
        except Exception:
            pass
        print("sigcheck: ERROR one or more installers failed Authenticode verification:", file=sys.stderr)
        print(
            "sigcheck: HINT If this is a timestamping failure, check "
            + f"{tauri_conf_display} -> bundle.windows.timestampUrl "
            + "(must be a reachable https:// timestamp server) and re-run the build.",
            file=sys.stderr,
        )
        for msg in failures:
            print("\n---\n", file=sys.stderr)
            print(msg, file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
