#!/usr/bin/env python3

"""
Release guardrail: ensure Windows installers make WebView2 available.

Formula's Windows build uses the Microsoft Edge WebView2 runtime. On machines where WebView2 is
not already installed, the installer must ensure the runtime is available (either by bundling a
bootstrapper/offline installer, or by shipping a fixed runtime in the app bundle).

This script inspects the *produced* Windows installers under the standard Cargo/Tauri bundle
output directories:
  - `<target>/release/bundle/**`
  - `<target>/<target-triple>/release/bundle/**`
(across common target roots like `apps/desktop/src-tauri/target`, `apps/desktop/target`, `target`,
and any configured `CARGO_TARGET_DIR` override),
and asserts that the installers contain a reference to:
  - the WebView2 Evergreen bootstrapper (`MicrosoftEdgeWebView2Setup.exe`), or
  - an offline runtime installer, or
  - a fixed runtime payload (detected via well-known runtime file names, e.g. `msedgewebview2.exe`).

We intentionally validate the built artifacts (not just tauri.conf.json) so CI fails if the
bundler behavior ever regresses.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Iterable


WEBVIEW2_MARKER_STRINGS = [
    # Evergreen bootstrapper (most common, small).
    "MicrosoftEdgeWebview2Setup.exe",
    "MicrosoftEdgeWebView2Setup.exe",
    # Bootstrapper download URLs (used by downloadBootstrapper mode; the exe may not be embedded).
    # We keep these as substring markers so we match multiple URL variants used by different tools.
    "go.microsoft.com/fwlink",
    "LinkId=2124703",
    "linkid=2124703",
    # Evergreen standalone/offline installers (less common, much larger).
    "MicrosoftEdgeWebView2RuntimeInstaller",
    # Fixed runtime bundle (when using webviewInstallMode.type = "fixedRuntime").
    "msedgewebview2.exe",
]


def _resolve_tauri_config_path(repo_root: Path) -> Path:
    override = os.environ.get("FORMULA_TAURI_CONF_PATH", "").strip()
    if override:
        p = Path(override)
        if not p.is_absolute():
            p = repo_root / p
        return p
    return repo_root / "apps" / "desktop" / "src-tauri" / "tauri.conf.json"


def _load_configured_webview_install_mode(conf_path: Path) -> str | None:
    """
    Read `bundle.windows.webviewInstallMode` from apps/desktop/src-tauri/tauri.conf.json.

    The config can be either a string (shorthand) or an object with a `type` field.
    """

    if not conf_path.is_file():
        return None
    try:
        config = json.loads(conf_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Invalid JSON in {conf_path}: {exc}") from exc

    windows = (config.get("bundle") or {}).get("windows") or {}
    mode = windows.get("webviewInstallMode")
    if isinstance(mode, str):
        return mode
    if isinstance(mode, dict):
        t = mode.get("type")
        return t if isinstance(t, str) else None
    return None


def _find_src_tauri_dirs(repo_root: Path) -> Iterable[Path]:
    """
    Best-effort discovery for Tauri projects without scanning huge directories.
    Mirrors the skip list used by scripts/desktop_bundle_size_report.py.
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
    # Keep the fallback walk bounded: Tauri projects should be shallow (e.g.
    # apps/desktop/src-tauri/tauri.conf.json). This avoids traversing arbitrarily deep trees
    # when the repo contains extracted artifacts or build output.
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


def _candidate_target_dirs(repo_root: Path) -> list[Path]:
    candidates: list[Path] = []

    # Some Tauri build flows export CARGO_TARGET_DIR (and/or use a non-default target dir).
    # Prefer it when available so this verifier can find bundles regardless of build layout.
    cargo_target_dir = os.environ.get("CARGO_TARGET_DIR", "").strip()
    if cargo_target_dir:
        p = Path(cargo_target_dir)
        if not p.is_absolute():
            p = repo_root / p
        if p.is_dir():
            candidates.append(p)

    for p in (
        repo_root / "apps" / "desktop" / "src-tauri" / "target",
        repo_root / "apps" / "desktop" / "target",
        repo_root / "target",
    ):
        if p.is_dir():
            candidates.append(p)

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


def _find_windows_installers(target_dir: Path) -> list[Path]:
    """
    Collect Windows installers produced by Tauri (NSIS .exe and MSI .msi).

    Paths look like:
      - target/release/bundle/nsis/*.exe
      - target/release/bundle/nsis-web/*.exe
      - target/release/bundle/msi/*.msi
      - target/<triple>/release/bundle/nsis/*.exe
    """

    installers: list[Path] = []
    # Avoid recursive `**/release/bundle/**` globs here: Cargo target directories can be very large
    # (especially on CI runners with caching enabled), and recursive globbing can be surprisingly slow.
    #
    # Tauri bundle outputs are always emitted at:
    # - <target_dir>/release/bundle/...
    # - <target_dir>/<triple>/release/bundle/...
    patterns = [
        "release/bundle/nsis/*.exe",
        "*/release/bundle/nsis/*.exe",
        "release/bundle/nsis-web/*.exe",
        "*/release/bundle/nsis-web/*.exe",
        "release/bundle/msi/*.msi",
        "*/release/bundle/msi/*.msi",
    ]
    for pattern in patterns:
        installers.extend([p for p in target_dir.glob(pattern) if p.is_file()])

    # Tauri may place helper executables (e.g. the WebView2 bootstrapper itself) alongside
    # the main installer outputs. We only want to validate the *installers* here.
    def is_installer_candidate(p: Path) -> bool:
        name = p.name.lower()
        if name.endswith(".msi"):
            return True
        if not name.endswith(".exe"):
            return False
        if "microsoftedgewebview2setup.exe" in name:
            return False
        if "microsoftedgewebview2runtimeinstaller" in name:
            return False
        return True

    installers = [p for p in installers if is_installer_candidate(p)]

    # De-dupe.
    seen: set[Path] = set()
    uniq: list[Path] = []
    for i in installers:
        key = i.resolve()
        if key in seen:
            continue
        seen.add(key)
        uniq.append(i)

    return sorted(uniq)


def _binary_contains_any(path: Path, markers: list[bytes]) -> bytes | None:
    """
    Streaming substring search over a binary file.
    Returns the marker found, or None.
    """

    if not markers:
        return None
    max_len = max(len(m) for m in markers)
    overlap = max_len - 1

    carry = b""
    with path.open("rb") as f:
        while True:
            chunk = f.read(1024 * 1024)  # 1 MiB
            if not chunk:
                break
            data = carry + chunk
            for marker in markers:
                if marker in data:
                    return marker
            carry = data[-overlap:] if overlap > 0 and len(data) >= overlap else data

    return None


def _find_7z() -> str | None:
    exe = shutil.which("7z") or shutil.which("7z.exe")
    if exe:
        return exe

    # GitHub's Windows runners typically have 7-Zip installed, but it may not always be on PATH.
    # Look in common install locations so the verifier remains robust.
    if os.name != "nt":
        return None

    candidate_paths: list[Path] = []
    for env_var in ("ProgramW6432", "ProgramFiles", "ProgramFiles(x86)"):
        root = os.environ.get(env_var)
        if root:
            candidate_paths.append(Path(root) / "7-Zip" / "7z.exe")

    # Hardcoded fallbacks in case the env vars aren't set.
    candidate_paths.extend(
        [
            Path("C:/Program Files/7-Zip/7z.exe"),
            Path("C:/Program Files (x86)/7-Zip/7z.exe"),
        ]
    )

    for candidate in candidate_paths:
        if candidate.is_file():
            return str(candidate)

    return None


def _7z_list(archive: Path) -> str | None:
    """
    Return `7z l` output if 7-Zip is available and can list the archive.
    """

    seven_zip = _find_7z()
    if not seven_zip:
        return None

    # -sccUTF-8 forces UTF-8 console encoding (important on Windows).
    cmd = [seven_zip, "l", "-sccUTF-8", "-ba", "-bd", str(archive)]
    try:
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=300)
    except subprocess.TimeoutExpired:
        return None
    if proc.returncode != 0:
        return None
    return proc.stdout.decode("utf-8", errors="replace")


def _7z_extract_and_find_marker(archive: Path, markers: list[str]) -> str | None:
    """
    Extract `archive` with 7-Zip and look for a file whose name contains a marker substring.
    """

    seven_zip = _find_7z()
    if not seven_zip:
        return None

    with tempfile.TemporaryDirectory(prefix="webview2-check-") as tmpdir:
        cmd = [seven_zip, "x", "-y", "-sccUTF-8", f"-o{tmpdir}", str(archive)]
        try:
            proc = subprocess.run(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=300
            )
        except subprocess.TimeoutExpired:
            return None
        if proc.returncode != 0:
            return None

        marker_lc = [m.lower() for m in markers]
        for root, _, files in os.walk(tmpdir):
            for f in files:
                name_lc = f.lower()
                # Avoid `zip(..., strict=...)` so this script can run on older Python versions too.
                for m_lc, m in zip(marker_lc, markers):
                    if m_lc in name_lc:
                        # Return the canonical marker string (not the filename we found).
                        return m
        return None


def _detect_webview2_marker(installer: Path) -> str | None:
    """
    Try multiple strategies to detect WebView2 installer wiring in a built installer.

    Prefer archive listing/extraction (strong signal); fall back to binary substring search.
    """

    # 1) 7z list (fast, no extraction).
    listing = _7z_list(installer)
    if listing is not None:
        listing_lc = listing.lower()
        for marker in WEBVIEW2_MARKER_STRINGS:
            if marker.lower() in listing_lc:
                return marker

    # 2) 7z extract and search filenames (more expensive, but still robust).
    extracted_marker = _7z_extract_and_find_marker(installer, WEBVIEW2_MARKER_STRINGS)
    if extracted_marker is not None:
        return extracted_marker

    # 3) Fallback: search the raw installer binary for either UTF-8/ASCII or UTF-16LE string literals.
    patterns: list[bytes] = []
    for marker in WEBVIEW2_MARKER_STRINGS:
        patterns.append(marker.encode("utf-8"))
        patterns.append(marker.encode("utf-16le"))

    found = _binary_contains_any(installer, patterns)
    if found is None:
        return None

    # Prefer returning the canonical marker string when we can identify it.
    for marker in WEBVIEW2_MARKER_STRINGS:
        if found == marker.encode("utf-8") or found == marker.encode("utf-16le"):
            return marker

    # Fallback: decode best-effort for debugging output.
    if b"\x00" in found:
        return found.decode("utf-16le", errors="replace")
    return found.decode("utf-8", errors="replace")


def main() -> int:
    # Resolve the repository root relative to this script location so callers can invoke the
    # verifier from arbitrary working directories.
    repo_root = Path(__file__).resolve().parents[2]

    tauri_conf_path = _resolve_tauri_config_path(repo_root)
    tauri_conf_display = tauri_conf_path
    try:
        tauri_conf_display = tauri_conf_path.relative_to(repo_root)
    except ValueError:
        pass

    configured_mode = _load_configured_webview_install_mode(tauri_conf_path)
    if configured_mode is None:
        print(
            f"webview2-check: ERROR bundle.windows.webviewInstallMode is not set in {tauri_conf_display}.\n"
            "The Windows installer must be configured to install WebView2 when it is missing (do not rely on users having it preinstalled).",
            file=sys.stderr,
        )
        return 2
    if configured_mode.strip().lower() == "skip":
        print(
            "webview2-check: ERROR bundle.windows.webviewInstallMode is set to 'skip'.\n"
            "This will produce installers that require users to manually install the WebView2 Runtime.",
            file=sys.stderr,
        )
        return 2

    seven_zip = _find_7z()
    if seven_zip:
        print(f"webview2-check: using 7z={seven_zip}")
    else:
        print(
            "webview2-check: 7z not found; falling back to binary marker scanning (less reliable)."
        )

    target_dirs = _candidate_target_dirs(repo_root)
    if not target_dirs:
        print("webview2-check: ERROR No Cargo target directories found.", file=sys.stderr)
        return 2

    installers: list[Path] = []
    for t in target_dirs:
        installers.extend(_find_windows_installers(t))

    if not installers:
        expected = repo_root / "apps" / "desktop" / "src-tauri" / "target" / "release" / "bundle" / "nsis"
        print(
            "webview2-check: ERROR No Windows installers found to inspect.\n"
            f"Searched under: {', '.join(str(t) for t in target_dirs)}\n"
            f"Expected something like: {expected}",
            file=sys.stderr,
        )
        return 1

    failures: list[str] = []
    for installer in installers:
        marker = _detect_webview2_marker(installer)
        rel = installer
        try:
            rel = installer.relative_to(repo_root)
        except ValueError:
            pass

        if marker is None:
            failures.append(f"- {rel} (no WebView2 installer markers found)")
            continue

        print(f"webview2-check: OK {rel} (found {marker})")

    if failures:
        print(
            "webview2-check: ERROR One or more Windows installers do not appear to bundle/reference the WebView2 runtime.\n"
            "Expected to find one of these markers:\n"
            + "\n".join(f"- {m}" for m in WEBVIEW2_MARKER_STRINGS)
            + "\n\nOffenders:\n"
            + "\n".join(failures),
            file=sys.stderr,
        )
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
