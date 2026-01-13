#!/usr/bin/env python3
"""
CI guardrail: verify that Linux packaging (.deb extracted root) includes a
`.desktop` file that advertises the MIME types corresponding to our configured
file associations and that the Exec line accepts a file/URL argument.

We intentionally inspect the *built* `.desktop` file (from an extracted .deb)
instead of only validating `tauri.conf.json` so we catch packaging regressions.
"""

from __future__ import annotations

import argparse
import configparser
import json
import sys
from pathlib import Path
from typing import Iterable


def _format_set(items: Iterable[str]) -> str:
    return ", ".join(sorted(set(items)))


def load_expected_deep_link_schemes(tauri_config_path: Path) -> set[str]:
    config = json.loads(tauri_config_path.read_text(encoding="utf-8"))
    plugins = config.get("plugins")
    if not isinstance(plugins, dict):
        return set()

    deep_link = plugins.get("deep-link")
    if not isinstance(deep_link, dict):
        return set()

    desktop = deep_link.get("desktop")
    schemes: set[str] = set()

    def add_from_protocol(protocol: object) -> None:
        if not isinstance(protocol, dict):
            return
        raw = protocol.get("schemes")
        if isinstance(raw, str):
            val = raw.strip().lower()
            if val:
                schemes.add(val)
        elif isinstance(raw, list):
            for item in raw:
                if not isinstance(item, str):
                    continue
                val = item.strip().lower()
                if val:
                    schemes.add(val)

    if isinstance(desktop, list):
        for protocol in desktop:
            add_from_protocol(protocol)
    else:
        add_from_protocol(desktop)

    return schemes


def load_expected_mime_types(tauri_config_path: Path) -> set[str]:
    config = json.loads(tauri_config_path.read_text(encoding="utf-8"))
    associations = config.get("bundle", {}).get("fileAssociations", [])
    expected: set[str] = set()
    missing_mime_type_exts: list[str] = []
    for assoc in associations:
        if not isinstance(assoc, dict):
            continue
        mime_type = assoc.get("mimeType")
        exts = assoc.get("ext", [])
        if isinstance(mime_type, str):
            mt = mime_type.strip().lower()
            if mt:
                expected.add(mt)
        elif isinstance(mime_type, list):
            for raw in mime_type:
                if isinstance(raw, str) and raw.strip():
                    expected.add(raw.strip().lower())
        else:
            # If any association lacks a MIME type, Linux packaging can't reliably
            # advertise it in `MimeType=` (the `.desktop` file uses MIME types, not
            # extensions). Fail fast with a clear message so the config is fixed
            # instead of papering over it.
            if isinstance(exts, list):
                missing_mime_type_exts.extend([str(e) for e in exts])
            else:
                missing_mime_type_exts.append(str(exts))

    if missing_mime_type_exts:
        raise SystemExit(
            "missing bundle.fileAssociations[].mimeType for extensions: "
            + ", ".join(missing_mime_type_exts)
        )
    if not expected:
        raise SystemExit(
            f"no file association MIME types found in {tauri_config_path} under bundle.fileAssociations"
        )
    return expected


def find_desktop_files(package_root: Path) -> list[Path]:
    candidates: list[Path] = []
    for rel in (
        Path("usr/share/applications"),
        Path("usr/local/share/applications"),
        Path("share/applications"),
    ):
        app_dir = package_root / rel
        if not app_dir.is_dir():
            continue
        candidates.extend(sorted(app_dir.glob("*.desktop")))
    if candidates:
        return candidates
    # Fallback: anything in the extracted root.
    return sorted(package_root.rglob("*.desktop"))


def parse_desktop_entry(path: Path) -> tuple[set[str], str]:
    parser = configparser.ConfigParser(interpolation=None)
    parser.optionxform = str  # keep case
    parser.read(path, encoding="utf-8")
    if "Desktop Entry" not in parser:
        raise SystemExit(f"{path} missing [Desktop Entry] section")
    entry = parser["Desktop Entry"]
    mime_raw = entry.get("MimeType", "").strip()
    # MIME types are case-insensitive by spec; normalize to lowercase so we don't fail on
    # capitalization differences (e.g. macroEnabled vs macroenabled).
    mime_types = {part.strip().lower() for part in mime_raw.split(";") if part.strip()}
    exec_line = entry.get("Exec", "").strip()
    return mime_types, exec_line


def exec_accepts_file_arg(exec_line: str) -> bool:
    return any(token in exec_line for token in ("%u", "%U", "%f", "%F"))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--package-root",
        "--deb-root",  # backwards-compat (older workflow/scripts)
        dest="package_root",
        type=Path,
        required=True,
        help="Path to extracted Linux package root (dpkg-deb -x output directory or rpm2cpio output directory)",
    )
    parser.add_argument(
        "--tauri-config",
        default=Path("apps/desktop/src-tauri/tauri.conf.json"),
        type=Path,
        help="Path to tauri.conf.json (source of truth for expected file associations)",
    )
    parser.add_argument(
        "--url-scheme",
        default="formula",
        help="Expected x-scheme-handler/<scheme> entry in the .desktop MimeType= list",
    )
    args = parser.parse_args()

    expected_mime_types = load_expected_mime_types(args.tauri_config)
    expected_schemes = load_expected_deep_link_schemes(args.tauri_config)
    if not expected_schemes:
        expected_scheme = args.url_scheme.strip().lower()
        if not expected_scheme:
            raise SystemExit("--url-scheme must be non-empty")
        expected_schemes = {expected_scheme}
    expected_scheme_mimes = {f"x-scheme-handler/{scheme}" for scheme in expected_schemes}
    desktop_files = find_desktop_files(args.package_root)
    if not desktop_files:
        print(f"[linux] ERROR: no .desktop files found under {args.package_root}", file=sys.stderr)
        return 1

    observed_mime_types: set[str] = set()
    desktop_entries: list[tuple[Path, set[str], str]] = []

    print(f"[linux] Extracted package root: {args.package_root}")
    print("[linux] Desktop files:")

    for desktop_file in desktop_files:
        mime_types, exec_line = parse_desktop_entry(desktop_file)
        observed_mime_types |= mime_types
        desktop_entries.append((desktop_file, mime_types, exec_line))
        print(f"[linux] - {desktop_file}")
        print(f"[linux]   Exec={exec_line!r}")
        print(f"[linux]   MimeType entries ({len(mime_types)}): {_format_set(mime_types)}")

    print(
        f"[linux] Expected MIME types from tauri.conf.json ({len(expected_mime_types)}): {_format_set(expected_mime_types)}"
    )
    print(f"[linux] Expected deep link scheme MIME types ({len(expected_scheme_mimes)}): {_format_set(expected_scheme_mimes)}")
    print(f"[linux] Observed MIME types from .desktop files ({len(observed_mime_types)}): {_format_set(observed_mime_types)}")

    missing_mime_types = expected_mime_types - observed_mime_types
    if missing_mime_types:
        print(
            f"[linux] ERROR: missing MIME types in .desktop MimeType=: {_format_set(missing_mime_types)}",
            file=sys.stderr,
        )
        return 1

    missing_scheme_mimes = expected_scheme_mimes - observed_mime_types
    if missing_scheme_mimes:
        print(
            f"[linux] ERROR: missing deep link scheme handler(s) in .desktop MimeType=: {_format_set(missing_scheme_mimes)}",
            file=sys.stderr,
        )
        return 1

    file_assoc_entries = [
        (path, exec_line)
        for (path, mime_types, exec_line) in desktop_entries
        if mime_types & expected_mime_types
    ]
    if not file_assoc_entries:
        print(
            "[linux] ERROR: no .desktop files advertise any expected file-association MIME types",
            file=sys.stderr,
        )
        return 1

    bad_file_assoc_exec = [
        f"{path} Exec={exec_line!r}"
        for (path, exec_line) in file_assoc_entries
        if not exec_accepts_file_arg(exec_line)
    ]
    if bad_file_assoc_exec:
        print(
            "[linux] ERROR: .desktop file(s) that advertise file associations must include a file/URL placeholder in Exec= (%u/%U/%f/%F)",
            file=sys.stderr,
        )
        print("[linux] Offending entries:", file=sys.stderr)
        for line in bad_file_assoc_exec:
            print(f"[linux] - {line}", file=sys.stderr)
        return 1

    scheme_entries = [
        (path, exec_line)
        for (path, mime_types, exec_line) in desktop_entries
        if mime_types & expected_scheme_mimes
    ]
    if not scheme_entries:
        print(
            "[linux] ERROR: no .desktop files advertise the expected x-scheme-handler MIME types",
            file=sys.stderr,
        )
        return 1

    bad_scheme_exec = [
        f"{path} Exec={exec_line!r}"
        for (path, exec_line) in scheme_entries
        if not exec_accepts_file_arg(exec_line)
    ]
    if bad_scheme_exec:
        print(
            "[linux] ERROR: .desktop file(s) that advertise x-scheme-handler/* must include a URL placeholder in Exec= (%u/%U/%f/%F)",
            file=sys.stderr,
        )
        print("[linux] Offending entries:", file=sys.stderr)
        for line in bad_scheme_exec:
            print(f"[linux] - {line}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
