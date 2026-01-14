#!/usr/bin/env python3
"""
CI guardrail: verify that Linux packaging includes `.desktop` metadata that
advertises our configured file associations and deep-link scheme handler.

This script operates on an extracted package/AppDir root:
- `.deb`: output of `dpkg-deb -x`
- `.rpm`: output of `rpm2cpio ... | cpio -idm`
- `.AppImage`: extracted `squashfs-root` (from `--appimage-extract`)

We intentionally inspect the *built* `.desktop` files rather than only validating
`tauri.conf.json` so we catch packaging regressions.
"""

from __future__ import annotations

import argparse
import configparser
import json
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Iterable


def _format_set(items: Iterable[str]) -> str:
    return ", ".join(sorted(set(items)))


def _normalize_scheme(value: str) -> str:
    """
    Normalize a deep-link scheme string.

    We accept common user-provided inputs like "formula://", "formula:", or "formula/" and normalize
    them to just "formula" so validations match how OS scheme handlers are actually registered
    (x-scheme-handler/<scheme>).
    """

    v = value.strip().lower()
    # Remove trailing "://", ":" or "/" segments.
    v = re.sub(r"[:/]+$", "", v)
    return v


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
            val = _normalize_scheme(raw)
            if val:
                schemes.add(val)
        elif isinstance(raw, list):
            for item in raw:
                if not isinstance(item, str):
                    continue
                val = _normalize_scheme(item)
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


def load_expected_doc_package_name(tauri_config_path: Path) -> str:
    """
    The Linux bundle validators install compliance artifacts under:
      /usr/share/doc/<package>/{LICENSE,NOTICE}

    For our Tauri builds we use `mainBinaryName` as the stable package/doc directory name
    (it is also the installed binary name under /usr/bin).
    """
    try:
        config = json.loads(tauri_config_path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return "formula-desktop"
    raw = config.get("mainBinaryName")
    if isinstance(raw, str) and raw.strip():
        return raw.strip()
    return "formula-desktop"


def load_expected_identifier(tauri_config_path: Path) -> str:
    """
    Tauri app identifier (reverse-DNS). We use it as the shared-mime-info XML filename:
      /usr/share/mime/packages/<identifier>.xml
    """
    try:
        config = json.loads(tauri_config_path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return ""
    raw = config.get("identifier")
    if isinstance(raw, str) and raw.strip():
        return raw.strip()
    return ""


def verify_compliance_artifacts(package_root: Path, package_name: str) -> None:
    doc_dir = package_root / "usr" / "share" / "doc" / package_name
    missing: list[Path] = []
    for filename in ("LICENSE", "NOTICE"):
        p = doc_dir / filename
        if not p.is_file():
            missing.append(p)
    if missing:
        formatted = "\n".join(f"- {p}" for p in missing)
        raise SystemExit(
            "[linux] ERROR: missing LICENSE/NOTICE compliance artifacts in package doc directory.\n"
            f"Expected under: {doc_dir}\n"
            f"Missing:\n{formatted}\n"
            "Hint: ensure apps/desktop/src-tauri/tauri.conf.json includes bundle.resources and "
            "bundle.linux.(deb|rpm).files mappings for /usr/share/doc/<package>/."
        )


def verify_parquet_mime_definition(package_root: Path, identifier: str) -> None:
    """
    Parquet is not consistently defined in shared-mime-info across distros.

    If we advertise a Parquet MIME type in `MimeType=` we should also ship a
    shared-mime-info definition so `*.parquet` resolves to that MIME type after
    install (via update-mime-database triggers).
    """

    mime_packages_dir = package_root / "usr" / "share" / "mime" / "packages"
    if not mime_packages_dir.is_dir():
        raise SystemExit(
            "[linux] ERROR: Parquet file association configured but no shared-mime-info packages dir found.\n"
            f"Expected: {mime_packages_dir}\n"
            "Hint: install a MIME definition under usr/share/mime/packages (e.g. apps/desktop/src-tauri/mime/...) "
            "and map it into Linux packages via bundle.linux.(deb|rpm|appimage).files in tauri.conf.json."
        )

    expected_mime = "application/vnd.apache.parquet"
    expected_glob = "*.parquet"
    identifier = identifier.strip()

    if not identifier:
        raise SystemExit(
            "[linux] ERROR: Parquet file association configured but tauri.conf.json identifier is missing.\n"
            "Hint: `identifier` is required to determine the expected shared-mime-info XML filename "
            "(/usr/share/mime/packages/<identifier>.xml)."
        )
    if "/" in identifier or "\\" in identifier:
        raise SystemExit(
            "[linux] ERROR: Parquet file association configured but tauri.conf.json identifier is not a valid filename "
            "(contains path separators).\n"
            f"identifier={identifier!r}\n"
            "Hint: `identifier` is used as the shared-mime-info XML filename "
            "(/usr/share/mime/packages/<identifier>.xml). It must not contain '/' or '\\\\'."
        )

    expected_xml = mime_packages_dir / f"{identifier}.xml"
    if not expected_xml.is_file():
        packaged = sorted(mime_packages_dir.glob("*.xml"))
        formatted = "\n".join(f"- {p}" for p in packaged) if packaged else "(no *.xml files found)"
        raise SystemExit(
            "[linux] ERROR: Parquet file association configured but expected shared-mime-info definition file is missing.\n"
            f"Expected: {expected_xml}\n"
            f"Found:\n{formatted}\n"
            "Hint: keep apps/desktop/src-tauri/mime/<identifier>.xml packaged via tauri.conf.json bundle.linux.*.files "
            "(where <identifier> comes from tauri.conf.json identifier)."
        )

    candidates = [expected_xml]
    for xml_path in candidates:
        try:
            tree = ET.parse(xml_path)
        except ET.ParseError as e:
            raise SystemExit(
                "[linux] ERROR: failed to parse expected Parquet shared-mime-info definition XML.\n"
                f"File: {xml_path}\n"
                f"Error: {e}"
            )
        root = tree.getroot()
        for mime_type in root.findall(".//{http://www.freedesktop.org/standards/shared-mime-info}mime-type"):
            if mime_type.get("type") != expected_mime:
                continue
            for glob in mime_type.findall("{http://www.freedesktop.org/standards/shared-mime-info}glob"):
                if glob.get("pattern") == expected_glob:
                    print(
                        f"[linux] Parquet MIME definition OK: {xml_path} defines {expected_mime} ({expected_glob})"
                    )
                    return

    raise SystemExit(
        "[linux] ERROR: Parquet file association configured but the expected shared-mime-info definition file is missing required content.\n"
        f"Expected: {mime_packages_dir}/{identifier}.xml\n"
        f"Expected to define:\n  - {expected_mime} with glob {expected_glob}\n"
        "Hint: ensure the packaged shared-mime-info XML includes a <glob pattern=\"*.parquet\" /> entry."
    )


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


def exec_targets_expected_binary(exec_line: str, expected_binary: str) -> bool:
    """
    Ensure we're validating *our* desktop integration entry, not some unrelated .desktop file.

    Linux packages can technically contain multiple .desktop files. Our file association checks
    should apply to the entry that actually launches Formula.

    We match either:
      - <expected_binary> (from tauri.conf.json mainBinaryName), or
      - AppRun (common for extracted AppImage AppDirs)
    as a token in the Exec= line (allowing an optional path prefix).
    """

    exec_line = exec_line.strip()
    if not exec_line:
        return False

    expected_binary = expected_binary.strip()
    if not expected_binary:
        expected_binary = "formula-desktop"

    token = re.escape(expected_binary)
    pattern = rf'(^|\s)["\']?(?:[^\s"\']*/)?(AppRun|{token})["\']?(\s|$)'
    return re.search(pattern, exec_line, flags=re.IGNORECASE) is not None


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
        "--expected-main-binary",
        default="",
        help=(
            "Override the expected Exec binary token used to select the app's .desktop entry "
            "(defaults to tauri.conf.json mainBinaryName). This does not affect the expected "
            "doc directory name."
        ),
    )
    parser.add_argument(
        "--doc-package-name",
        default="",
        help=(
            "Override the expected /usr/share/doc/<package> directory name used to validate "
            "LICENSE/NOTICE compliance artifacts (defaults to tauri.conf.json mainBinaryName)."
        ),
    )
    parser.add_argument(
        "--url-scheme",
        default="formula",
        help="Expected x-scheme-handler/<scheme> entry in the .desktop MimeType= list",
    )
    args = parser.parse_args()

    expected_mime_types = load_expected_mime_types(args.tauri_config)
    expected_schemes = load_expected_deep_link_schemes(args.tauri_config)
    default_name = load_expected_doc_package_name(args.tauri_config)
    expected_identifier = load_expected_identifier(args.tauri_config)
    expected_doc_pkg = args.doc_package_name.strip() or default_name
    expected_main_binary = args.expected_main_binary.strip() or default_name
    if not expected_schemes:
        expected_scheme = _normalize_scheme(args.url_scheme)
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
    print(f"[linux] Expected doc package name: {expected_doc_pkg!r} (for /usr/share/doc/<package>/)")
    print(f"[linux] Expected main Exec binary: {expected_main_binary!r} (or 'AppRun')")
    if expected_identifier:
        print(f"[linux] Expected Tauri identifier: {expected_identifier!r} (for /usr/share/mime/packages/<identifier>.xml)")
    print(f"[linux] Observed MIME types from all .desktop files ({len(observed_mime_types)}): {_format_set(observed_mime_types)}")

    app_entries = [
        (path, mime_types, exec_line)
        for (path, mime_types, exec_line) in desktop_entries
        if exec_targets_expected_binary(exec_line, expected_main_binary)
    ]
    if not app_entries:
        print(
            f"[linux] ERROR: no .desktop files appear to target the expected executable '{expected_main_binary}' (or AppRun).",
            file=sys.stderr,
        )
        print("[linux] .desktop Exec entries inspected:", file=sys.stderr)
        for (path, _mime_types, exec_line) in desktop_entries:
            print(f"[linux] - {path} Exec={exec_line!r}", file=sys.stderr)
        return 1

    observed_mime_types_app: set[str] = set()
    for _path, mime_types, _exec in app_entries:
        observed_mime_types_app |= mime_types
    print(
        f"[linux] Desktop entries targeting app ({len(app_entries)}): "
        + ", ".join(str(p) for (p, _m, _e) in app_entries)
    )
    print(
        f"[linux] Observed MIME types from app .desktop entries ({len(observed_mime_types_app)}): {_format_set(observed_mime_types_app)}"
    )

    missing_mime_types = expected_mime_types - observed_mime_types_app
    if missing_mime_types:
        print(
            f"[linux] ERROR: missing MIME types in .desktop MimeType=: {_format_set(missing_mime_types)}",
            file=sys.stderr,
        )
        return 1

    missing_scheme_mimes = expected_scheme_mimes - observed_mime_types_app
    if missing_scheme_mimes:
        print(
            f"[linux] ERROR: missing deep link scheme handler(s) in .desktop MimeType=: {_format_set(missing_scheme_mimes)}",
            file=sys.stderr,
        )
        return 1

    file_assoc_entries = [
        (path, exec_line)
        for (path, mime_types, exec_line) in app_entries
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
        for (path, mime_types, exec_line) in app_entries
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

    # Finally, validate that the built package ships OSS/compliance artifacts.
    verify_compliance_artifacts(args.package_root, expected_doc_pkg)

    # Parquet file association requires a shared-mime-info definition on many distros.
    if "application/vnd.apache.parquet" in expected_mime_types:
        verify_parquet_mime_definition(args.package_root, expected_identifier)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
