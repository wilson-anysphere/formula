#!/usr/bin/env python3
"""
CI guardrail: verify that the built macOS `.app` bundle actually contains the
expected file associations and URL scheme registration.

This inspects the *built* `Contents/Info.plist` (not the repo template) so we
catch packaging regressions where `tauri.conf.json` or `Info.plist` changes do
not propagate into final artifacts.
"""

from __future__ import annotations

import argparse
import json
import plistlib
import sys
from pathlib import Path
from typing import Any, Iterable


def _normalize_ext(ext: str) -> str:
    return ext.strip().lstrip(".").lower()


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


def load_expected_extensions(tauri_config_path: Path) -> set[str]:
    config = json.loads(tauri_config_path.read_text(encoding="utf-8"))
    associations = config.get("bundle", {}).get("fileAssociations", [])
    expected: set[str] = set()
    for assoc in associations:
        if not isinstance(assoc, dict):
            continue
        for ext in assoc.get("ext", []) or []:
            if not isinstance(ext, str):
                continue
            normalized = _normalize_ext(ext)
            if normalized:
                expected.add(normalized)

    if not expected:
        raise SystemExit(
            f"no file association extensions found in {tauri_config_path} under bundle.fileAssociations"
        )
    return expected


def _extensions_from_document_types(document_types: Any) -> set[str]:
    out: set[str] = set()
    if not isinstance(document_types, list):
        return out
    for entry in document_types:
        if not isinstance(entry, dict):
            continue
        raw_exts = entry.get("CFBundleTypeExtensions")
        if isinstance(raw_exts, str):
            normalized = _normalize_ext(raw_exts)
            if normalized:
                out.add(normalized)
        elif isinstance(raw_exts, list):
            for ext in raw_exts:
                if not isinstance(ext, str):
                    continue
                normalized = _normalize_ext(ext)
                if normalized:
                    out.add(normalized)
    return out


def _extensions_from_uti_declarations(plist: dict[str, Any]) -> set[str]:
    out: set[str] = set()
    for key in ("UTExportedTypeDeclarations", "UTImportedTypeDeclarations"):
        decls = plist.get(key)
        if not isinstance(decls, list):
            continue
        for decl in decls:
            if not isinstance(decl, dict):
                continue
            tags = decl.get("UTTypeTagSpecification")
            if not isinstance(tags, dict):
                continue
            raw_exts = tags.get("public.filename-extension")
            if isinstance(raw_exts, str):
                normalized = _normalize_ext(raw_exts)
                if normalized:
                    out.add(normalized)
            elif isinstance(raw_exts, list):
                for ext in raw_exts:
                    if not isinstance(ext, str):
                        continue
                    normalized = _normalize_ext(ext)
                    if normalized:
                        out.add(normalized)
    return out


def extract_registered_extensions(plist: dict[str, Any]) -> set[str]:
    return _extensions_from_document_types(plist.get("CFBundleDocumentTypes")) | _extensions_from_uti_declarations(
        plist
    )


def extract_url_schemes(plist: dict[str, Any]) -> set[str]:
    out: set[str] = set()
    url_types = plist.get("CFBundleURLTypes")
    if not isinstance(url_types, list):
        return out
    for entry in url_types:
        if not isinstance(entry, dict):
            continue
        schemes = entry.get("CFBundleURLSchemes")
        if isinstance(schemes, str):
            scheme = schemes.strip().lower()
            if scheme:
                out.add(scheme)
        elif isinstance(schemes, list):
            for scheme in schemes:
                if not isinstance(scheme, str):
                    continue
                normalized = scheme.strip().lower()
                if normalized:
                    out.add(normalized)
    return out


def _format_set(items: Iterable[str]) -> str:
    return ", ".join(sorted(set(items)))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--info-plist",
        required=True,
        type=Path,
        help="Path to the built .app bundle Info.plist (e.g. Formula.app/Contents/Info.plist)",
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
        help="Expected CFBundleURLSchemes entry (deep link scheme)",
    )
    args = parser.parse_args()

    expected_exts = load_expected_extensions(args.tauri_config)
    expected_schemes = load_expected_deep_link_schemes(args.tauri_config)
    if not expected_schemes:
        expected_scheme = args.url_scheme.strip().lower()
        if not expected_scheme:
            raise SystemExit("--url-scheme must be non-empty")
        expected_schemes = {expected_scheme}

    with args.info_plist.open("rb") as f:
        plist = plistlib.load(f)
    if not isinstance(plist, dict):
        raise SystemExit(f"unexpected plist type in {args.info_plist}: {type(plist)}")

    registered_exts = extract_registered_extensions(plist)
    registered_schemes = extract_url_schemes(plist)

    print(f"[macos] Info.plist: {args.info_plist}")
    print(f"[macos] Expected file extensions ({len(expected_exts)}): {_format_set(expected_exts)}")
    print(
        f"[macos] Observed file extensions ({len(registered_exts)}): {_format_set(registered_exts)}"
    )
    print(f"[macos] Expected URL schemes ({len(expected_schemes)}): {_format_set(expected_schemes)}")
    print(f"[macos] Observed URL schemes ({len(registered_schemes)}): {_format_set(registered_schemes)}")

    missing_exts = expected_exts - registered_exts
    missing_schemes = expected_schemes - registered_schemes

    if missing_exts or missing_schemes:
        print("[macos] ERROR: bundle registration mismatch", file=sys.stderr)
        if missing_exts:
            print(
                f"[macos] Missing extensions: {_format_set(missing_exts)}",
                file=sys.stderr,
            )
        if missing_schemes:
            print(
                f"[macos] Missing URL schemes: {_format_set(missing_schemes)}",
                file=sys.stderr,
            )
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
