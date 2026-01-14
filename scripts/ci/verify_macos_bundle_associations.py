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
import os
import plistlib
import re
import sys
from pathlib import Path
from typing import Any, Iterable


def _normalize_ext(ext: str) -> str:
    return ext.strip().lstrip(".").lower()


def _normalize_scheme(value: str) -> str:
    """
    Normalize a deep-link scheme string.

    We accept common user-provided inputs like "formula://", "formula:", or "formula/" and normalize
    them to just "formula" so validations match how schemes are registered in Info.plist
    (CFBundleURLSchemes expects only the scheme name).
    """

    v = value.strip().lower()
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


def load_expected_extensions(tauri_config_path: Path) -> set[str]:
    config = json.loads(tauri_config_path.read_text(encoding="utf-8"))
    associations = config.get("bundle", {}).get("fileAssociations", [])
    expected: set[str] = set()
    for assoc in associations:
        if not isinstance(assoc, dict):
            continue
        raw = assoc.get("ext")
        exts: list[str] = []
        if isinstance(raw, str):
            exts = [raw]
        elif isinstance(raw, list):
            exts = [item for item in raw if isinstance(item, str)]
        for ext in exts:
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


def _utis_from_document_types(document_types: Any) -> set[str]:
    out: set[str] = set()
    if not isinstance(document_types, list):
        return out
    for entry in document_types:
        if not isinstance(entry, dict):
            continue
        raw_utis = entry.get("LSItemContentTypes")
        if isinstance(raw_utis, str):
            uti = raw_utis.strip().lower()
            if uti:
                out.add(uti)
        elif isinstance(raw_utis, list):
            for item in raw_utis:
                if not isinstance(item, str):
                    continue
                uti = item.strip().lower()
                if uti:
                    out.add(uti)
    return out


def _uti_extension_map(plist: dict[str, Any]) -> dict[str, set[str]]:
    out: dict[str, set[str]] = {}
    for key in ("UTExportedTypeDeclarations", "UTImportedTypeDeclarations"):
        decls = plist.get(key)
        if not isinstance(decls, list):
            continue
        for decl in decls:
            if not isinstance(decl, dict):
                continue
            uti_raw = decl.get("UTTypeIdentifier")
            if not isinstance(uti_raw, str):
                continue
            uti = uti_raw.strip().lower()
            if not uti:
                continue
            tags = decl.get("UTTypeTagSpecification")
            if not isinstance(tags, dict):
                continue
            raw_exts = tags.get("public.filename-extension")
            if isinstance(raw_exts, str):
                normalized = _normalize_ext(raw_exts)
                if normalized:
                    out.setdefault(uti, set()).add(normalized)
            elif isinstance(raw_exts, list):
                for ext in raw_exts:
                    if not isinstance(ext, str):
                        continue
                    normalized = _normalize_ext(ext)
                    if normalized:
                        out.setdefault(uti, set()).add(normalized)
    return out


def extract_registered_extensions(plist: dict[str, Any]) -> set[str]:
    doc_types = plist.get("CFBundleDocumentTypes")
    doc_exts = _extensions_from_document_types(doc_types)
    doc_utis = _utis_from_document_types(doc_types)
    uti_map = _uti_extension_map(plist)
    ext_via_utis: set[str] = set()
    for uti in doc_utis:
        ext_via_utis |= uti_map.get(uti, set())
    return doc_exts | ext_via_utis


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

    repo_root = Path(__file__).resolve().parents[2]
    override = os.environ.get("FORMULA_TAURI_CONF_PATH", "").strip()
    if override:
        tauri_default = Path(override)
        if not tauri_default.is_absolute():
            tauri_default = repo_root / tauri_default
    else:
        tauri_default = repo_root / "apps" / "desktop" / "src-tauri" / "tauri.conf.json"

    parser.add_argument(
        "--tauri-config",
        default=tauri_default,
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
        expected_scheme = _normalize_scheme(args.url_scheme)
        if not expected_scheme:
            raise SystemExit("--url-scheme must be non-empty")
        expected_schemes = {expected_scheme}
    invalid_expected_schemes = {s for s in expected_schemes if ":" in s or "/" in s}
    if invalid_expected_schemes:
        print(
            "[macos] ERROR: invalid deep-link scheme(s) configured in tauri.conf.json (expected scheme names only, no ':' or '/')",
            file=sys.stderr,
        )
        print(f"[macos] Invalid scheme value(s): {_format_set(invalid_expected_schemes)}", file=sys.stderr)
        return 1

    with args.info_plist.open("rb") as f:
        plist = plistlib.load(f)
    if not isinstance(plist, dict):
        raise SystemExit(f"unexpected plist type in {args.info_plist}: {type(plist)}")

    doc_types = plist.get("CFBundleDocumentTypes")
    if doc_types is None:
        raise SystemExit("[macos] ERROR: CFBundleDocumentTypes missing from built Info.plist")
    if not isinstance(doc_types, list):
        raise SystemExit(
            f"[macos] ERROR: CFBundleDocumentTypes has unexpected type {type(doc_types)} in {args.info_plist}"
        )

    document_type_exts = _extensions_from_document_types(doc_types)
    document_type_utis = _utis_from_document_types(doc_types)
    uti_map = _uti_extension_map(plist)
    uti_exts: set[str] = set()
    for exts in uti_map.values():
        uti_exts |= exts
    registered_via_utis: set[str] = set()
    for uti in document_type_utis:
        registered_via_utis |= uti_map.get(uti, set())
    registered_exts = document_type_exts | registered_via_utis
    registered_schemes = extract_url_schemes(plist)
    invalid_schemes = {s for s in registered_schemes if ":" in s or "/" in s}
    if invalid_schemes:
        print(
            "[macos] ERROR: Info.plist declares invalid CFBundleURLSchemes value(s) (expected scheme names only, no ':' or '/')",
            file=sys.stderr,
        )
        print(f"[macos] Invalid scheme value(s): {_format_set(invalid_schemes)}", file=sys.stderr)
        return 1

    print(f"[macos] Info.plist: {args.info_plist}")
    print(f"[macos] Expected file extensions ({len(expected_exts)}): {_format_set(expected_exts)}")
    print(
        f"[macos] Observed file extensions (CFBundleDocumentTypes, {len(document_type_exts)}): {_format_set(document_type_exts)}"
    )
    if document_type_utis:
        print(f"[macos] Observed content types (LSItemContentTypes, {len(document_type_utis)}): {_format_set(document_type_utis)}")
    if uti_map:
        if uti_exts:
            print(f"[macos] Observed file extensions (UT*TypeDeclarations, {len(uti_exts)}): {_format_set(uti_exts)}")
        if registered_via_utis:
            print(
                f"[macos] Observed file extensions (via LSItemContentTypes -> UT*TypeDeclarations, {len(registered_via_utis)}): {_format_set(registered_via_utis)}"
            )
    print(
        f"[macos] Observed file extensions (combined, {len(registered_exts)}): {_format_set(registered_exts)}"
    )
    print(f"[macos] Expected URL schemes ({len(expected_schemes)}): {_format_set(expected_schemes)}")
    print(f"[macos] Observed URL schemes ({len(registered_schemes)}): {_format_set(registered_schemes)}")

    if not document_type_exts and not document_type_utis:
        print(
            "[macos] ERROR: CFBundleDocumentTypes did not declare any CFBundleTypeExtensions or LSItemContentTypes",
            file=sys.stderr,
        )
        return 1

    # File associations are driven by CFBundleDocumentTypes; UT*TypeDeclarations are helpful for
    # defining UTIs but are not sufficient on their own to register "open with" handling. We accept
    # extensions registered either directly via CFBundleTypeExtensions or indirectly via
    # CFBundleDocumentTypes.LSItemContentTypes + UT*TypeDeclarations defining filename extensions.
    missing_exts = expected_exts - registered_exts
    missing_schemes = expected_schemes - registered_schemes

    if missing_exts or missing_schemes:
        print("[macos] ERROR: bundle registration mismatch", file=sys.stderr)
        if missing_exts:
            print(
                f"[macos] Missing extensions: {_format_set(missing_exts)}",
                file=sys.stderr,
            )
            missing_only_in_uti = missing_exts & uti_exts
            if missing_only_in_uti:
                print(
                    f"[macos] Note: these extensions were present in UT*TypeDeclarations but did not appear to be registered via CFBundleTypeExtensions or LSItemContentTypes: {_format_set(missing_only_in_uti)}",
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
