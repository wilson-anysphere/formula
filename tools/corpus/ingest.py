#!/usr/bin/env python3

from __future__ import annotations

import argparse
from pathlib import Path

from .crypto import get_fernet_key_from_env
from .sanitize import SanitizeOptions, sanitize_xlsx_bytes, scan_xlsx_bytes_for_leaks
from . import triage as triage_mod
from .util import WorkbookInput, ensure_dir, read_workbook_input, sha256_hex, utc_now_iso, write_json


def _triage_sanitized_workbook(workbook: WorkbookInput) -> dict:
    """Run corpus triage on an in-memory workbook blob.

    This is intentionally a small wrapper around `tools.corpus.triage.triage_workbook` so we can
    share defaults between `triage.py` and `ingest.py` and unit test the invocation without
    requiring a Rust toolchain.
    """

    rust_exe = triage_mod._build_rust_helper()
    diff_ignore = set(triage_mod.DEFAULT_DIFF_IGNORE)
    return triage_mod.triage_workbook(
        workbook,
        rust_exe=rust_exe,
        diff_ignore=diff_ignore,
        diff_limit=25,
        recalc=False,
        render_smoke=False,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Ingest an XLSX into the (private) compatibility corpus.")
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--corpus-dir", type=Path, default=Path("tools/corpus/private"))
    parser.add_argument(
        "--fernet-key-env",
        default="CORPUS_ENCRYPTION_KEY",
        help="Env var containing Fernet key used to encrypt the original workbook.",
    )

    parser.add_argument("--no-redact-cell-values", action="store_true")
    parser.add_argument("--hash-strings", action="store_true")
    parser.add_argument("--hash-salt", help="Required when --hash-strings is set.")
    parser.add_argument("--keep-external-links", action="store_true")
    parser.add_argument("--keep-secrets", action="store_true")
    parser.add_argument("--no-scrub-metadata", action="store_true")
    parser.add_argument(
        "--rename-sheets",
        action="store_true",
        help="Deterministically rename sheets to Sheet1, Sheet2, ... (updates formulas).",
    )
    parser.add_argument(
        "--no-leak-scan",
        action="store_true",
        help="Skip leak scanning the sanitized output (not recommended).",
    )
    parser.add_argument(
        "--leak-scan-string",
        action="append",
        default=[],
        help="Plaintext string expected not to appear in sanitized output. Can be repeated.",
    )
    parser.add_argument("--no-triage", action="store_true", help="Skip triage (faster ingest).")
    args = parser.parse_args()

    fernet_key = get_fernet_key_from_env(args.fernet_key_env)

    workbook = read_workbook_input(args.input)
    raw = workbook.data
    workbook_id = sha256_hex(raw)[:16]
    ext = "".join(Path(workbook.display_name).suffixes) or ".xlsx"

    corpus_dir: Path = args.corpus_dir
    originals_dir = corpus_dir / "originals"
    sanitized_dir = corpus_dir / "sanitized"
    metadata_dir = corpus_dir / "metadata"
    reports_dir = corpus_dir / "reports"
    for d in (originals_dir, sanitized_dir, metadata_dir, reports_dir):
        ensure_dir(d)

    from cryptography.fernet import Fernet

    f = Fernet(fernet_key.encode("utf-8"))
    (originals_dir / f"{workbook_id}{ext}.enc").write_bytes(f.encrypt(raw))

    options = SanitizeOptions(
        redact_cell_values=not args.no_redact_cell_values,
        hash_strings=args.hash_strings,
        hash_salt=args.hash_salt,
        remove_external_links=not args.keep_external_links,
        remove_secrets=not args.keep_secrets,
        scrub_metadata=not args.no_scrub_metadata,
        rename_sheets=args.rename_sheets,
    )
    sanitized_bytes, sanitize_summary = sanitize_xlsx_bytes(raw, options=options)

    if not args.no_leak_scan:
        scan = scan_xlsx_bytes_for_leaks(sanitized_bytes, plaintext_strings=args.leak_scan_string)
        if not scan.ok:
            print(f"Leak scan failed ({len(scan.findings)} findings); refusing to ingest.")
            for f in scan.findings[:25]:
                print(f"  {f.kind} in {f.part_name} sha256={f.match_sha256[:16]}")
            return 1

    sanitized_path = sanitized_dir / f"{workbook_id}{ext}"
    sanitized_path.write_bytes(sanitized_bytes)

    write_json(
        metadata_dir / f"{workbook_id}.json",
        {
            "id": workbook_id,
            "ingested_at": utc_now_iso(),
            "input_filename": args.input.name,
            "original_sha256": sha256_hex(raw),
            "sanitized_sha256": sha256_hex(sanitized_bytes),
            "sanitize_options": options.__dict__,
            "sanitize_summary": {
                "removed_parts": sanitize_summary.removed_parts,
                "rewritten_parts": sanitize_summary.rewritten_parts,
            },
        },
    )

    if not args.no_triage:
        report = _triage_sanitized_workbook(
            WorkbookInput(display_name=workbook.display_name, data=sanitized_bytes)
        )
        write_json(reports_dir / f"{workbook_id}.json", report)

    print(f"Ingested {args.input} -> {workbook_id}")
    print(f"  Encrypted original: {originals_dir / f'{workbook_id}{ext}.enc'}")
    print(f"  Sanitized:          {sanitized_path}")
    if not args.no_triage:
        print(f"  Triage report:      {reports_dir / f'{workbook_id}.json'}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
