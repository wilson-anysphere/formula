#!/usr/bin/env python3

from __future__ import annotations

import argparse
import base64
import json
from pathlib import Path
from typing import Any

from .sanitize import SanitizeOptions, sanitize_xlsx_bytes, scan_xlsx_bytes_for_leaks
from .triage import DEFAULT_DIFF_IGNORE, _build_rust_helper, triage_workbook
from .util import WorkbookInput, ensure_dir, load_json, read_workbook_input, sha256_hex, write_json


def _encode_b64_text(data: bytes) -> bytes:
    """Return base64-encoded bytes suitable for writing to `*.b64` fixtures.

    We intentionally use `encodebytes` to wrap long lines and include a trailing newline,
    matching the style of existing fixtures in `tools/corpus/public/`.
    """

    # `encodebytes` always appends a trailing newline.
    return base64.encodebytes(data)


def _write_public_fixture(
    out_path: Path, workbook_bytes: bytes, *, force: bool
) -> None:
    """Write `workbook_bytes` to `out_path` as base64, refusing to overwrite unless `force`."""

    if out_path.exists():
        existing = read_workbook_input(out_path).data
        if existing == workbook_bytes:
            # Idempotent: already up to date.
            return
        if not force:
            raise FileExistsError(
                f"Refusing to overwrite existing fixture {out_path} (bytes differ). "
                f"Re-run with --force to overwrite."
            )

    ensure_dir(out_path.parent)
    out_path.write_bytes(_encode_b64_text(workbook_bytes))


def extract_public_expectations(report: dict[str, Any]) -> dict[str, Any]:
    """Extract the public expectations subset from a triage report.

    The public expectations file is used as a regression gate in PR CI. Keep it minimal
    and privacy-safe.
    """

    result = report.get("result") or {}
    open_ok = result.get("open_ok")
    if open_ok is not True:
        raise ValueError(
            "Refusing to promote: triage report indicates open_ok is not true "
            f"(open_ok={open_ok!r})."
        )

    round_trip_ok_raw = result.get("round_trip_ok")
    if not isinstance(round_trip_ok_raw, bool):
        # Be strict: we expect triage to set this to a bool even if diff failed.
        raise ValueError(
            "Refusing to promote: triage report missing boolean round_trip_ok "
            f"(round_trip_ok={round_trip_ok_raw!r})."
        )

    diff_critical = result.get("diff_critical_count", 0)
    if not isinstance(diff_critical, int):
        raise ValueError(
            "Refusing to promote: triage report missing integer diff_critical_count "
            f"(diff_critical_count={diff_critical!r})."
        )

    # Requirements: open_ok must be true in the expectations entry.
    return {
        "open_ok": True,
        "round_trip_ok": round_trip_ok_raw,
        "diff_critical_count": diff_critical,
    }


def upsert_expectations_entry(
    *,
    expectations: dict[str, Any],
    workbook_name: str,
    entry: dict[str, Any],
    force: bool,
) -> tuple[dict[str, Any], bool]:
    """Insert/update one expectations entry.

    Returns: (updated_expectations, changed)
    """

    current = expectations.get(workbook_name)
    if current == entry:
        return expectations, False
    if current is not None and not force:
        raise FileExistsError(
            f"Refusing to overwrite existing expectations entry for {workbook_name}. "
            "Re-run with --force to overwrite."
        )
    out = dict(expectations)
    out[workbook_name] = entry
    return out, True


def update_public_expectations_file(
    expectations_path: Path, *, workbook_name: str, report: dict[str, Any], force: bool
) -> bool:
    """Update (or create) the public expectations file for one workbook.

    Returns True if the file content changed.
    """

    expectations: dict[str, Any]
    if expectations_path.exists():
        expectations = load_json(expectations_path)
        if not isinstance(expectations, dict):
            raise ValueError(f"Expected {expectations_path} to be a JSON object mapping names -> expectations.")
    else:
        expectations = {}

    entry = extract_public_expectations(report)
    updated, changed = upsert_expectations_entry(
        expectations=expectations, workbook_name=workbook_name, entry=entry, force=force
    )
    if changed:
        write_json(expectations_path, updated)
    return changed


def _run_public_triage(workbook: WorkbookInput, *, diff_limit: int = 25) -> dict[str, Any]:
    rust_exe = _build_rust_helper()
    return triage_workbook(
        workbook,
        rust_exe=rust_exe,
        diff_ignore=set(DEFAULT_DIFF_IGNORE),
        diff_limit=diff_limit,
        recalc=False,
        render_smoke=False,
    )


def _coerce_display_name(name: str, *, default_ext: str) -> str:
    name = name.strip()
    if not name:
        raise ValueError("--name must not be empty")
    if name.endswith((".xlsx", ".xlsm")):
        return name
    if name.endswith(".b64"):
        # Be helpful if the user passes the fixture filename itself.
        name = name[: -len(".b64")]
    if not name.endswith((".xlsx", ".xlsm")):
        name += default_ext
    return name


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Promote a sanitized workbook into the public compatibility corpus subset."
    )
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument(
        "--name",
        help=(
            "Optional fixture name (without extension). Defaults to the input filename. "
            "Examples: 'my-case' or 'my-case.xlsx'."
        ),
    )
    parser.add_argument("--public-dir", type=Path, default=Path("tools/corpus/public"))
    parser.add_argument("--force", action="store_true", help="Overwrite existing fixture/expectations.")

    # Safety controls
    parser.add_argument(
        "--confirm-sanitized",
        action="store_true",
        help=(
            "Assert that the input workbook is already sanitized and safe to publish. "
            "When set, leak scanning is skipped (not recommended)."
        ),
    )

    # Optional sanitization (mirrors `tools/corpus/ingest.py` flags).
    parser.add_argument("--sanitize", action="store_true", help="Run sanitizer on the input bytes first.")
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
        "--leak-scan-string",
        action="append",
        default=[],
        help="Plaintext string expected not to appear in output. Can be repeated.",
    )

    parser.add_argument(
        "--triage-out",
        type=Path,
        default=Path("tools/corpus/out/promote-public"),
        help="Directory for writing the generated triage report (gitignored).",
    )
    args = parser.parse_args()

    wb_in = read_workbook_input(args.input)
    display_name = (
        _coerce_display_name(args.name, default_ext=Path(wb_in.display_name).suffix) if args.name else wb_in.display_name
    )

    # Note: `WorkbookInput` must not contain local paths; use `display_name`.
    workbook_bytes = wb_in.data

    if args.sanitize:
        options = SanitizeOptions(
            redact_cell_values=not args.no_redact_cell_values,
            hash_strings=args.hash_strings,
            hash_salt=args.hash_salt,
            remove_external_links=not args.keep_external_links,
            remove_secrets=not args.keep_secrets,
            scrub_metadata=not args.no_scrub_metadata,
            rename_sheets=args.rename_sheets,
        )
        workbook_bytes, _summary = sanitize_xlsx_bytes(workbook_bytes, options=options)

    if not args.confirm_sanitized:
        scan = scan_xlsx_bytes_for_leaks(workbook_bytes, plaintext_strings=args.leak_scan_string)
        if not scan.ok:
            print(
                f"Leak scan failed ({len(scan.findings)} findings); refusing to promote to public corpus."
            )
            for f in scan.findings[:25]:
                print(f"  {f.kind} in {f.part_name} sha256={f.match_sha256[:16]}")
            return 1

    public_dir: Path = args.public_dir
    fixture_path = public_dir / f"{display_name}.b64"
    expectations_path = public_dir / "expectations.json"

    # 1) Run triage against the exact bytes we intend to publish.
    report = _run_public_triage(WorkbookInput(display_name=display_name, data=workbook_bytes))

    triage_out: Path = args.triage_out
    ensure_dir(triage_out)
    report_id = (report.get("sha256") or sha256_hex(workbook_bytes))[:16]
    report_path = triage_out / f"{report_id}.json"
    write_json(report_path, report)

    # 2) Validate we won't overwrite existing entries unless --force is set.
    try:
        entry = extract_public_expectations(report)

        if fixture_path.exists():
            existing_fixture_bytes = read_workbook_input(fixture_path).data
            if existing_fixture_bytes != workbook_bytes and not args.force:
                raise FileExistsError(
                    f"Refusing to overwrite existing fixture {fixture_path} (bytes differ). "
                    "Re-run with --force to overwrite."
                )

        current_expectations: dict[str, Any] = {}
        if expectations_path.exists():
            current_expectations = load_json(expectations_path)
            if not isinstance(current_expectations, dict):
                raise ValueError(
                    f"Expected {expectations_path} to be a JSON object mapping names -> expectations."
                )

        _updated, _changed = upsert_expectations_entry(
            expectations=current_expectations,
            workbook_name=display_name,
            entry=entry,
            force=args.force,
        )
    except (FileExistsError, ValueError) as e:
        print(str(e))
        return 1

    # 3) Write the public fixture and upsert expectations.json.
    try:
        _write_public_fixture(fixture_path, workbook_bytes, force=args.force)
    except FileExistsError as e:
        print(str(e))
        return 1

    # Safe to update now; we already validated overwrites above.
    changed = update_public_expectations_file(
        expectations_path, workbook_name=display_name, report=report, force=args.force
    )

    # Optional: write a convenience summary for humans.
    summary = {
        "fixture": str(fixture_path),
        "expectations": str(expectations_path),
        "expectations_changed": changed,
        "triage_report": str(report_path),
    }
    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
