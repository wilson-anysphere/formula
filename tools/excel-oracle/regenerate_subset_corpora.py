#!/usr/bin/env python3
"""
Regenerate small subset corpora under tools/excel-oracle/ from the canonical cases corpus.

Why this exists
---------------

The canonical Excel-oracle corpus lives at:

  tests/compatibility/excel-oracle/cases.json

We also keep a few *small* subset corpora (under tools/excel-oracle/) to make it easy to run
targeted compatibility checks in real Excel (Windows + COM automation) and then merge the results
back into the pinned dataset by canonical caseId.

Those subset corpora must stay aligned with the canonical corpus (same case IDs/formulas), so we
can patch/pin datasets without manually editing JSON.

This script rewrites the subset files deterministically by filtering the canonical corpus by tags.

Files written
-------------

- tools/excel-oracle/odd_coupon_validation_cases.json
  - All cases tagged `odd_coupon_validation`

- tools/excel-oracle/odd_coupon_boundary_cases.json
  - All cases tagged `odd_coupon` AND `boundary`

- tools/excel-oracle/odd_coupon_long_stub_cases.json
  - All cases tagged `odd_coupon` AND `long_stub`

- tools/excel-oracle/odd_coupon_basis4_cases.json
  - All cases tagged `odd_coupon` AND `basis4`

- tools/excel-oracle/odd_coupon_invalid_schedule_cases.json
  - All cases tagged `odd_coupon` AND `invalid_schedule`

After running, commit the resulting diffs.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def _load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=False) + "\n",
        encoding="utf-8",
        newline="\n",
    )


def _has_tag(case: dict[str, Any], tag: str) -> bool:
    tags = case.get("tags", [])
    return isinstance(tags, list) and tag in tags


def _resolve_under_repo_root(repo_root: Path, raw: str) -> Path:
    p = Path(raw)
    if p.is_absolute():
        return p
    return repo_root / p


def _expected_subset_payload(*, case_set: str, cases: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "schemaVersion": 1,
        "caseSet": case_set,
        "defaultSheet": "Sheet1",
        "cases": cases,
    }


def main() -> int:
    p = argparse.ArgumentParser(
        description=(
            "Regenerate small convenience subset corpora under tools/excel-oracle/ "
            "from the canonical cases corpus."
        )
    )
    p.add_argument(
        "--cases",
        default="tests/compatibility/excel-oracle/cases.json",
        help="Path to canonical cases.json (default: %(default)s)",
    )
    p.add_argument(
        "--out-dir",
        default="tools/excel-oracle",
        help="Directory to write subset corpora into (default: %(default)s)",
    )
    p.add_argument(
        "--check",
        action="store_true",
        help="Validate that the subset corpora are up to date, without rewriting files.",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Print what would be written (case counts + paths) without creating or modifying any files.",
    )
    args = p.parse_args()

    repo_root = Path(__file__).resolve().parents[2]
    corpus_path = _resolve_under_repo_root(repo_root, args.cases).resolve()
    out_dir = _resolve_under_repo_root(repo_root, args.out_dir).resolve()

    corpus = _load_json(corpus_path)
    cases = [c for c in corpus.get("cases", []) if isinstance(c, dict)]

    validation_cases = [c for c in cases if _has_tag(c, "odd_coupon_validation")]
    boundary_cases = [c for c in cases if _has_tag(c, "odd_coupon") and _has_tag(c, "boundary")]
    long_stub_cases = [c for c in cases if _has_tag(c, "odd_coupon") and _has_tag(c, "long_stub")]
    basis4_cases = [c for c in cases if _has_tag(c, "odd_coupon") and _has_tag(c, "basis4")]
    invalid_schedule_cases = [
        c for c in cases if _has_tag(c, "odd_coupon") and _has_tag(c, "invalid_schedule")
    ]

    targets = [
        (
            out_dir / "odd_coupon_validation_cases.json",
            _expected_subset_payload(
                case_set="financial-odd-coupon-validation", cases=validation_cases
            ),
            len(validation_cases),
        ),
        (
            out_dir / "odd_coupon_boundary_cases.json",
            _expected_subset_payload(case_set="odd-coupon-boundaries", cases=boundary_cases),
            len(boundary_cases),
        ),
        (
            out_dir / "odd_coupon_long_stub_cases.json",
            _expected_subset_payload(case_set="financial-odd-coupon-long", cases=long_stub_cases),
            len(long_stub_cases),
        ),
        (
            out_dir / "odd_coupon_basis4_cases.json",
            _expected_subset_payload(case_set="financial-odd-coupon-basis4", cases=basis4_cases),
            len(basis4_cases),
        ),
        (
            out_dir / "odd_coupon_invalid_schedule_cases.json",
            _expected_subset_payload(
                case_set="financial-odd-coupon-invalid-schedule", cases=invalid_schedule_cases
            ),
            len(invalid_schedule_cases),
        ),
    ]

    if args.check:
        mismatched: list[Path] = []
        for path, expected, _count in targets:
            if not path.is_file():
                mismatched.append(path)
                continue
            actual = _load_json(path)
            if actual != expected:
                mismatched.append(path)
        if mismatched:
            rels = []
            for m in mismatched:
                try:
                    rels.append(m.resolve().relative_to(repo_root.resolve()).as_posix())
                except Exception:
                    rels.append(m.as_posix())
            sys_err = "\n".join(f"- {r}" for r in rels)
            raise SystemExit(
                "Subset corpora are out of date. Re-run:\n"
                "  python tools/excel-oracle/regenerate_subset_corpora.py\n"
                f"Mismatched files:\n{sys_err}"
            )
        print("Subset corpora are up to date.")
        return 0

    if args.dry_run:
        for path, _payload, count in targets:
            # Print paths relative to repo root for stability/readability.
            try:
                rel = path.resolve().relative_to(repo_root.resolve()).as_posix()
            except Exception:
                rel = path.as_posix()
            print(f"Would write {count} cases -> {rel}")
        return 0

    for path, payload, count in targets:
        _write_json(path, payload)
        # Print paths relative to repo root for stability/readability.
        try:
            rel = path.resolve().relative_to(repo_root.resolve()).as_posix()
        except Exception:
            rel = path.as_posix()
        print(f"Wrote {count} cases -> {rel}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
