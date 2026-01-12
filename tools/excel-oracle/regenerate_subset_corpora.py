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

- tools/excel-oracle/odd_coupon_invalid_schedule_cases.json
  - All cases tagged `odd_coupon` AND `invalid_schedule`

After running, commit the resulting diffs.
"""

from __future__ import annotations

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


def main() -> int:
    repo_root = Path(__file__).resolve().parents[2]
    corpus_path = repo_root / "tests" / "compatibility" / "excel-oracle" / "cases.json"
    corpus = _load_json(corpus_path)

    cases = [c for c in corpus.get("cases", []) if isinstance(c, dict)]

    validation_cases = [c for c in cases if _has_tag(c, "odd_coupon_validation")]
    boundary_cases = [c for c in cases if _has_tag(c, "odd_coupon") and _has_tag(c, "boundary")]
    long_stub_cases = [c for c in cases if _has_tag(c, "odd_coupon") and _has_tag(c, "long_stub")]
    invalid_schedule_cases = [
        c for c in cases if _has_tag(c, "odd_coupon") and _has_tag(c, "invalid_schedule")
    ]

    _write_json(
        repo_root / "tools" / "excel-oracle" / "odd_coupon_validation_cases.json",
        {
            "schemaVersion": 1,
            "caseSet": "financial-odd-coupon-validation",
            "defaultSheet": "Sheet1",
            "cases": validation_cases,
        },
    )
    _write_json(
        repo_root / "tools" / "excel-oracle" / "odd_coupon_boundary_cases.json",
        {
            "schemaVersion": 1,
            "caseSet": "odd-coupon-boundaries",
            "defaultSheet": "Sheet1",
            "cases": boundary_cases,
        },
    )
    _write_json(
        repo_root / "tools" / "excel-oracle" / "odd_coupon_long_stub_cases.json",
        {
            "schemaVersion": 1,
            "caseSet": "financial-odd-coupon-long",
            "defaultSheet": "Sheet1",
            "cases": long_stub_cases,
        },
    )
    _write_json(
        repo_root / "tools" / "excel-oracle" / "odd_coupon_invalid_schedule_cases.json",
        {
            "schemaVersion": 1,
            "caseSet": "financial-odd-coupon-invalid-schedule",
            "defaultSheet": "Sheet1",
            "cases": invalid_schedule_cases,
        },
    )

    print(f"Wrote {len(validation_cases)} cases -> tools/excel-oracle/odd_coupon_validation_cases.json")
    print(f"Wrote {len(boundary_cases)} cases -> tools/excel-oracle/odd_coupon_boundary_cases.json")
    print(f"Wrote {len(long_stub_cases)} cases -> tools/excel-oracle/odd_coupon_long_stub_cases.json")
    print(
        f"Wrote {len(invalid_schedule_cases)} cases -> tools/excel-oracle/odd_coupon_invalid_schedule_cases.json"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
