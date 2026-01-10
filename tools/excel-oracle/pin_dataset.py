#!/usr/bin/env python3
"""
Pin an Excel oracle dataset for CI.

Typical flow on a Windows machine with Excel installed:

1) Generate oracle data:
   powershell -ExecutionPolicy Bypass -File tools/excel-oracle/run-excel-oracle.ps1 `
     -CasesPath tests/compatibility/excel-oracle/cases.json `
     -OutPath  tests/compatibility/excel-oracle/datasets/excel-oracle.json

2) Pin it (optionally versioned):
   python tools/excel-oracle/pin_dataset.py \
     --dataset tests/compatibility/excel-oracle/datasets/excel-oracle.json \
     --pinned tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json \
     --versioned-dir tests/compatibility/excel-oracle/datasets/versioned

The pinned dataset can be committed, allowing CI to validate engine behavior
even on Windows runners without Excel installed.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
from pathlib import Path
from typing import Any


def _load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _sanitize_fragment(text: str) -> str:
    # Keep filenames portable and reasonably readable.
    safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", text.strip())
    safe = re.sub(r"_+", "_", safe).strip("_")
    return safe or "unknown"


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--dataset", required=True, help="Path to generated oracle dataset JSON")
    p.add_argument(
        "--pinned",
        default="tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json",
        help="Where to copy the dataset for CI to consume (default: %(default)s)",
    )
    p.add_argument(
        "--versioned-dir",
        default="",
        help="If set, also write a version-tagged copy into this directory.",
    )
    args = p.parse_args()

    dataset_path = Path(args.dataset)
    pinned_path = Path(args.pinned)
    versioned_dir = Path(args.versioned_dir) if args.versioned_dir else None

    payload = _load_json(dataset_path)
    source = payload.get("source", {})
    case_set = payload.get("caseSet", {})

    if not isinstance(source, dict) or source.get("kind") != "excel":
        raise SystemExit(
            "Refusing to pin dataset that does not come from real Excel. "
            f"source.kind={source.get('kind')!r}"
        )

    excel_version = _sanitize_fragment(str(source.get("version", "unknown")))
    excel_build = _sanitize_fragment(str(source.get("build", "unknown")))
    cases_sha = _sanitize_fragment(str(case_set.get("sha256", "unknown")))

    pinned_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(dataset_path, pinned_path)
    print(f"Pinned dataset -> {pinned_path.as_posix()}")

    if versioned_dir is not None:
        versioned_dir.mkdir(parents=True, exist_ok=True)
        versioned_name = f"excel-{excel_version}-build-{excel_build}-cases-{cases_sha[:8]}.json"
        versioned_path = versioned_dir / versioned_name
        shutil.copyfile(dataset_path, versioned_path)
        print(f"Versioned copy -> {versioned_path.as_posix()}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
