#!/usr/bin/env python3
"""
Guardrail: ensure all tracked TOML files parse cleanly.

Why:
- Duplicate keys (a common copy/paste/merge error) are invalid TOML and break Cargo/workflows.
- Failing fast in CI keeps regressions small and avoids confusing downstream failures.

Notes:
- Uses Python's stdlib `tomllib` (Python 3.11+), which rejects duplicate keys.
"""

from __future__ import annotations

import subprocess
import sys
import tomllib


def git_ls_files() -> list[str]:
    out = subprocess.check_output(["git", "ls-files"], text=True)
    return [line.strip() for line in out.splitlines() if line.strip()]


def parse_toml(path: str) -> None:
    with open(path, "rb") as f:
        tomllib.load(f)


def main() -> int:
    files = [p for p in git_ls_files() if p.endswith(".toml")]
    if not files:
        print("No tracked .toml files found; skipping TOML parse guard.", file=sys.stderr)
        return 0

    failed = False
    for path in files:
        try:
            parse_toml(path)
        except Exception as err:  # noqa: BLE001 - guard script; surface any parse error
            failed = True
            print(f"error: failed to parse TOML: {path}", file=sys.stderr)
            print(f"  {err}", file=sys.stderr)

    if failed:
        return 1

    print(f"TOML parse: OK ({len(files)} files).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

