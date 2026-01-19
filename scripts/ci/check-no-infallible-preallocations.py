#!/usr/bin/env python3
"""
Guardrail: prevent production Rust code from reintroducing infallible preallocations.

Many std collection constructors like `Vec::with_capacity(n)` and
`HashMap::with_capacity_and_hasher(n, ...)` abort on allocation failure (OOM). In
production, capacities can be derived from untrusted files/inputs; prefer
`*_::new()` plus best-effort `try_reserve*` to avoid immediate aborts.

We scan production Rust sources under:
  - `crates/**/src/**/*.rs`
  - `apps/desktop/src-tauri/src/**/*.rs`
and fail if we find common infallible preallocation constructors for standard collections.

We intentionally skip:
  - build scripts (`**/build.rs`) by virtue of scanning only `src/**`
  - `**/src/bin/**` (tooling binaries)
  - test-only modules/files (best-effort):
    - entire items annotated with `#[cfg(test)]` / `#[cfg(all(test, ...))]` / etc.
    - entire items annotated with `#[test]` / `#[tokio::test]` / `#[bench]` (even if not nested in a `#[cfg(test)] mod ...`)
    - files named `tests.rs`, `fuzz_tests.rs`, `tests_proptest.rs`

This is a pragmatic CI guard, not a full Rust parser; keep it fast and low-friction.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path

from rust_guard_utils import should_scan_file, strip_cfg_test_items


def _generic_turbofish() -> str:
    # Rough match for `::<...>::` (single line only).
    return r"(?:<[^\\n]*?>\s*::\s*)?"


RE_FORBIDDEN = [
    (
        "Vec::with_capacity",
        re.compile(rf"\b(?:std\s*::\s*vec\s*::\s*)?Vec\s*::\s*{_generic_turbofish()}with_capacity\s*\("),
    ),
    (
        "VecDeque::with_capacity",
        re.compile(
            rf"\b(?:std\s*::\s*collections\s*::\s*)?VecDeque\s*::\s*{_generic_turbofish()}with_capacity\s*\("
        ),
    ),
    (
        "String::with_capacity",
        re.compile(r"\b(?:std\s*::\s*string\s*::\s*)?String\s*::\s*with_capacity\s*\("),
    ),
    (
        "HashMap::with_capacity",
        re.compile(
            rf"\b(?:std\s*::\s*collections\s*::\s*)?HashMap\s*::\s*{_generic_turbofish()}with_capacity\s*\("
        ),
    ),
    (
        "HashMap::with_capacity_and_hasher",
        re.compile(
            rf"\b(?:std\s*::\s*collections\s*::\s*)?HashMap\s*::\s*{_generic_turbofish()}with_capacity_and_hasher\s*\("
        ),
    ),
    (
        "HashSet::with_capacity",
        re.compile(
            rf"\b(?:std\s*::\s*collections\s*::\s*)?HashSet\s*::\s*{_generic_turbofish()}with_capacity\s*\("
        ),
    ),
    (
        "HashSet::with_capacity_and_hasher",
        re.compile(
            rf"\b(?:std\s*::\s*collections\s*::\s*)?HashSet\s*::\s*{_generic_turbofish()}with_capacity_and_hasher\s*\("
        ),
    ),
]

SKIP_BASENAMES = {
    "tests.rs",
    "fuzz_tests.rs",
    "tests_proptest.rs",
}


@dataclass(frozen=True)
class Finding:
    path: Path
    line: int
    kind: str
    snippet: str


def scan_file(path: Path) -> list[Finding]:
    text = path.read_text(encoding="utf-8", errors="replace")

    needles = [
        "with_capacity",
        "with_capacity_and_hasher",
        "VecDeque",
        "HashMap",
        "HashSet",
        "String",
    ]
    if not any(needle in text for needle in needles):
        return []

    lines = text.splitlines()
    filtered = strip_cfg_test_items(lines)

    findings: list[Finding] = []
    for idx, line in enumerate(filtered, start=1):
        stripped = line.lstrip()
        if stripped.startswith("//"):
            continue

        # Best-effort: ignore trailing line comments to avoid false positives in commentary.
        code = line.split("//", 1)[0]
        if code.strip() == "":
            continue

        for kind, rx in RE_FORBIDDEN:
            if rx.search(code):
                findings.append(
                    Finding(
                        path=path,
                        line=idx,
                        kind=kind,
                        snippet=line.rstrip("\n"),
                    )
                )
                break

    return findings


def main() -> int:
    repo_root = Path(__file__).resolve().parents[2]
    scan_roots = [
        repo_root / "crates",
        repo_root / "apps" / "desktop" / "src-tauri",
    ]

    for root in scan_roots:
        if not root.is_dir():
            print(f"error: expected Rust source root at {root}", file=sys.stderr)
            return 2

    all_findings: list[Finding] = []
    for root in scan_roots:
        for path in root.rglob("*.rs"):
            if not should_scan_file(path, SKIP_BASENAMES):
                continue
            all_findings.extend(scan_file(path))

    if not all_findings:
        return 0

    print("Found forbidden infallible preallocations in production Rust sources:\n", file=sys.stderr)
    for f in all_findings[:200]:
        rel = f.path.relative_to(repo_root)
        print(f"- {rel}:{f.line}: {f.kind}: {f.snippet}", file=sys.stderr)

    if len(all_findings) > 200:
        print(
            f"\n... and {len(all_findings) - 200} more matches (showing first 200)",
            file=sys.stderr,
        )

    return 1


if __name__ == "__main__":
    raise SystemExit(main())

