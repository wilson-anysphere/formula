#!/usr/bin/env python3
"""
Guardrail: prevent production Rust code from reintroducing infallible preallocations.

Many std collection constructors like `Vec::with_capacity(n)` and
`HashMap::with_capacity_and_hasher(n, ...)` abort on allocation failure (OOM). In
production, capacities can be derived from untrusted files/inputs; prefer
`*_::new()` plus best-effort `try_reserve*` to avoid immediate aborts.

We scan Rust sources under `crates/**/src/**/*.rs` and fail if we find common
infallible preallocation constructors for standard collections.

We intentionally skip:
  - `crates/**/src/bin/**` (tooling binaries)
  - test-only modules/files (best-effort):
    - entire items annotated with `#[cfg(test)]` / `#[cfg(all(test, ...))]` / etc.
    - files named `tests.rs`, `fuzz_tests.rs`, `tests_proptest.rs`

This is a pragmatic CI guard, not a full Rust parser; keep it fast and low-friction.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path


RE_CFG_TEST = re.compile(r"^\s*#\s*\[\s*cfg\s*\((?P<inner>.*)\)\s*\]\s*$")


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


def is_cfg_test_attr(line: str) -> bool:
    m = RE_CFG_TEST.match(line)
    if not m:
        return False
    inner = m.group("inner")
    return re.search(r"\btest\b", inner) is not None


def brace_delta(line: str) -> int:
    # Heuristic. We do not attempt to fully parse Rust; this is good enough
    # for skipping common `#[cfg(test)] mod tests { ... }` blocks.
    return line.count("{") - line.count("}")


def strip_cfg_test_items(lines: list[str]) -> list[str]:
    out: list[str] = []

    brace_depth: int = 0
    pending_cfg_test: bool = False
    skipping: bool = False
    skip_until_depth: int | None = None

    for line in lines:
        if skipping:
            brace_depth += brace_delta(line)
            if skip_until_depth is not None and brace_depth == skip_until_depth:
                skipping = False
                skip_until_depth = None
            continue

        if pending_cfg_test:
            stripped = line.strip()
            if stripped == "":
                continue
            if line.lstrip().startswith("#["):
                continue
            if line.lstrip().startswith("//"):
                continue

            start_depth = brace_depth
            brace_depth += brace_delta(line)

            # Most cfg(test) items are blocks (modules, fns, impls). If the next item doesn't open
            # a block, treat it as a single-line item and skip just this line.
            if brace_depth > start_depth:
                skipping = True
                skip_until_depth = start_depth
            pending_cfg_test = False
            continue

        if is_cfg_test_attr(line):
            pending_cfg_test = True
            continue

        out.append(line)
        brace_depth += brace_delta(line)

    return out


def should_scan_file(path: Path) -> bool:
    if path.name in SKIP_BASENAMES:
        return False

    parts = path.parts
    try:
        src_idx = parts.index("src")
    except ValueError:
        return False

    # Skip `crates/**/src/bin/**`.
    if src_idx + 1 < len(parts) and parts[src_idx + 1] == "bin":
        return False

    return True


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
    crates_dir = repo_root / "crates"

    if not crates_dir.is_dir():
        print(f"error: expected crates dir at {crates_dir}", file=sys.stderr)
        return 2

    all_findings: list[Finding] = []
    for path in crates_dir.rglob("*.rs"):
        if not should_scan_file(path):
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

