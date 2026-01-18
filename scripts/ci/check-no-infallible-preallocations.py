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


RE_CFG_TEST = re.compile(r"^\s*#\s*\[\s*cfg\s*\((?P<inner>.*)\)\s*\]\s*$")
RE_TEST_ATTR = re.compile(
    r"^\s*#\s*\[\s*(?:[A-Za-z_]\w*::)*test(?:\s*\([^)]*\))?\s*\]\s*$"
)
RE_BENCH_ATTR = re.compile(r"^\s*#\s*\[\s*bench\s*\]\s*$")


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


def is_test_item_attr(line: str) -> bool:
    return RE_TEST_ATTR.match(line) is not None or RE_BENCH_ATTR.match(line) is not None


@dataclass
class BraceParseState:
    in_block_comment: bool = False
    in_string: bool = False
    string_escape: bool = False
    raw_string_hashes: int | None = None


def brace_delta_with_state(line: str, st: BraceParseState) -> int:
    # Count `{`/`}` in Rust code, ignoring those inside strings/chars/comments.
    delta = 0
    i = 0
    n = len(line)

    while i < n:
        ch = line[i]
        nxt = line[i + 1] if i + 1 < n else ""

        if st.in_block_comment:
            if ch == "*" and nxt == "/":
                st.in_block_comment = False
                i += 2
                continue
            i += 1
            continue

        if st.raw_string_hashes is not None:
            if ch == '"':
                hashes = st.raw_string_hashes
                if line[i + 1 : i + 1 + hashes] == ("#" * hashes):
                    st.raw_string_hashes = None
                    i += 1 + hashes
                    continue
            i += 1
            continue

        if st.in_string:
            if st.string_escape:
                st.string_escape = False
            else:
                if ch == "\\":
                    st.string_escape = True
                elif ch == '"':
                    st.in_string = False
            i += 1
            continue

        # Line comment start.
        if ch == "/" and nxt == "/":
            break

        # Block comment start.
        if ch == "/" and nxt == "*":
            st.in_block_comment = True
            i += 2
            continue

        # Raw string start: r###"..."### or br###"..."###.
        if ch in {"r", "b"}:
            prev_ok = i == 0 or not (line[i - 1].isalnum() or line[i - 1] == "_")
            if prev_ok:
                j = i
                if ch == "b" and nxt == "r":
                    j = i + 1
                if j < n and line[j] == "r":
                    k = j + 1
                    hashes = 0
                    while k < n and line[k] == "#":
                        hashes += 1
                        k += 1
                    if k < n and line[k] == '"':
                        st.raw_string_hashes = hashes
                        i = k + 1
                        continue

        # Normal string start.
        if ch == '"':
            st.in_string = True
            i += 1
            continue

        # Char literal / byte char literal. Only treat as a char literal if we can find the
        # closing `'` shortly after; otherwise it's likely a lifetime like `'a`.
        if ch == "'":
            j = i + 1
            esc = False
            found_close = False
            while j < n and (j - i) <= 8:
                cj = line[j]
                if esc:
                    esc = False
                else:
                    if cj == "\\":
                        esc = True
                    elif cj == "'":
                        found_close = True
                        break
                j += 1

            if found_close:
                i = j + 1
                continue

            i += 1
            continue

        if ch == "{":
            delta += 1
        elif ch == "}":
            delta -= 1

        i += 1

    return delta


def strip_cfg_test_items(lines: list[str]) -> list[str]:
    out: list[str] = []

    brace_depth: int = 0
    brace_state = BraceParseState()
    pending_cfg_test: bool = False
    pending_test_item: bool = False
    skipping: bool = False
    skip_until_depth: int | None = None

    for line in lines:
        if skipping:
            brace_depth += brace_delta_with_state(line, brace_state)
            if skip_until_depth is not None and brace_depth == skip_until_depth:
                skipping = False
                skip_until_depth = None
            continue

        if pending_cfg_test or pending_test_item:
            stripped = line.strip()
            if stripped == "":
                continue
            if line.lstrip().startswith("#["):
                continue
            if line.lstrip().startswith("//"):
                continue

            start_depth = brace_depth
            brace_depth += brace_delta_with_state(line, brace_state)

            # Most cfg(test) items are blocks (modules, fns, impls). If the next item doesn't open
            # a block, treat it as a single-line item and skip just this line.
            if brace_depth > start_depth:
                skipping = True
                skip_until_depth = start_depth
            pending_cfg_test = False
            pending_test_item = False
            continue

        if is_cfg_test_attr(line):
            pending_cfg_test = True
            continue

        if is_test_item_attr(line):
            pending_test_item = True
            continue

        out.append(line)
        brace_depth += brace_delta_with_state(line, brace_state)

    return out


def should_scan_file(path: Path) -> bool:
    if path.name in SKIP_BASENAMES:
        return False

    parts = path.parts
    try:
        src_idx = parts.index("src")
    except ValueError:
        return False

    # Skip `**/src/bin/**`.
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

