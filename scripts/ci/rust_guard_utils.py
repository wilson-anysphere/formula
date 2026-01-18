#!/usr/bin/env python3
"""
Shared helpers for Rust CI guard scripts.

These scripts are intentionally *not* full Rust parsers. They implement:
- a best-effort "skip test-only items" filter (`strip_cfg_test_items`)
- a minimal brace counter that ignores strings/chars/comments
- a common file eligibility predicate for "production Rust sources"

Keep this module dependency-free and fast.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


RE_CFG_TEST = re.compile(r"^\s*#\s*\[\s*cfg\s*\((?P<inner>.*)\)\s*\]\s*$")
RE_TEST_ATTR = re.compile(
    r"^\s*#\s*\[\s*(?:[A-Za-z_]\w*::)*test(?:\s*\([^)]*\))?\s*\]\s*$"
)
RE_BENCH_ATTR = re.compile(r"^\s*#\s*\[\s*bench\s*\]\s*$")


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
    """
    Count `{`/`}` in Rust code, ignoring those inside strings/chars/comments.

    This keeps the cfg(test)/#[test] skipper robust even when test helpers contain
    strings like `" }); "` or raw strings with lots of braces.
    """
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

        # Char literal / byte char literal.
        # Only treat as a char literal if we can find the closing `'` shortly after;
        # otherwise it's likely a lifetime like `'a`.
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


def should_scan_file(path: Path, skip_basenames: set[str]) -> bool:
    if path.name in skip_basenames:
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

