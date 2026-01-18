#!/usr/bin/env python3
"""
Guardrail: prevent production Rust code from reintroducing panic-style constructs.

We scan production Rust sources under:
  - `crates/**/src/**/*.rs`
  - `apps/desktop/src-tauri/src/**/*.rs`
and fail if we find any of:
  - `.unwrap()`, `.unwrap_err()`, `.unwrap_unchecked()`
  - `.expect(...)`
  - `panic!(...)`, `unreachable!(...)`
  - debug-only output / placeholders: `println!`, `eprintln!`, `dbg!`, `todo!`, `unimplemented!`

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

from rust_guard_utils import should_scan_file, strip_cfg_test_items


RE_FORBIDDEN = [
    ("unwrap()", re.compile(r"\.\s*unwrap\s*\(\s*\)")),
    ("unwrap_err()", re.compile(r"\.\s*unwrap_err\s*\(\s*\)")),
    ("unwrap_unchecked()", re.compile(r"\.\s*unwrap_unchecked\s*\(\s*\)")),
    # Many internal parsers use an `expect(TokenKind)` method; only flag `expect("...")` / `expect(r#"..."#)`.
    ("expect(<string>)", re.compile(r"\.\s*expect\s*\(\s*(?:r#*)?\"")),
    ("panic!", re.compile(r"\bpanic!\s*\(")),
    ("unreachable!", re.compile(r"\bunreachable!\s*\(")),
    ("println!", re.compile(r"\bprintln!\s*\(")),
    ("eprintln!", re.compile(r"\beprintln!\s*\(")),
    ("dbg!", re.compile(r"\bdbg!\s*\(")),
    ("todo!", re.compile(r"\btodo!\s*\(")),
    ("unimplemented!", re.compile(r"\bunimplemented!\s*\(")),
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

    # Fast prefilter to avoid work on the vast majority of files.
    needles = [
        "unwrap",
        "expect",
        "panic!",
        "unreachable!",
        "println!",
        "eprintln!",
        "dbg!",
        "todo!",
        "unimplemented!",
    ]
    if not any(needle in text for needle in needles):
        return []

    lines = text.splitlines()
    filtered = strip_cfg_test_items(lines)

    findings: list[Finding] = []
    for idx, line in enumerate(filtered, start=1):
        if line.lstrip().startswith("//"):
            continue

        for kind, rx in RE_FORBIDDEN:
            if rx.search(line):
                # `println!/eprintln!/dbg!` are fine for CLI entrypoints (`src/main.rs`), but we
                # still want to enforce the unwrap/panic invariants everywhere.
                if (
                    kind in {"println!", "eprintln!", "dbg!"}
                    and path.name == "main.rs"
                    and len(path.parts) >= 2
                    and path.parts[-2] == "src"
                ):
                    break
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

    print("Found forbidden panic-style constructs in production Rust sources:\n", file=sys.stderr)
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

