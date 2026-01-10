# `@formula/vba-migrate`

AI-assisted VBA migration tooling (VBA → Python/TypeScript) with:

- **Analyzer**: identifies object model usage (`Range`/`Cells`/`Worksheets`), external references, and risky/unsupported constructs.
- **Converter**: uses an injected LLM client to produce Python (Formula Python API) and TypeScript (Formula scripting API) output.
- **Deterministic post-processing**: strips markdown fences, normalizes common VBA-ish artifacts (`.Value`, `.Formula`, `sheet.Range(...)`), and runs a compile/syntax check.
- **Validator**: executes a small supported subset of VBA and the converted script against the same in-memory workbook and reports cell diffs + mismatches.

This package is intentionally minimal: it is designed to be a testable pipeline that can later be wired into the real VBA parser/executor and scripting runtimes.

