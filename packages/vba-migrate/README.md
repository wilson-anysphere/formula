# `@formula/vba-migrate`

AI-assisted VBA migration tooling (VBA → Python/TypeScript) with:

- **Analyzer**: identifies object model usage (`Range`/`Cells`/`Worksheets`), external references, and risky/unsupported constructs.
- **Converter**: uses an injected LLM client to produce Python (Formula Python API) and TypeScript (Formula scripting API) output.
- **Deterministic post-processing**: strips markdown fences, normalizes common VBA-ish artifacts (`.Value`, `.Formula`, `sheet.Range(...)`), and runs a compile/syntax check.
- **Validator**: runs the original macro via a **pluggable execution oracle** (default: Rust CLI powered by `crates/formula-vba-runtime`), runs the converted script, and reports actionable workbook diffs + mismatches.
- **Batch CLI**: `vba-migrate --dir <path>` scans a directory of `.xlsm` (or JSON fixtures), converts, validates, and emits a summary report.

This package is intentionally minimal: it is designed to be a testable pipeline that can later be wired into the full Excel object model and scripting runtimes.

## LLM integration

`VbaMigrator` accepts either:

- a `complete({ prompt, temperature }) -> string` style client (used by tests), or
- an `LLMClient.chat({ messages })` style client (see `packages/llm`).
