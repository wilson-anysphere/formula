# External workbook references — task tracker

This document exists to keep the “external workbook references” workstream coherent and to avoid
duplicated follow-up work across the task queue.

## DONE (landed)

- DONE — Debug trace external 3D spans with sheet order (`699939da`)
- DONE — External structured refs via provider metadata (`971bc00e`)
- DONE — Database functions external key parsing refactor (`30dfd78d`)
- DONE — Parser roundtrip for external 3D spans (`b0b3a606`)
- DONE — Debug trace path-qualified external refs (`e8a16064`)
- DONE — Task 375: External 3D precedents in `Engine::precedents`
- DONE — Task 377: `SHEET()` external index when sheet order exists
- DONE — Task 380: External invalidation (sheet/workbook) + external dependency indexing
- DONE — Task 384: Bytecode direct external refs
- DONE — Task 385: Docs

## Remaining (open)

- TODO — Task 379: Unit tests for external key parsing helpers
  - Pointers: `crates/formula-engine/src/eval/evaluator.rs`
    - `split_external_sheet_key`
    - `split_external_sheet_span_key`
- TODO — Task 381: Dynamic external deps (ensure dependency tracing/indexing covers dynamic ref producers)
  - Pointers:
    - `crates/formula-engine/src/engine.rs`
      - `analyze_external_dependencies` + `set_cell_external_refs`
- TODO — Task 383: Copy rewrite for external refs
  - Pointers: `crates/formula-engine/src/editing/rewrite.rs` (`rewrite_formula_for_copy_delta`, etc.)
- TODO — Bytecode: make `INDIRECT` reject external workbook refs (match AST + `tests/indirect.rs`)
  - Pointers:
    - `crates/formula-engine/src/bytecode/runtime.rs` (`fn_indirect`, external sheet handling)
    - `crates/formula-engine/tests/bytecode_external_refs.rs` (currently expects INDIRECT external refs to resolve)
