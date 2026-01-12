# Excel Feature Parity (code-driven)

This document replaces the old hand-maintained “✅/⬜ per function” checklists, which quickly drift
and become misleading.

Instead, parity is reported from authoritative, versioned data sources in the repo and validated by
tests/scripts.

## Sources of truth

### Implemented function catalog (authoritative)

The source of truth for “what the engine currently implements” is:

- **[`shared/functionCatalog.json`](../shared/functionCatalog.json)** (and
  [`shared/functionCatalog.mjs`](../shared/functionCatalog.mjs))
  - Kept in sync with the Rust registry by
    **[`crates/formula-engine/tests/function_catalog_sync.rs`](../crates/formula-engine/tests/function_catalog_sync.rs)**.
  - Used by downstream tooling (JS/TS, docs, scripts) without having to compile Rust.

For a current count snapshot (generated), see [“Current snapshot (counts only)”](#current-snapshot-counts-only)
below.

Why it’s committed:

- Stable interface for non-Rust tooling (autocomplete, docs, parity scripts, etc).
- Reviewable diffs (sorted, normalized names).
- CI enforces correctness via `function_catalog_sync`.

### Excel “function universe” references (informational)

Excel has multiple naming/encoding systems and many “functions” that aren’t worksheet functions
(legacy XLM macro functions, reserved slots, etc). The following sources are useful for context and
coverage tracking:

- **BIFF/XLSB built-in function table (FTAB)**:
  [`crates/formula-biff/src/ftab.rs`](../crates/formula-biff/src/ftab.rs)
  - Lists BIFF `iftab` ids `0..=484` and their canonical names.
  - **Does not** include many modern functions that Excel stores as “future/UDF” calls in BIFF.
- **Excel oracle corpus**:
  [`tests/compatibility/excel-oracle/cases.json`](../tests/compatibility/excel-oracle/cases.json)
  - A curated set of formulas + expected outputs, used to validate behavior against Excel.

## Regenerating / checking

Update the committed function catalog (requires a Rust toolchain):

```bash
pnpm -w run generate:function-catalog
```

(Under the hood this runs `node scripts/generate-function-catalog.js`.)

Validate that the committed catalog matches the Rust registry:

```bash
bash scripts/cargo_agent.sh test -p formula-engine function_catalog_sync
```

Generate a parity report between `shared/functionCatalog.json` and the BIFF FTAB name table:

```bash
pnpm -w run report:function-parity
```

Lightweight parity summary (also includes oracle corpus stats):

```bash
node tools/parity/report_functions.mjs
```

Useful report flags:

- `node tools/parity/report_functions.mjs --list-missing` (FTAB names missing from the engine catalog)
- `node tools/parity/report_functions.mjs --list-oracle-missing` (oracle function-like tokens missing
  from the engine catalog; may include intentional “unknown function” test cases)

Note: FTAB includes many legacy XLM/macro-only functions and reserved slots; the FTAB-based “missing”
count is **not** a direct “worksheet function parity” percentage.

## Current snapshot (counts only)

This section is a snapshot of the summary produced by:

```bash
pnpm -w run report:function-parity
```

To regenerate: run the command above and replace the block between the markers below.

To update this file automatically:

```bash
pnpm -w run report:function-parity -- --update-doc
```

<!-- BEGIN GENERATED: report-function-parity -->
```text
Function parity report (catalog ↔ BIFF FTAB)

Catalog functions (shared/functionCatalog.json): 428
FTAB functions (crates/formula-biff/src/ftab.rs): 478
Catalog ∩ FTAB (case-insensitive name match): 324
FTAB \ Catalog (missing from catalog): 154
Catalog \ FTAB (not present in FTAB): 104
```
<!-- END GENERATED: report-function-parity -->

## Notes on tricky function families

Some function families have especially subtle Excel behavior and deserve dedicated documentation.
For example, odd-coupon bond functions (ODDF\*/ODDL\*) involve Excel-specific conventions around
day-count bases, coupon schedules, accrued interest, and yield solver behavior. See:
[`docs/financial-odd-coupon-bonds.md`](financial-odd-coupon-bonds.md).

## Dependency threshold: 65,536 → “full recalc mode”

Excel has a behavior/implementation detail where, beyond a **65,536 dependency threshold**, it may
switch away from incremental dependency-driven recalculation into a “full recalc” strategy.

In this repo the requirement is tracked here:

- [`instructions/core-engine.md`](../instructions/core-engine.md) (see “Dependency threshold: 65,536
  before full recalc mode”)

Relevant implementation areas (where this is/will be enforced):

- Dependency graph enforcement (limit + fallback):
  [`crates/formula-engine/src/graph/dependency_graph.rs`](../crates/formula-engine/src/graph/dependency_graph.rs)
  - `DependencyGraph::DEFAULT_DIRTY_MARK_LIMIT` (65,536)
  - `DependencyGraph::mark_dirty` falls back to “full recalc” by marking all formula cells dirty when
    propagation exceeds the limit
- Regression tests:
  [`crates/formula-engine/tests/graph_dirty_mark_limit.rs`](../crates/formula-engine/tests/graph_dirty_mark_limit.rs)
  (plus general scaling tests like
  [`crates/formula-engine/tests/graph_stress.rs`](../crates/formula-engine/tests/graph_stress.rs))

Note: The parity report above focuses on **function coverage**. The dependency threshold is a
separate (performance/behavior) parity requirement.
