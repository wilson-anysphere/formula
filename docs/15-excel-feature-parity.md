# Excel Feature Parity (code-driven)

This document replaces the old hand-maintained “✅/⬜ per function” checklists, which quickly drift
and become misleading.

Instead, **function parity is reported from authoritative, versioned data sources in the repo** and
validated by tests.

## Sources of truth

### Implemented function catalog (authoritative)

The source of truth for “what the engine currently implements” is:

- **[`shared/functionCatalog.json`](../shared/functionCatalog.json)** (and
  [`shared/functionCatalog.mjs`](../shared/functionCatalog.mjs))
  - Kept in sync with the Rust registry by
    **[`crates/formula-engine/tests/function_catalog_sync.rs`](../crates/formula-engine/tests/function_catalog_sync.rs)**.
  - Used by downstream tooling (JS/TS, docs, scripts) without having to compile Rust.

Current implemented function count (from `shared/functionCatalog.json`): **319**

> This number will change over time. Run the parity report script below to get the current values
> in your checkout.

### Excel “function universe” references (informational)

Excel has multiple naming/encoding systems and many “functions” that aren’t worksheet functions
(legacy XLM macro functions, reserved slots, etc). The following sources are useful for context and
coverage tracking:

- **BIFF/XLSB built-in function table (FTAB)**:
  [`crates/formula-biff/src/ftab.rs`](../crates/formula-biff/src/ftab.rs)
  - Lists BIFF `iftab` ids `0..=484` and their canonical names.
  - **Does not** include many modern `_xlfn.` functions (which are commonly encoded as “user-defined”
    calls in BIFF).
- **Excel oracle corpus**:
  [`tests/compatibility/excel-oracle/cases.json`](../tests/compatibility/excel-oracle/cases.json)
  - A curated set of formulas + expected outputs, used to validate behavior against Excel.
  - Good for measuring “real-world” coverage, but not a comprehensive universe of all functions.

## Parity summary report (script)

To avoid baking volatile numbers into docs, use the repo script:

```bash
node tools/parity/report_functions.mjs
```

(Source: [`tools/parity/report_functions.mjs`](../tools/parity/report_functions.mjs))

To print the list of FTAB names that don’t exist in the engine catalog yet:

```bash
node tools/parity/report_functions.mjs --list-missing
```

The report prints:

- Implemented function count (from `shared/functionCatalog.json`)
- FTAB non-empty name count (from `crates/formula-biff/src/ftab.rs`)
- Approximate count of FTAB names missing from the catalog

Note: FTAB includes many legacy XLM/macro-only functions and reserved slots; the “missing from
catalog” number is **not** a direct “worksheet function parity” percentage.

### Example output

```text
Excel function parity (code-driven)
Implemented functions (shared/functionCatalog.json): <count>
BIFF FTAB function names (non-empty): <count>
FTAB names missing from engine catalog (approx): <count>
```

## Notes on tricky function families

Some function families have especially subtle Excel behavior and deserve dedicated documentation.
For example, odd-coupon bond functions (ODDF\*/ODDL\*) involve Excel-specific conventions around
day-count bases, coupon schedules, accrued interest, and yield solver behavior. See:
[`docs/financial-odd-coupon-bonds.md`](financial-odd-coupon-bonds.md).

## Updating / regenerating the function catalog

When adding or changing built-in functions in Rust, regenerate the catalog with:

```bash
node scripts/generate-function-catalog.js
```

(Source: [`scripts/generate-function-catalog.js`](../scripts/generate-function-catalog.js))

This writes:

- `shared/functionCatalog.json`
- `shared/functionCatalog.mjs`

Why it’s committed:

- It’s a **stable interface** for non-Rust tooling (autocomplete, docs, parity scripts, etc).
- It makes diffs reviewable (catalog is sorted and normalized).
- CI enforces correctness via `crates/formula-engine/tests/function_catalog_sync.rs`.

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

Note: The parity report script above focuses on **function coverage**. The dependency threshold is a
separate (performance/behavior) parity requirement.
