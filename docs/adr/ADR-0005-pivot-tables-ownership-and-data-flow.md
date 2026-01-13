# ADR-0005: PivotTables ownership and data flow across crates

- **Status:** Accepted
- **Date:** 2026-01-13

## Context

PivotTables cut across multiple layers of the stack:

- workbook persistence / IPC-friendly schema,
- worksheet/range-based pivot computation,
- Data Model (PowerPivot-style) pivot computation,
- XLSX (OpenXML) import/export and round-trip fidelity,
- JS/WASM API surface.

During development we accumulated **multiple partial pivot implementations**:

- `formula-model::pivots::{PivotTable, PivotManager}` implements a small in-memory pivot + slicer/timeline filter loop.
- `formula-engine::pivot` implements a newer cache-backed pivot engine (row/col/value fields, totals, “show as”).
- `formula-xlsx::pivots::engine_bridge` converts parsed OpenXML pivots *directly* into `formula-engine::pivot` types.

Without an explicit ownership boundary, future work tends to:

- re-introduce “just one more pivot engine” in whichever crate is being touched,
- bake OpenXML-specific details into the model layer,
- or couple worksheet pivots and Data Model pivots into one monolith.

This ADR defines **one canonical pivot schema** and the **single intended data flow** so we don’t regress into duplicate implementations.

### Terminology (to avoid “pivot cache” confusion)

- **OpenXML PivotCache**: the XLSX parts under `xl/pivotCache/` (definition + records) that Excel uses
  to persist pivot state in a file. These are owned by `formula-xlsx` (import/export + round-trip
  fidelity) and are *not* the canonical in-app computation cache.
- **Engine PivotCache**: the runtime cache used to compute worksheet/range pivots (e.g.
  `formula_engine::pivot::PivotCache`). This is owned by `formula-engine` and is treated as
  **runtime-only** (rebuildable from the worksheet source).
- **Data Model pivot**: a pivot whose source is the PowerPivot/Data Model; computed by `formula-dax`
  (group-by + measures under a filter context).

## Decision

### 1) Crate ownership (single source of truth)

Quick summary:

| Crate | Owns | Does **not** own |
|---|---|---|
| `formula-model` | Canonical pivot *definitions* (PivotTable/PivotConfig/PivotSource), slicer/timeline state, serialization/IPC schema | Pivot computation, cache building, OpenXML parsing |
| `formula-engine` | Worksheet/range pivot computation, runtime cache building/invalidation, writing results to sheet | Persisted pivot schema, Data Model measure/relationship evaluation, OpenXML I/O |
| `formula-dax` | Data Model pivot computation (group-by, measures, relationships) | Worksheets/ranges, OpenXML I/O, pivot rendering into sheet cells |
| `formula-xlsx` | OpenXML pivot cache/table + slicer/timeline parts import/export + round-trip preservation, bridge to model schema | Pivot computation engine |
| `formula-wasm` | JS API surface to mutate model + call engine refresh/compute | Pivot computation |

#### `formula-model` — canonical persisted/IPC schema (NO computation)

Owns the **serialization-friendly, stable workbook schema** for pivots and related UX objects:

- Pivot identities and definitions:
  - `PivotTableId`
  - `PivotTable` definition record (name, destination, config, source reference)
  - `PivotSource` (range/table vs DataModel)
- Pivot configuration:
  - `PivotConfig` (row/col/value/filter fields, layout, totals)
  - field/value specs (`PivotField`, `ValueField`, `FilterField`)
  - calculated fields/items *as definitions* (raw formula strings), not evaluated values
- Slicers/timelines as workbook objects:
  - ids, display metadata, selection state, and connections to pivots
  - a model-level filter representation that can be applied by either compute backend

Non-goal: `formula-model` MUST NOT contain a pivot engine, cache builder, or worksheet cell evaluation logic.

#### `formula-engine` — worksheet/range pivot computation + cache building

Owns the runtime computation for pivots whose source is a worksheet/range/table:

- extracting a rectangular dataset from a worksheet range (headers + records),
- building and maintaining runtime `PivotCache` instances,
- computing pivot outputs for a given `PivotConfig` and filter state,
- writing the result grid into a destination range (or returning “cell writes” to the caller),
- cache invalidation (source range edits, refresh requests, slicer/timeline selection changes).

`formula-engine` is the orchestrator for “refresh pivot / refresh all” at runtime.

#### `formula-dax` — Data Model pivot computation

Owns pivot computation where the source is the **Data Model**:

- DAX measure evaluation,
- relationships (RELATED / star-schema behavior),
- columnar group-by fast paths,
- evaluation under a filter context.

`formula-dax` does not know about worksheets, ranges, OpenXML, or rendering a 2D sheet grid.

#### `formula-xlsx` — OpenXML pivot import/export + bridging to model schema

Owns XLSX/OpenXML concerns:

- parsing and writing:
  - `xl/pivotTables/*.xml`
  - `xl/pivotCache/pivotCacheDefinition*.xml` + `pivotCacheRecords*.xml`
  - slicer/timeline parts (`xl/slicers/*`, `xl/slicerCaches/*`, `xl/timelines/*`, ...)
- preserving unknown/unmodeled XML for round-trip fidelity,
- bridging OpenXML pivot parts into the **model schema** (`formula-model`),
- emitting OpenXML parts from the model schema (plus any preserved XML stubs).

Non-goal: `formula-xlsx` MUST NOT become an alternative pivot compute engine. Its job is file I/O and fidelity.

#### `formula-wasm` — JS API surface

Owns the JavaScript-facing API:

- create/update pivot definitions in `formula-model`,
- invoke `formula-engine` refresh operations,
- expose slicer/timeline selection manipulation,
- return computed results (either via engine-applied cell writes or as explicit result payloads).

Non-goal: `formula-wasm` MUST NOT implement pivot computation in JS or WASM glue code.

---

### 2) Canonical data flows

The flows below describe the intended end-to-end ownership boundaries. The exact API names may change, but the directionality should not.

#### A) Create pivot (range/table source)

1. **UI / host** collects:
   - source range/table,
   - destination (top-left cell),
   - pivot layout (row/col/value/filter fields).
2. **`formula-wasm`**:
   - constructs a `formula-model` `PivotTable` definition with a stable `PivotTableId`,
   - writes it into the workbook model.
3. **`formula-engine`**:
   - reads the source range values,
   - builds or reuses a runtime cache keyed by `PivotSource`,
   - computes the pivot result grid,
   - applies the output to the destination range (or returns cell writes).

Key property: the workbook model stores the **definition**, not a computed cache.

#### Identity note (PivotTable IDs)

- The **canonical pivot identity** is the model-layer `PivotTableId` (UUID) stored in `formula-model`.
- `formula-engine::pivot::PivotTable.id` is currently a runtime-only string (`pivot-1`, `pivot-2`, …) used for
  convenience in tests and transient compute objects; it is **not stable across sessions** and must not be persisted.
- XLSX pivot parts are identified by *part names/relationships* (`xl/pivotTables/pivotTableN.xml`, rel ids, etc.) and
  may include user-visible names like `"PivotTable1"`. Import should map these to stable `PivotTableId`s in the model.

#### B) Create pivot (Data Model source)

1. **UI / host** chooses:
   - a base table (fact table) and group-by columns,
   - measure list (DAX expressions or references to stored measures),
   - destination range.
2. **`formula-wasm`** writes a `PivotTable` definition with `PivotSource::DataModel`.
3. **`formula-engine`** builds a `formula-dax` request:
   - group-by columns,
   - measures,
   - filter context derived from slicers/timelines.
4. **`formula-dax`** executes the grouped query and returns a tabular result.
5. **`formula-engine`** converts the result into sheet cell writes and updates the destination range.

#### C) Refresh pivot (single)

Refresh can be triggered by:

- explicit user action (“Refresh”),
- source edits (range/table contents change),
- slicer/timeline selection changes,
- Data Model changes (table refresh, relationship changes, measure edits).

1. **`formula-engine`** looks up the pivot definition in `formula-model` by `PivotTableId`.
2. It resolves the pivot’s source:
   - **Range/table pivot:** (re)build cache from worksheet data.
   - **Data Model pivot:** build a DAX filter context.
3. It resolves filters:
   - static pivot `filter_fields`,
   - plus dynamic filters from connected slicers/timelines (see mapping below).
4. It recomputes and writes output.

#### D) Refresh all

“Refresh All” is conceptually a batch of refresh-pivot operations:

1. **Orchestrator (host/UI)** selects the target set:
   - all pivots in the workbook (and optionally Power Query / Data Model refresh first).
2. **`formula-engine`** performs refreshes with two important invariants:
   - **cache sharing:** multiple pivots referencing the same `PivotSource` should share a rebuilt cache where possible,
   - **deterministic writes:** pivot output writes must be applied in a deterministic order to avoid nondeterministic overlap behavior.

#### E) XLSX import

1. **`formula-xlsx`** parses OpenXML pivot parts and slicer/timeline parts.
2. It creates/updates the workbook’s `formula-model` pivot schema:
   - pivot definitions (including source references),
   - slicer/timeline objects and their connections,
   - selection state when present.
3. It preserves the original OpenXML pivot parts for round-trip fidelity.
4. **Runtime compute choice:**
   - The engine may compute pivots from worksheet data (preferred for correctness when the sheet differs from cached OpenXML records),
   - while `formula-xlsx` keeps the original pivot parts available for export.

#### F) XLSX export

1. **`formula-xlsx`** serializes workbook state (cells, styles, etc.) and pivot model state.
2. For pivot parts:
   - **preferred:** emit OpenXML pivot parts from the model schema and refresh cache records from current worksheet data when requested/required,
   - **otherwise:** preserve/imported pivot parts when we cannot fully regenerate a compatible OpenXML representation yet.

Pragmatic note: Excel is tolerant of many pivot-cache mismatches, but users are not. When we can refresh caches from worksheet data, we should.

---

### 3) Slicers/timelines → pivot filters

Slicers and timelines are workbook-level UX objects that affect one or more pivots. They must map into a common filter representation that both backends can apply.

#### Slicer mapping (discrete selections)

- **Slicer selection = All**  
  ⇒ no filter clause (equivalent to “allow all values”).

- **Slicer selection = Items(Set\<…\>)**  
  ⇒ an “IN” filter on the slicer’s bound field/column.

Application:

- **Range/table pivots (`formula-engine`)**: filter rows in the pivot cache by checking whether the row’s field value is in the allowed set.
- **Data Model pivots (`formula-dax`)**: translate to filter-context constraints on the referenced column; relationship propagation is handled by the model/engine.

OpenXML note:

- XLSX slicer caches may store item identifiers as either display strings or indices (e.g. `<slicerCacheItem x="...">`)
  into a pivot cache shared-items table.
- Resolving these identifiers into **typed values** (number/date/text/bool) is owned by `formula-xlsx` (because it requires
  OpenXML pivot cache parsing). When a typed mapping is unavailable, best-effort text matching is acceptable but should be
  treated as a fidelity gap.

#### Timeline mapping (date range selection)

- Timeline selection is an inclusive range: `[start, end]` (either endpoint may be missing).

Application:

- **Range/table pivots (`formula-engine`)**: filter cache rows where the bound field is a date and satisfies the range.
- **Data Model pivots (`formula-dax`)**: translate to filter-context constraints of the form:
  - `Column >= start` (if start is set)
  - `Column <= end` (if end is set)

Important: timelines are not representable as a pure “allowed set” without exploding into many individual dates; the model-level filter representation must preserve the range intent.

Implementation note:

- `formula-dax::FilterContext` is currently expressed primarily as **allowed-value sets** per column.
  Until the DAX layer supports native range predicates in its filter context, timeline ranges may need to be:
  - materialized into an allowed set (potentially large), or
  - applied by extending the filter representation to carry `>=`/`<=` predicates.

## Migration note (current state + planned consolidation)

### Current state (today)

- `formula-model::pivots::PivotManager` (and related `DataTable`, `PivotTable::refresh`, slicer/timeline filters) is a **legacy MVP** used by unit tests and early prototypes.
  - It lives in the model crate but performs computation, which violates the desired crate boundary.
- `formula-engine::pivot` contains the newer cache-backed pivot engine and defines its own pivot config/value types.
  - Note: types like `formula_engine::pivot::PivotTable` and `formula_engine::pivot::PivotConfig` are **compute-layer**
    constructs today. They are convenient for tests and bridges (e.g. XLSX → engine), but they are not yet the canonical
    persisted workbook schema.
- `formula-xlsx` currently contains a direct bridge from OpenXML pivots → `formula-engine::pivot` types (`crates/formula-xlsx/src/pivots/engine_bridge.rs`).
- Slicers/timelines:
  - `formula-xlsx` can parse slicer/timeline parts and convert selection state into `formula_model::pivots::slicers::RowFilter`.
  - The **legacy** `formula-model::pivots::PivotManager` applies those filters when refreshing its in-memory pivot.
  - The **new** `formula-engine::pivot` path currently models filters as `PivotConfig.filter_fields` (allowed-value sets) and is not yet wired to consume workbook slicer/timeline state directly.

### Migration plan (what future work should converge to)

1. **Move/define the canonical pivot schema in `formula-model`** (PivotTable definition, PivotConfig, PivotSource, filter representation).
2. **Make `formula-engine` consume `formula-model` pivot schema** for computation, instead of defining parallel config/value types.
3. **Make `formula-xlsx` map OpenXML ↔ `formula-model` schema**, not OpenXML ↔ engine types.
4. **Deprecate and remove `formula-model::pivots::PivotManager`** once:
   - workbook schema contains pivot definitions,
   - `formula-engine` refresh is wired into the host,
   - slicer/timeline state is stored in the model and applied through the engine.

### Known limitations (explicitly not implemented yet)

These are present in the schema in some form but are not fully evaluated end-to-end:

- **Calculated fields / calculated items**:
  - Definitions exist (name + formula text),
  - but evaluation is not yet implemented in the new engine path.
- **Slicer/timeline integration in the new engine path**:
  - Selection parsing exists (`formula-xlsx`) and model-level filter types exist (`formula-model`),
  - but end-to-end “change slicer selection → engine refresh pivot output” is not yet unified under the `formula-engine::pivot` compute path.
- **Full OpenXML parity**:
  - many pivot/table style and layout knobs are preserved for round-trip,
  - but may be ignored by the compute engine.
- **Pivot charts**:
  - OpenXML parts can be preserved,
  - but chart data-binding + refresh is not yet fully wired through the new pivot engine.

## Consequences

- New pivot-related features MUST start by asking: “Is this schema, compute, or I/O?”
  - schema ⇒ `formula-model`,
  - range compute ⇒ `formula-engine`,
  - Data Model compute ⇒ `formula-dax`,
  - XLSX ⇒ `formula-xlsx`,
  - JS API ⇒ `formula-wasm`.
- Do not add new pivot engines in UI glue code or file I/O layers.
- Conversion helpers (OpenXML ↔ model, model ↔ compute requests) are expected and encouraged; duplicated compute implementations are not.

## Current implementation pointers

- Legacy model pivot engine (to be removed): `crates/formula-model/src/pivots/mod.rs`
- Model slicers/timelines types (selection state): `crates/formula-model/src/pivots/slicers.rs`
- Engine pivot compute + cache: `crates/formula-engine/src/pivot/mod.rs`
- OpenXML pivot parsing/preservation: `crates/formula-xlsx/src/pivots/*`
- OpenXML → engine bridge (migration target to “→ model”): `crates/formula-xlsx/src/pivots/engine_bridge.rs`
- OpenXML slicer/timeline parsing + mapping helpers: `crates/formula-xlsx/src/pivots/slicers/mod.rs`
- Data Model pivot computation: `crates/formula-dax/src/pivot.rs`
- WASM/JS API crate: `crates/formula-wasm/`
