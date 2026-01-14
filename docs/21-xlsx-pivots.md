# XLSX Pivot Compatibility (PivotTables, PivotCaches, Slicers, Timelines)

This doc tracks **what `formula-xlsx` currently supports** for Excel pivots and related UX parts,
what we **intentionally preserve without interpreting**, and the **known gaps / roadmap**.

For the intended *cross-crate* ownership boundaries and end-to-end refresh flows (model schema vs
compute vs XLSX import/export), see
[ADR-0005: PivotTables ownership and data flow across crates](./adr/ADR-0005-pivot-tables-ownership-and-data-flow.md).

This document is intentionally scoped to **OpenXML I/O + round-trip fidelity**. Pivot *computation*
lives elsewhere:

- Worksheet/range pivots: `formula-engine`
- Data Model pivots (measures/relationships): `formula-dax`

Note: `formula-xlsx` also contains a best-effort OpenXML → engine bridge (for in-app computation),
but that is a convenience layer, not the canonical schema boundary (see ADR-0005).
The bridge lives at `crates/formula-xlsx/src/pivots/engine_bridge.rs` and is a migration target
(OpenXML → `formula-model` schema, then model → compute), not a long-term architecture decision.

The goal is team clarity: pivots are a large subsystem with many interdependent OPC parts, and
regressions are easy to introduce when changing ZIP/relationship handling.

---

## Scope and definitions

In this doc:

- **Pivot tables** live under `xl/pivotTables/pivotTable*.xml` and are attached to worksheets via
  `xl/worksheets/_rels/sheetN.xml.rels` relationships (`…/pivotTable`) and/or `<pivotTables>` blocks
  in `xl/worksheets/sheetN.xml`.
- **Pivot caches** (OpenXML PivotCaches, per ADR-0005 terminology) typically live under:
  - `xl/pivotCache/pivotCacheDefinition*.xml`
  - `xl/pivotCache/pivotCacheRecords*.xml`
  and are wired from `xl/workbook.xml` via `<pivotCaches>` + `xl/_rels/workbook.xml.rels`
  (`…/pivotCacheDefinition`). Cache definitions wire to cache records via the cache definition’s
  `.rels` (`…/pivotCacheRecords`).
- **Slicers** live under `xl/slicers/slicer*.xml` with cache definitions under
  `xl/slicerCaches/slicerCache*.xml`.
- **Timelines** live under `xl/timelines/timeline*.xml` with cache definitions under
  `xl/timelineCaches/timelineCacheDefinition*.xml`.

---

## What we currently support

### 1) Round-trip preservation of pivot-related parts

We preserve pivot-related parts **byte-for-byte** during round-trip (at the OPC part payload level):

- Pivot tables: `xl/pivotTables/**`
- Pivot caches: `xl/pivotCache/**`
- Slicers: `xl/slicers/**`, `xl/slicerCaches/**`
- Timelines: `xl/timelines/**`, `xl/timelineCaches/**`

This is exercised by integration tests that compare the original **part bytes** (e.g.
`XlsxPackage::part("xl/pivotTables/pivotTable1.xml")`) with the round-tripped part bytes:

- `XlsxDocument` round-trips (load → save): see `crates/formula-xlsx/tests/roundtrip_preserves_pivot_tables.rs`.
- `XlsxPackage` round-trips (`write_to_bytes`): see `crates/formula-xlsx/tests/roundtrip_preserves_slicers_and_pivot_charts.rs`.

For “regenerate the workbook XML from a model” pipelines (where parts would otherwise be dropped),
we also support extracting and re-applying the minimal pivot subgraph:

- `XlsxPackage::preserve_pivot_parts()` / `apply_preserved_pivot_parts()`
- Streaming extraction variant: `pivots::preserve::preserve_pivot_parts_from_reader`

These preserve and re-attach:

- `xl/workbook.xml` `<pivotCaches>` subtree + required relationships
- `xl/worksheets/sheetN.xml` `<pivotTables>` subtree + required relationships
- Required `[Content_Types].xml` overrides for the preserved parts

### 2) Pivot cache definition parsing (source + cacheField metadata)

We parse `pivotCacheDefinition*.xml` into `PivotCacheDefinition`:

- Cache metadata:
  - `pivotCacheDefinition@recordCount`
  - `pivotCacheDefinition@refreshOnLoad`
  - `pivotCacheDefinition@createdVersion`, `@refreshedVersion`
- Cache source:
  - `cacheSource@type` (`worksheet`, `external`, `consolidation`, `scenario`, or `Unknown`)
  - `cacheSource@connectionId` (for external sources)
  - `worksheetSource@sheet` + `worksheetSource@ref`
    - Best-effort normalization for non-standard producers that encode `Sheet1!A1:C5` in `ref`
      instead of using `sheet="..."`.
- `cacheField` attributes (best-effort, tolerant to namespaces/unknown tags):
  - `name`, `caption`, `propertyName`, `numFmtId`
  - `databaseField`, `serverField`, `uniqueList`
  - `formula`, `sqlType`, `hierarchy`, `level`, `mappingCount`

Note: when a PivotCache is backed by the **Data Model / Power Pivot**, Excel may encode field
captions using DAX-like strings (and multiple equivalent spellings), for example:
  
- `Table[Column]` vs `'Table Name'[Column]` (quoted table identifiers)
- `]` escaped as `]]` inside `[...]` identifiers
- measures sometimes stored as `Total Sales` (no brackets) or `[Total Sales]`

`formula-xlsx` preserves these caption strings as-is. The compute layer resolves multiple encodings
when binding structured `PivotFieldRef` values to cache fields.

Implementation: `crates/formula-xlsx/src/pivots/cache_definition.rs`.

### 3) Pivot cache record streaming parser (`m/n/s/b/e/d/x`)

We provide a **streaming** parser for `pivotCacheRecords*.xml` that avoids building a full DOM:

- API: `PivotCacheRecordsReader` (`XlsxPackage::pivot_cache_records` and `XlsxDocument::pivot_cache_records`)
- Supported record item tags:
  - `<m/>` → `PivotCacheValue::Missing`
  - `<n v="..."/>` → `Number(f64)`
  - `<s v="..."/>` → `String(String)`
  - `<b v="0|1"/>` → `Bool(bool)`
  - `<e v="..."/>` → `Error(String)`
  - `<d v="..."/>` → `DateTime(String)` (keeps raw text; Excel commonly stores ISO strings here)
  - `<x v="..."/>` → `Index(u32)` (shared-item index; decoding is a known gap below)

The parser is namespace-insensitive and is tolerant of producers that emit `<n><v>42</v></n>` style
wrappers (in addition to self-closing `v="..."` attributes).

Implementation: `crates/formula-xlsx/src/pivots/cache_records.rs`.

### 4) Pivot graph resolver (worksheet → pivot table → cache definition/records)

We resolve the relationship graph between:

- Worksheets (`xl/worksheets/sheetN.xml`)
- Pivot tables (`xl/pivotTables/pivotTableN.xml`)
- Pivot caches (`xl/pivotCache/pivotCacheDefinition*.xml` → `.rels` → `pivotCacheRecords*.xml`)

API: `XlsxPackage::pivot_graph()` returns `XlsxPivotGraph` with one `PivotTableInstance` per pivot
table.

Resolver behavior is intentionally tolerant of malformed/incomplete workbooks:

- If `workbook.xml` omits `<pivotCaches>` or workbook `.rels` entries are missing, we fall back to
  the common `pivotCacheDefinition{cacheId}.xml` / `pivotCacheRecords{cacheId}.xml` naming pattern.
- Pivot tables are still returned even when their sheet/caches cannot be resolved.

Implementation: `crates/formula-xlsx/src/pivots/graph.rs`.

### 5) Pivot chart binding parser

Excel stores pivot charts as normal chart parts (`xl/charts/chartN.xml`) and binds them via
`<c:pivotSource name="..." r:id="..."/>`.

We scan chart parts for `<pivotSource>` and resolve the `r:id` to a target part:

- API: `XlsxPackage::pivot_chart_parts() -> Vec<PivotChartPart>`
- Extracted fields:
  - `PivotChartPart.part_name` (`xl/charts/chartN.xml`)
  - `pivot_source_name` (the user-visible pivot source name)
  - `pivot_source_part` (resolved relationship target; typically a pivot table part)

Implementation: `crates/formula-xlsx/src/pivots/pivot_charts.rs`.

### 6) Slicer/timeline discovery + selection parsing

We parse slicer/timeline parts for UX wiring and (best-effort) selection state:

- API: `XlsxPackage::pivot_slicer_parts() -> PivotSlicerParts`

For **slicers** we capture:

- Part metadata: `xl/slicers/slicer*.xml` (`name`, `uid`)
- Cache binding: resolve relationship to `xl/slicerCaches/slicerCache*.xml`
- Cache metadata: `slicerCache@name`, `slicerCache@sourceName`
- Connected pivot tables: resolve slicer cache relationships to pivot table parts
- Drawing placement: discover which `xl/drawings/drawingN.xml` reference the slicer via drawing `.rels`
- Selection state: parse `slicerCacheItem` keys + selected flags
  - When Excel does not persist explicit selection state, slicers behave as “All selected”.
    We represent that as `selected_items: None`.

For **timelines** we capture:

- Part metadata: `xl/timelines/timeline*.xml` (`name`, `uid`)
- Cache binding: resolve relationship to `xl/timelineCaches/timelineCacheDefinition*.xml`
- Cache metadata: `timelineCacheDefinition@name`, `@sourceName`, `@baseField`, `@level`
- Connected pivot tables: resolve timeline cache relationships to pivot table parts
- Drawing placement: discover which `xl/drawings/drawingN.xml` reference the timeline via drawing `.rels`
- Selection state: best-effort extraction of `start`/`end` dates from either the timeline part or
  the timeline cache part.
  - Supports ISO `YYYY-MM-DD...`, `YYYYMMDD`, and numeric “Excel serial” encodings (see gaps below).

Implementation: `crates/formula-xlsx/src/pivots/slicers/mod.rs`.

### 7) Pivot cache refresh-from-worksheet

We can refresh a pivot cache definition/records from its `worksheetSource`:

- API: `XlsxPackage::refresh_pivot_cache_from_worksheet(cache_definition_part)`

Behavior (intentionally conservative):

- Reads `worksheetSource@sheet` + `@ref` from the cache definition.
- Loads cell values from the worksheet range (including shared strings).
  - Shared strings are discovered via `workbook.xml.rels` (and fall back to `xl/sharedStrings.xml`).
- Updates:
  - `pivotCacheDefinition@recordCount`
  - `<cacheFields>` list (names from the header row of the worksheet source range)
  - `pivotCacheRecords@count` + the `<r>` record list
- Preserves unrelated XML nodes/attributes where possible.

Implementation: `crates/formula-xlsx/src/pivots/refresh.rs`.

---

## Known gaps / planned work

These are the current “sharp edges” where we either don’t parse enough to interpret the file, or we
intentionally ignore fidelity-sensitive options:

1) **`sharedItems` decoding for `<x>` cache record values (partial)**
   - We parse `<x v="..."/>` as `PivotCacheValue::Index(u32)` and parse
      `pivotCacheDefinition*.xml` `cacheField/sharedItems` tables into `PivotCacheField.shared_items`.
   - Callers can resolve indices → actual values via `PivotCacheDefinition::resolve_record_value`
     (the pivot → engine bridge does this when building the source table).
   - We also expose `PivotCacheDefinition::resolve_shared_item(field_idx, index) -> ScalarValue`
     for resolving slicer/timeline item identifiers stored as shared-item indices.
   - The pivot → engine bridge includes helpers to translate slicer selections that use `x` indices
     into typed pivot-engine filters (see `pivots::engine_bridge::*pivot_slicer_parts*`).

2) **Mapping slicers/timelines to specific cache fields**
    - We discover slicers/timelines and the pivot tables they are connected to, but we do not yet
      join that information with pivot cache field metadata (e.g. `baseField` → cache field name) to
      produce a stable “this slicer filters field X” mapping.
    - As a best-effort fallback, the pivot → engine bridge can infer the cache field for some
      slicers by matching slicer item keys against the pivot cache’s unique values. This is
      heuristic and may fail for ambiguous fields or typed (non-text) slicers.

3) **Pivot table sort/manual ordering (partial)**
   - We parse `pivotField@sortType` and best-effort manual item ordering (including shared-item
     indices when `sharedItems` metadata is available) and map it into the pivot-engine config.
   - Remaining work: support the full range of Excel sort options (e.g. sorting by a value field),
     and ensure round-trip fidelity for all sort-related XML.

4) **`showDataAs` fidelity (partial)**
    - We map `dataField@showDataAs` into pivot-engine `ValueField.show_as`, and also map
      `baseField`/`baseItem` into engine `base_field`/`base_item` when possible.
    - The pivot engine now implements all `ShowAsType` variants currently represented in the
      `formula-model` schema (including base-item show-as variants like `percentOf` and
      `percentDifferenceFrom`).
    - Remaining work: expand the schema + bridge to cover additional Excel `showDataAs` variants not
      currently modeled (those will currently be treated as normal values).

---

## Roadmap (high-level)

This is the rough sequencing we expect to follow for pivot-related fidelity work:

1) **Pivot cache value decoding**
   - Parse `cacheField/sharedItems` from `pivotCacheDefinition*.xml`. (done)
   - Decode `<x>` indices in `pivotCacheRecords*.xml` into typed values. (done for pivot-engine
     source conversion; slicer/timeline selection resolution is supported via `resolve_shared_item`)

2) **Slicers/timelines → cache field mapping**
   - Join `pivot_graph()` results with slicer/timeline cache metadata (`baseField`, `sourceName`,
     relationships) to produce a stable mapping from a slicer/timeline to a specific cache field
     (name + index).

3) **Timeline date system correctness**
   - Respect workbook `date1904` when interpreting timeline serial values.

4) **Pivot table display fidelity**
    - Parse and preserve sort/manual ordering.
    - `showDataAs` mapping is wired into the engine config and the current set of modeled
      transformations are implemented in `formula-engine`; remaining work is to extend the modeled
      set of Excel variants and preserve/round-trip the corresponding XML without normalization.

5) **Authoring/editing pivots**
   - Move beyond “preserve and parse” into “modify and write”: controlled updates to
     `pivotTableDefinition` and cache parts while keeping relationship IDs stable and round-trip-safe.

---

## Testing guidance (fixtures)

The following fixtures cover different parts of the pivot subsystem:

| Fixture | Path | What it covers |
|---|---|---|
| `pivot-fixture.xlsx` | `crates/formula-xlsx/tests/fixtures/pivot-fixture.xlsx` (mirrored in the xlsx-diff corpus at `fixtures/xlsx/pivots/pivot-fixture.xlsx`) | Minimal pivot table + cache definition/records; round-trip preservation; cache definition parsing; cache record streaming parser; pivot → engine bridge smoke test. |
| `pivot-full.xlsx` | `crates/formula-xlsx/tests/fixtures/pivot-full.xlsx` | Full relationship chain (`workbook.xml` `<pivotCaches>` + worksheet `.rels` + cache definition `.rels`); `pivot_graph()` end-to-end; preserve/apply of workbook + sheet pivot subtrees. |
| `pivot_slicers_and_chart.xlsx` | `crates/formula-xlsx/tests/fixtures/pivot_slicers_and_chart.xlsx` | Slicer + timeline discovery; drawing placement via drawing `.rels`; pivot chart binding (`<c:pivotSource>`); round-trip preservation of slicer/timeline/chart parts. |
| `slicer-selection.xlsx` | `crates/formula-xlsx/tests/fixtures/slicer-selection.xlsx` | Explicit slicer selection parsing (subset selected vs “all selected”). |
| `timeline-selection.xlsx` | `crates/formula-xlsx/tests/fixtures/timeline-selection.xlsx` | Timeline selection parsing (`start`/`end` extraction + normalization). |

If you add new pivot parsing/writing behavior, prefer extending these fixtures (or adding a new
targeted fixture) and asserting on the parsed structs **and** on byte-for-byte round-trip (OPC part
payload bytes) when applicable.
