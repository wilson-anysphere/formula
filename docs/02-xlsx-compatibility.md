# XLSX Compatibility Layer

## Overview

Perfect XLSX compatibility is the foundation of user trust. Users must be confident that their complex financial models, scientific calculators, and business-critical workbooks will load, calculate, and save without any loss of fidelity.

## Related docs

- [adr/ADR-0005-pivot-tables-ownership-and-data-flow.md](./adr/ADR-0005-pivot-tables-ownership-and-data-flow.md) — PivotTables ownership + data flow across crates (schema vs compute vs XLSX IO)
- [encrypted-workbooks.md](./encrypted-workbooks.md) — overview + links for password-protected Excel files (OOXML `EncryptedPackage`, legacy `.xls` `FILEPASS`)
- [20-xlsx-rich-data.md](./20-xlsx-rich-data.md) — Excel `richData` / rich values (including “image in cell”; naming varies: `richValue*` vs `rdrichvalue*`)
- [20-images-in-cells.md](./20-images-in-cells.md) — Excel “Images in Cell” (`IMAGE()` / “Place in Cell”) packaging + schema notes
- [20-images-in-cells-richdata.md](./20-images-in-cells-richdata.md) — RichData (`richValue*` / `rdrichvalue*`) tables used by images-in-cells
- [xlsx-embedded-images-in-cells.md](./xlsx-embedded-images-in-cells.md) — confirmed “Place in Cell” chain (the `rdRichValue*` schema; the fixture used there is encoded as `t="e"`/`#VALUE!`, but other real Excel files use other cached-value encodings)
- [xlsx-comments.md](./xlsx-comments.md) — legacy notes + threaded comments + persons parts (relationships, parsing, preservation)
- [21-encrypted-workbooks.md](./21-encrypted-workbooks.md) — password-protected / encrypted Excel workbooks (OOXML `EncryptedPackage`, legacy `.xls` `FILEPASS`)
- [21-offcrypto.md](./21-offcrypto.md) — MS-OFFCRYPTO details for encrypted OOXML workbooks (`EncryptionInfo` + `EncryptedPackage`)
- [22-ooxml-encryption.md](./22-ooxml-encryption.md) — OOXML password decryption reference for MS-OFFCRYPTO Agile 4.4 (HMAC target bytes, IV/salt gotchas, error semantics)
- [21-xlsx-pivots.md](./21-xlsx-pivots.md) — pivot tables/caches/slicers/timelines compatibility notes + roadmap

---

## Compatibility checklist (L1 Read / L4 Round-trip)

This repo’s File I/O workstream uses the following shorthand (see [`instructions/file-io.md`](../instructions/file-io.md)):

- **L1 (Read)**: the workbook opens and all data is visible/usable.
- **L4 (Round-trip)**: we can save and reopen in Excel without unintended diffs or fidelity loss.

| Feature | L1 impact | L4 impact | Preservation / patching summary |
|---|---:|---:|---|
| Data validations (`<dataValidations>`) | Low (UI affordance) | High | Preserve the `<dataValidations>` subtree byte-for-byte on round-trip; ensure schema-ordering when inserting/replacing nearby sections (e.g. `<mergeCells>` must come before `<dataValidations>`). |
| Row/column default styles (`row/@s`, `col/@style`) | Medium (formatting/render) | High | The reader imports `row/@s`+`row/@customFormat` and `col/@style`+`col/@customFormat` into `Worksheet.row_properties/col_properties[*].style_id`. `write_workbook*` emits these defaults from the model; `formula_xlsx::write::write_to_vec` preserves them on round-trip and re-emits them when it has to regenerate `<cols>` or synthesize `<row>` elements. Remaining risk: `<cols>` regeneration is semantic (may drop unmodeled col attrs / collapse ranges), and truly-empty `<row>` placeholders with no modeled properties may still be dropped. |
| Rich text (`sharedStrings` runs + `inlineStr`) | High (display fidelity) | High | Parse rich runs for display; preserve raw `<si>` records and unchanged inline `<is>` subtrees to avoid reserializing formatting. |
| Sheet view state (`<sheetViews>` beyond zoom/freeze) | Medium (UI state) | Medium | Preserve the full `<sheetViews>` block when unchanged; if we need to update view state, we rewrite `<sheetViews>` as a minimal modeled block (zoom + frozen/split pane + topLeftCell + selection + gridlines/headings/zeros), which may drop other unmodeled view state. |
| Worksheet protection (`<sheetProtection>`) | Low (editability) | High | Preserve modern hashing attrs (`algorithmName`, `hashValue`, `saltValue`, `spinCount`) byte-for-byte when protection settings are unchanged; avoid rewriting `<sheetProtection>` unless the model edits it. |
| Tables (`xl/tables/table*.xml`) | Medium (structured refs/filters) | High | Preserve `xl/tables/*` parts + worksheet `<tableParts>` + relationships. Table `<autoFilter>` shares worksheet `<autoFilter>` semantics; preserve advanced criteria via `raw_xml`. |
| Typed date cells (`c/@t="d"`) | Medium (display fidelity) | High | Stored as ISO-8601 text in `<v>`; should behave as an Excel date serial for calc. Round-trip must preserve `t` and the original ISO string even if the in-memory model stores a plain string. |
| Comments (legacy notes + threaded) | Medium (collab/review) | High | Parse comment text/authors into the model; preserve all comment-related OPC parts byte-for-byte unless explicitly rewriting comment XML. |
| Outline metadata (`outlinePr`, row/col `outlineLevel`/`collapsed`/`hidden`) | Medium (grouping UX) | High | Outline can be parsed into `formula_model::Outline` via the `formula_xlsx::outline` helpers; on write, we preserve existing outline attrs when rows/cols are preserved, and we emit outline attrs for newly-written rows/cols based on `Worksheet.outline`. |
| OPC robustness (path normalization) | High (open more files) | High | Normalize part lookup for leading `/`, Windows `\` separators, ASCII case differences, and percent-decoding fallback. Relationship targets also strip `#fragment` / `?query`. `.rels` bytes are preserved unless we must explicitly repair/append relationships. |

---

## Feature preservation rules

This section documents the **current preservation contract** for specific XLSX features: what we parse into the workbook model vs what we keep around for lossless round-trip.

### Data validations (`<dataValidations>`)

Excel stores dropdown lists, custom validation formulas, and input/error prompts under `<dataValidations>` in each worksheet XML.

#### Preservation strategy

- **Parsed into model:** currently **not** parsed into `Worksheet.data_validations` (future work).
- **Preserved byte-for-byte:** the original `<dataValidations>` subtree is preserved on round-trip. (We do not currently generate validations from the model.)

#### Patch/write rules

- We treat `<dataValidations>` as **schema-ordered** relative to other worksheet sections. When inserting/replacing nearby blocks (e.g. `<mergeCells>`), we insert in the correct position **before** `<dataValidations>` to avoid Excel warnings.
- When we do not touch validations, we do not normalize/rewrite the block (attribute order, unknown extension nodes, etc.).

### Row/column default styles (`row/@s`, `col/@style`)

SpreadsheetML supports row/column-level default formatting:

- `row/@s` (+ `row/@customFormat`) applies a default `xf` style index to all cells in a row that do not specify `c/@s`.
- `col/@style` applies a default `xf` style index to all cells in a column range (`<col min="…" max="…">`).

#### Preservation strategy

- **Parsed into model:** the XLSX reader maps:
  - `row/@s` + `row/@customFormat` → `Worksheet.row_properties[*].style_id`
  - `col/@style` + `col/@customFormat` → `Worksheet.col_properties[*].style_id`
- **Preserved byte-for-byte:**
  - `row/@s` + `row/@customFormat` are preserved on round-trip because style-only rows are now represented in `Worksheet.row_properties` (so the sheet-data patcher does not treat them as droppable “empty rows”).
  - `col/@style` + `col/@customFormat` are preserved:
    - byte-for-byte when the existing `<cols>` section is preserved, and
    - semantically when `<cols>` is regenerated (the regenerated `<col>` ranges include `style` + `customFormat` when `Worksheet.col_properties[*].style_id` is set).

#### Patch/write rules

- Treat row/col default styles as **formatting-critical** round-trip metadata: avoid dropping rows/cols that exist solely to carry `row/@s` or `col/@style`.
- When patching existing worksheet XML, prefer preserving the original `<row>` / `<cols>` elements rather than regenerating them.
- When we *must* synthesize or regenerate these sections:
  - the workbook writer (`write_workbook*`) emits row/col default styles from `RowProperties.style_id` / `ColProperties.style_id`
  - the XlsxDocument round-trip writer (`formula_xlsx::write::write_to_vec`) emits them when generating `<row>` / `<cols>` content (note that regenerated `<cols>` may still drop **unmodeled** `<col>` attributes and range fragmentation).

### Rich text (shared strings `<r>` runs + `inlineStr`)

Excel can encode rich text in two places:

1. **`xl/sharedStrings.xml`**: `<si>` entries can contain a simple `<t>` or multiple `<r>` runs with formatting (`<rPr>`).
2. **Inline strings (`t="inlineStr"`)**: `<c t="inlineStr"><is>…</is></c>` can also contain `<r>` runs.

#### Preservation strategy

- **Parsed into model:** visible text plus run-level style for shared strings; inline strings are parsed at least to their visible text.
- **Preserved byte-for-byte:**
  - Existing shared-string `<si>` XML is preserved so we don’t reserialize rich runs we didn’t generate.
  - Inline `<is>` payloads are preserved when the cell is unchanged (so rich runs survive no-op saves).

#### Patch/write rules

- Prefer updating shared strings via an *append-only* editor (do not rewrite existing `<si>` records).
- Only re-emit an inline string’s `<is>` subtree when the cell value actually changes.

### Sheet view state (`<sheetViews>` beyond zoom/freeze)

`<sheetViews>` contains user interface state such as:

- selection (`<selection activeCell="…" sqref="…"/>`)
- gridlines/headings visibility (`showGridLines`, `showRowColHeaders`, etc.)
- split panes (freeze and non-freeze splits)
- zoom, view mode, and other sheet window flags

#### Preservation strategy

- **Parsed into model:** best-effort subset:
  - `sheetView/@zoomScale` → `Worksheet.zoom` + `Worksheet.view.zoom`
  - `sheetView/@showGridLines`, `@showHeadings`, `@showZeros` → `Worksheet.view.*`
  - panes:
    - frozen panes (`<pane state="frozen|frozenSplit" xSplit="…" ySplit="…">`) → `Worksheet.frozen_rows` / `Worksheet.frozen_cols` (+ `Worksheet.view.pane.frozen_*`)
    - split panes (`<pane … xSplit="…" ySplit="…">`) → `Worksheet.view.pane.x_split` / `y_split`
    - `pane/@topLeftCell` → `Worksheet.view.pane.top_left_cell`
  - selection (`<selection activeCell="…" sqref="…"/>`) → `Worksheet.view.selection`
- **Preserved byte-for-byte:** the full `<sheetViews>` subtree is preserved **only** when we do not need to update modeled view state. When we update, we replace `<sheetViews>` with a minimal modeled block (zoom + pane + selection + gridlines/headings/zeros), which may drop other view state and extension payloads.

#### Patch/write rules

- When `Worksheet.zoom` / frozen panes are unchanged, we keep the original `<sheetViews>` bytes unchanged.
- When modeled view state changes, we replace `<sheetViews>` with a minimal block and do not currently attempt to merge/preserve unmodeled state such as:
  - extra `<sheetView>` attributes we don’t model
  - multiple `<sheetView>` entries / non-zero `workbookViewId`
  - extension payloads (`<extLst>`)

### Worksheet protection (`<sheetProtection>`)

Worksheet protection is stored in `xl/worksheets/sheetN.xml` as a `<sheetProtection …/>` element. It
contains:

- a legacy `password="ABCD"` hash (16-bit, not cryptographically secure)
- allow-list booleans like `formatCells`, `insertRows`, `sort`, `autoFilter`, etc.
- inverted “protected” flags: `objects` and `scenarios`
- (newer Excel) modern hashing attributes like `algorithmName`, `hashValue`, `saltValue`,
  `spinCount`, etc.

#### Preservation strategy

- **Parsed into model:** a best-effort subset of allow-list flags plus the legacy `password` hash
  into `Worksheet.sheet_protection`. Unsupported attributes are ignored.
- **Preserved byte-for-byte:** modern hashing attributes (e.g. `algorithmName`, `hashValue`,
  `saltValue`, `spinCount`) must be preserved on round-trip even if the public model only exposes
  the legacy `password` hash.
  - In practice this means: if we aren’t changing protection, we should avoid rewriting
    `<sheetProtection>` (see the patch-in-place logic in `crates/formula-xlsx/src/write/mod.rs`).

#### Patch/write rules

- If protection settings change, the writer replaces the `<sheetProtection>` element using the
  modeled fields. This currently drops unmodeled attributes (including modern hashing attributes).
- When inserting/replacing `<sheetProtection>`, preserve schema ordering (Excel expects
  `sheetData`, then optional `sheetCalcPr`, then `sheetProtection`).

### Tables (`xl/tables/table*.xml`)

Excel “tables” (ListObjects) are stored in separate parts like `xl/tables/table1.xml` and linked
from the worksheet via `<tableParts>` plus a `.../relationships/table` relationship in
`xl/worksheets/_rels/sheetN.xml.rels`.

Table part parsing/writing lives in `crates/formula-xlsx/src/tables/xml.rs`.

#### Table `<autoFilter>` (same semantics as worksheet `<autoFilter>`)

A table part (`table.xml`, e.g. `xl/tables/table1.xml`) can contain an `<autoFilter>` element whose
schema/behavior matches the worksheet-level `<autoFilter>` (same
`filterColumn`/`filters`/`customFilters`/`dynamicFilter`/`sortState` vocabulary).

- Worksheet autoFilter parsing/writing lives in `crates/formula-xlsx/src/autofilter/*` and is applied
  during round-trip patching in `crates/formula-xlsx/src/write/mod.rs`.

For forward compatibility, any advanced filter criteria we don’t model (e.g. newer filter elements or
`extLst` payloads) should be preserved by storing the original XML fragments in the `raw_xml` fields
on the AutoFilter / FilterColumn structs and re-emitting them unchanged. Worksheet autoFilters
already do this via `crates/formula-xlsx/src/autofilter/*`; table autoFilters should follow the same
pattern when we extend `crates/formula-xlsx/src/tables/xml.rs` beyond the common subset.

### Comments (legacy notes + threaded comments)

#### Preservation strategy

- **Parsed into model:** comment anchors + visible text + authors (both legacy notes and threaded comments).
- **Preserved byte-for-byte:** comment-adjacent OPC parts (VML shapes, commentsExt, persons, and related `.rels`) unless the caller explicitly rewrites comment XML.

See [`docs/xlsx-comments.md`](./xlsx-comments.md) for the full part layout, relationship types, and preservation rules.

### Outline metadata (`outlinePr`, row/col `outlineLevel`/`collapsed`/`hidden`)

Outline/grouping state is split across:

- `<sheetPr><outlinePr …/></sheetPr>`
- row attributes: `row/@outlineLevel`, `row/@collapsed`, `row/@hidden`
- column attributes: `col/@outlineLevel`, `col/@collapsed`, `col/@hidden`

#### Preservation strategy

- **Parsed into model:** outline metadata is represented as `formula_model::Outline` (`Worksheet.outline`). It can be parsed on-demand via `formula_xlsx::outline::{read_outline_from_xlsx_bytes, read_outline_from_worksheet_xml}`.
- **Preserved byte-for-byte:** outline-related attributes remain byte-for-byte intact as long as we preserve the original `<row>` / `<cols>` elements. When we regenerate `<cols>` or synthesize new `<row>` elements, we only re-emit the outline attributes we model.

#### Patch/write rules

- When writing outline changes, update **only**:
  - `outlinePr` attributes
  - row/col outline attributes (`outlineLevel`, `collapsed`, `hidden`)
- Recompute derived hidden state (outline-hidden vs user-hidden) before writing.
- Keep schema order: if inserting a new `<cols>` section for outline columns, place it before `<sheetData>`.

### OPC robustness (relationship targets + part lookup)

Real-world XLSX producers sometimes emit relationship targets that diverge from canonical OPC part names.

#### Preservation strategy

- **Parsed into model:** none (this is lookup/IO plumbing).
- **Preserved byte-for-byte:** `.rels` parts and `[Content_Types].xml` are preserved unless we must explicitly repair/append entries.

#### Patch/write rules

- Relationship target normalization is *lookup-only*:
  - tolerate Windows path separators (`\`)
  - ignore URI fragments/queries (`#…`, `?…`) when mapping to ZIP entry names
  - fall back to case-insensitive and percent-decoded lookups (and tolerate a leading `/`) when a target does not resolve as-is
- Never rewrite a relationship `Target` string just to “normalize” it; preserve original strings unless the relationship itself is being created/repaired.

## XLSX File Format Structure

XLSX is a ZIP archive following Open Packaging Conventions (ECMA-376).

**Exception:** Excel “Encrypt with Password” / “Password to open” files are **not ZIP** archives even
when they use a `.xlsx`/`.xlsm` extension. Excel wraps the encrypted workbook in an **OLE/CFB**
container with `EncryptionInfo` + `EncryptedPackage` streams. See:

- [`docs/encrypted-workbooks.md`](./encrypted-workbooks.md) (overview + entrypoints)
- [`docs/21-encrypted-workbooks.md`](./21-encrypted-workbooks.md) (high-level background + terminology)
- [`docs/21-offcrypto.md`](./21-offcrypto.md) (developer-facing MS-OFFCRYPTO debugging + API usage)

```
workbook.xlsx (ZIP archive)
├── [Content_Types].xml          # MIME type declarations
├── _rels/
│   └── .rels                    # Package relationships
├── docProps/
│   ├── app.xml                  # Application properties
│   └── core.xml                 # Core properties (author, dates)
├── xl/
│   ├── workbook.xml             # Workbook structure, sheet refs
│   ├── styles.xml               # All cell formatting
│   ├── sharedStrings.xml        # Deduplicated text strings
│   ├── comments1.xml            # (optional) legacy cell notes ("comments" in OOXML terminology)
│   ├── threadedComments/        # (optional) modern threaded comments
│   │   └── threadedComments1.xml
│   ├── persons/                 # (optional) people directory for threaded comments
│   │   └── persons1.xml
│   ├── commentsExt1.xml         # (optional) comment extension metadata (may be unreferenced)
│   ├── cellimages.xml           # (optional) workbook-level cell image store (name/casing varies; may also appear as xl/cellImages.xml; observed in real Excel fixtures: fixtures/xlsx/rich-data/images-in-cell.xlsx; fixtures/xlsx/images-in-cells/image-in-cell.xlsx)
│   ├── calcChain.xml            # Calculation order hints
│   ├── metadata.xml             # Cell/value metadata (Excel "Rich Data")
│   ├── richData/                # Excel 365+ rich values (data types, in-cell images; naming/casing varies)
│   │   ├── rdrichvalue.xml              # or: richValue.xml / richValues.xml (naming varies)
│   │   ├── rdrichvaluestructure.xml     # or: richValueStructure.xml
│   │   ├── rdRichValueTypes.xml         # or: richValueTypes.xml (casing varies)
│   │   ├── richValueRel.xml             # Indirection to rich-value relationships (e.g. images)
│   │   └── _rels/
│   │       └── richValueRel.xml.rels    # richValueRel -> xl/media/* (image binaries)
│   ├── theme/
│   │   └── theme1.xml           # Color/font theme
│   ├── worksheets/
│   │   ├── sheet1.xml           # Cell data, formulas
│   │   └── sheet2.xml
│   ├── drawings/
│   │   ├── drawing1.xml         # Charts, shapes, images
│   │   └── vmlDrawing1.vml      # (optional) VML shapes for legacy notes/comments
│   ├── media/
│   │   └── image1.png           # Embedded image blobs (used by drawings and in-cell images)
│   ├── charts/
│   │   └── chart1.xml           # Chart definitions
│   ├── tables/
│   │   └── table1.xml           # Table definitions
│   ├── pivotTables/
│   │   └── pivotTable1.xml      # Pivot table definitions
│   ├── pivotCache/
│   │   ├── pivotCacheDefinition1.xml
│   │   └── pivotCacheRecords1.xml
│   ├── slicers/                 # (optional) slicer definitions (pivot UX)
│   ├── slicerCaches/            # (optional) slicer cache definitions
│   ├── timelines/               # (optional) timeline definitions (pivot UX)
│   ├── timelineCaches/          # (optional) timeline cache definitions
│   ├── queryTables/
│   │   └── queryTable1.xml      # External data queries
│   ├── connections.xml          # External data connections
│   ├── externalLinks/
│   │   └── externalLink1.xml    # Links to other workbooks
│   ├── customXml/               # Power Query definitions (base64)
│   └── vbaProject.bin           # VBA macros (binary)
└── xl/_rels/
    ├── workbook.xml.rels        # Workbook relationships
    ├── cellimages.xml.rels      # (optional) relationships for cellimages.xml -> xl/media/* (name/casing varies in the wild)
    └── metadata.xml.rels        # (optional) relationships from metadata.xml -> xl/richData/*
```

See also:
- [`docs/20-images-in-cells-richdata.md`](./20-images-in-cells-richdata.md) — Excel RichData parts used by “Images in Cell” / `IMAGE()` (naming varies: `richValue*` vs `rdrichvalue*`).
- [`docs/20-images-in-cells.md`](./20-images-in-cells.md) — high-level “images in cells” overview + round-trip constraints.

---

## Key Components

### Workbook Sheet List (Order / Name / Visibility)

Sheet *tabs* come from `xl/workbook.xml`:

```xml
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <!-- Sheet order is the tab order -->
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Hidden" sheetId="2" r:id="rId2" state="hidden"/>
    <sheet name="VeryHidden" sheetId="3" r:id="rId3" state="veryHidden"/>
  </sheets>
</workbook>
```

Key rules for Excel-like behavior + safe round-trip:
- **Order matters**: the order of `<sheet>` elements is the user-visible tab order.
- **`sheetId` is stable**: do not renumber sheets when reordering; new sheets typically use `max(sheetId)+1`.
- **Visibility**:
  - Missing `state` ⇒ visible.
  - `state="hidden"` ⇒ hidden (user can unhide).
  - `state="veryHidden"` ⇒ very hidden (Excel UI does not offer unhide; only VBA).
- **`r:id` must be preserved**: other parts refer to the worksheet via the relationship ID in `xl/_rels/workbook.xml.rels`.

### Sheet Tab Color

Tab color is stored per worksheet (not in `workbook.xml`) inside `xl/worksheets/sheetN.xml`:

```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr>
    <tabColor rgb="FFFF0000"/>
  </sheetPr>
  <!-- ... -->
</worksheet>
```

Notes:
- Excel can store color as `rgb`, `theme`/`tint`, or `indexed`. We must parse and write these attributes without loss.
- Missing `<tabColor>` means “no custom tab color”.

### Worksheet XML Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <selection activeCell="A1" sqref="A1"/>
    </sheetView>
  </sheetViews>
  
  <sheetFormatPr defaultRowHeight="15"/>
  
  <cols>
    <col min="1" max="1" width="12.5" style="1" customWidth="1"/>
  </cols>
  
  <sheetData>
    <row r="1" spans="1:3">
      <c r="A1" s="1" t="s">           <!-- t="s" = shared string -->
        <v>0</v>                        <!-- Index into sharedStrings -->
      </c>
      <c r="B1" s="2">                  <!-- No t = number -->
        <v>42.5</v>
      </c>
      <c r="C1" s="3">
        <f>A1+B1</f>                    <!-- Formula -->
        <v>42.5</v>                      <!-- Cached value -->
      </c>
    </row>
  </sheetData>
  
  <conditionalFormatting sqref="A1:C10">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
      <formula>100</formula>
    </cfRule>
  </conditionalFormatting>
  
  <dataValidations count="1">
    <dataValidation type="list" sqref="D1:D100">
      <formula1>"Option1,Option2,Option3"</formula1>
    </dataValidation>
  </dataValidations>
  
  <hyperlinks>
    <hyperlink ref="E1" r:id="rId1"/>
  </hyperlinks>
  
  <mergeCells count="1">
    <mergeCell ref="F1:G2"/>
  </mergeCells>
</worksheet>
```

### Cell Value Types
 
| Type Attribute | Meaning | Value Content |
|---------------|---------|---------------|
| (absent) | Number | Raw numeric value |
| `t="s"` | Shared String | Index into sharedStrings.xml |
| `t="str"` | Inline String | String in `<v>` element |
| `t="inlineStr"` | Rich Text | `<is><t>text</t></is>` |
| `t="b"` | Boolean | 0 or 1 |
| `t="e"` | Error | Error string (#VALUE!, etc.) |
| `t="d"` | Date (typed) | ISO-8601 text in `<v>`.<br>Excel should interpret as a date serial for calculation.<br>Round-trip must preserve `t` and the original ISO string. |

Implementation note (Formula): `formula-xlsx` currently treats `t="d"` as an *opaque* cell type and
stores the value as a plain string in the workbook model, while preserving the original `t` + raw
`<v>` text in round-trip metadata. This lets us rewrite `sheetData` without corrupting typed date
cells (see `crates/formula-xlsx/src/write/mod.rs`).

#### Images in Cells (`IMAGE()` / “Place in Cell”) (Rich Data + `metadata.xml`)

Newer Excel builds can store **images as cell values** (“Place in Cell” pictures, and the `IMAGE()` function) using workbook-level Rich Data parts (`xl/metadata.xml` + `xl/richData/*`) and worksheet cell metadata attributes like `c/@vm`.

This is distinct from legacy “floating” images stored under `xl/drawings/*`.

Further reading:
- [20-images-in-cells.md](./20-images-in-cells.md)
- [20-images-in-cells-richdata.md](./20-images-in-cells-richdata.md)
- [20-xlsx-rich-data.md](./20-xlsx-rich-data.md)
- [xlsx-embedded-images-in-cells.md](./xlsx-embedded-images-in-cells.md) (concrete “Place in Cell” schema walkthrough + exact URIs)

##### Worksheet cell encoding

Excel has been observed to use multiple worksheet-level encodings for “Place in Cell” images. In all
observed variants, the image binding comes from metadata pointers (`vm`, sometimes `cm`) into
`xl/metadata.xml` + `xl/richData/*` (not the cached `<v>` value).

```xml
<!-- Variant A: error cell with cached "#VALUE!" (observed in fixtures/xlsx/basic/image-in-cell.xlsx) -->
<c t="e" vm="N"><v>#VALUE!</v></c>

<!-- Variant B: placeholder numeric cached value (observed in fixtures/xlsx/rich-data/images-in-cell.xlsx) -->
<c vm="N" cm="M"><v>0</v></c>

<!-- Variant C: normal formula cell with vm metadata (observed in fixtures/xlsx/images-in-cells/image-in-cell.xlsx) -->
<c vm="N"><f>_xlfn.IMAGE("...")</f><v>0</v></c>
```

Indexing note:
- Excel commonly uses **1-based** `vm` values, but some producers/fixtures use **0-based** `vm` values. For round-trip safety, treat `vm` as an opaque integer index and preserve it exactly.
- Other common cell attributes (unrelated to the image binding) still apply, e.g. `r="A1"` (cell reference) and `s="…"` (style index).

Note on `IMAGE()` vs “Place in Cell”:
- “Place in Cell” pictures have been observed to use multiple encodings (including both variants shown
  above), depending on Excel build / producer.
- `IMAGE()` function results may instead be stored as a normal formula cell (e.g. `<f>_xlfn.IMAGE(...)</f>`) with `vm` metadata attached; preserve `vm`/`cm` and rich-data parts the same way.
  - Observed in: `fixtures/xlsx/images-in-cells/image-in-cell.xlsx`

##### Mapping chain (high-level)

`sheetN.xml c@vm` → `xl/metadata.xml <valueMetadata>` → `xl/richData/rdrichvalue.xml` (or `xl/richData/richValue*.xml`) → `xl/richData/richValueRel*.xml` → `xl/richData/_rels/richValueRel*.xml.rels` → `xl/media/imageN.*`

##### `xl/richData/rdrichvalue.xml` (rich value instances)

`xl/richData/rdrichvalue.xml` is the workbook-level table of rich value *instances*. The `i="…"` from `xlrd:rvb` selects a record from this table.

The exact element vocabulary inside each rich value varies by Excel version and feature, but for in-cell
images the rich value ultimately encodes a **relationship slot index** (an integer) that points into the
RichData relationship-slot table part (often named `xl/richData/richValueRel.xml`).

Representative shape (observed in the real Excel fixture `fixtures/xlsx/basic/image-in-cell.xlsx`):

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <!-- rich value index 0 -->
  <rv s="0">
    <!-- `_rvRel:LocalImageIdentifier` (relationship-slot index into richValueRel.xml) -->
    <v>0</v>
    <!-- `CalcOrigin` (Excel flag; preserve) -->
    <v>5</v>
  </rv>
</rvData>
```

##### 1) `vm="N"` maps into `xl/metadata.xml`

`xl/metadata.xml` stores workbook-level metadata tables. For rich values, `vm="N"` selects the `N`th `<bk>` (“metadata block”) inside `<valueMetadata>`, which then links (via extension metadata) to a rich value record stored under `xl/richData/`.

Example (`xl/metadata.xml`, representative):

```xml
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000"/>
  </metadataTypes>

  <!-- Rich-value indirection table (referenced from <rc v="…">) -->
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{...}">
          <!-- i = index into xl/richData/rdrichvalue.xml (or xl/richData/richValue*.xml) -->
          <xlrd:rvb i="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>

  <!-- vm="1" selects the first <bk> (Excel often uses 1-based vm). -->
  <valueMetadata count="1">
    <bk>
      <!-- t = index into <metadataTypes> (0-based or 1-based; 1-based is observed in the Excel fixtures in this repo),
           v = (0-based) index into <futureMetadata> -->
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
```

##### 2) Workbook relationship types (workbook → richData parts)

Excel wires the rich-data parts via OPC relationships. Relationship type URIs used for in-cell images include:

- `http://schemas.microsoft.com/office/2022/10/relationships/richValueRel`
- `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue`
- `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure`
- `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes`
- (variant seen in some producers/fixtures) `http://schemas.microsoft.com/office/2017/06/relationships/richValue`
- (variant seen in some producers/fixtures) `http://schemas.microsoft.com/office/2017/06/relationships/richValueRel`
- (variant, when richData parts are related from `xl/metadata.xml` via `xl/_rels/metadata.xml.rels`)
  - `http://schemas.microsoft.com/office/2017/relationships/richValue`
  - `http://schemas.microsoft.com/office/2017/relationships/richValueRel`
  - `http://schemas.microsoft.com/office/2017/relationships/richValueTypes`
  - `http://schemas.microsoft.com/office/2017/relationships/richValueStructure`
- (inside `xl/richData/_rels/richValueRel.xml.rels`) `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`

Representative `xl/_rels/workbook.xml.rels` snippet:

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <!-- ... -->
  <Relationship Id="rIdMeta"
                <!-- Excel has been observed to use either `.../sheetMetadata` or `.../metadata` here. -->
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
                Target="metadata.xml"/>

  <Relationship Id="rIdRV"
                Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue"
                Target="richData/rdrichvalue.xml"/>
  <Relationship Id="rIdRVS"
                Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure"
                Target="richData/rdrichvaluestructure.xml"/>
  <Relationship Id="rIdRVT"
                Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes"
                Target="richData/rdRichValueTypes.xml"/>

  <Relationship Id="rIdRel"
                Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel"
                Target="richData/richValueRel.xml"/>
</Relationships>
```

Another common variant (unprefixed `richValue*.xml` names):

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <!-- ... -->
  <Relationship Id="rIdMeta"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"
                Target="metadata.xml"/>

  <Relationship Id="rIdRV"
                Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue"
                Target="richData/richValue.xml"/>
  <Relationship Id="rIdRel"
                Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel"
                Target="richData/richValueRel.xml"/>
</Relationships>
```

Some packages instead attach the richData parts to `xl/metadata.xml` via `xl/_rels/metadata.xml.rels` (rather than directly from `xl/workbook.xml`):

```xml
<!-- xl/_rels/metadata.xml.rels -->
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.microsoft.com/office/2017/relationships/richValueTypes"
                Target="richData/richValueTypes.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.microsoft.com/office/2017/relationships/richValueStructure"
                Target="richData/richValueStructure.xml"/>
  <Relationship Id="rId3"
                Type="http://schemas.microsoft.com/office/2017/relationships/richValueRel"
                Target="richData/richValueRel.xml"/>
  <Relationship Id="rId4"
                Type="http://schemas.microsoft.com/office/2017/relationships/richValue"
                Target="richData/richValue.xml"/>
</Relationships>
```

##### `xl/richData/rdRichValueTypes.xml` + `xl/richData/rdrichvaluestructure.xml` (supporting schema tables; XML shape varies)

When workbooks use the `rdRichValue*` naming scheme, Excel can emit additional “schema” parts under
`xl/richData/`.

These parts are **not always required** to follow the **image byte chain** (which mainly depends on
`xl/metadata.xml` → `xl/richData/rdrichvalue.xml` → `xl/richData/richValueRel.xml` → `.rels` → `xl/media/*`),
but they are important for interpreting rich value payloads and should be preserved for round-trip safety.

Observed in the real Excel fixture `fixtures/xlsx/basic/image-in-cell.xlsx`:

`xl/richData/rdrichvaluestructure.xml` (structure `_localImage` and the positional key list):

```xml
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <s t="_localImage">
    <k n="_rvRel:LocalImageIdentifier" t="i"/>
    <k n="CalcOrigin" t="i"/>
  </s>
</rvStructures>
```

`xl/richData/rdRichValueTypes.xml` (key flag metadata; note the different root and namespace):

```xml
<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
             xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
             mc:Ignorable="x"
             xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <global>
    <keyFlags>
      <key name="_Self">
        <flag name="ExcludeFromFile" value="1"/>
        <flag name="ExcludeFromCalcComparison" value="1"/>
      </key>
      <!-- ... -->
    </keyFlags>
  </global>
</rvTypesInfo>
```

See also:

- [`docs/20-images-in-cells-richdata.md`](./20-images-in-cells-richdata.md) — broader richValue*/rdRichValue* tables + index-base notes
- [`fixtures/xlsx/basic/image-in-cell.md`](../fixtures/xlsx/basic/image-in-cell.md) — fixture walkthrough

##### 3) `richValueRel.xml` → `xl/media/*` via `.rels`

`xl/richData/richValueRel.xml` is an **ordered table** that maps a small integer slot index (referenced from rich values) to an `r:id`, which is then resolved via `xl/richData/_rels/richValueRel.xml.rels`.

`xl/richData/richValueRel.xml`:

```xml
<!-- Root name and namespace are version-dependent.
     Three fixture-backed variants are observed in the in-repo `.xlsx` fixtures; tests may use additional
     arbitrary namespaces, so treat namespace URIs as opaque. -->

<!-- `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` -->
<richValueRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <!-- slot 0 -->
  <rel r:id="rId1"/>
</richValueRel>

<!-- `fixtures/xlsx/rich-data/images-in-cell.xlsx` -->
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <!-- slot 0 -->
  <rels>
    <rel r:id="rId1"/>
  </rels>
</rvRel>

<!-- `fixtures/xlsx/basic/image-in-cell.xlsx` -->
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <!-- slot 0 -->
  <rel r:id="rId1"/>
</richValueRels>
```

`xl/richData/_rels/richValueRel.xml.rels`:

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="../media/image1.png"/>
</Relationships>
```

Important indexing note:

* `richValueRel.xml` is an **ordered** `<rel>` list; rich values reference it by integer slot index.
* `richValueRel.xml.rels` is an **unordered** map from relationship ID (`rId*`) to `Target`; do not assume
  the `.rels` file lists `<Relationship>` entries in the same order as `richValueRel.xml`.

##### 4) `[Content_Types].xml` overrides

In this repo’s fixture corpus, workbooks that include the extra metadata/richData parts also include
explicit `[Content_Types].xml` overrides for them (e.g. `fixtures/xlsx/basic/image-in-cell.xlsx`,
`fixtures/xlsx/basic/image-in-cell-richdata.xlsx`, `fixtures/xlsx/rich-data/images-in-cell.xlsx`).
Preserve whatever the source workbook uses; do not hardcode a single required set.

```xml
<Override PartName="/xl/metadata.xml"
          ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>

<Override PartName="/xl/richData/rdrichvalue.xml"
          ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>
<Override PartName="/xl/richData/rdrichvaluestructure.xml"
          ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>
<Override PartName="/xl/richData/rdRichValueTypes.xml"
          ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>
<Override PartName="/xl/richData/richValueRel.xml"
          ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
```

For the unprefixed `richValue*.xml` naming variant, content types are typically similar (preserve whatever the source workbook uses):

```xml
<Override PartName="/xl/richData/richValue.xml"
          ContentType="application/vnd.ms-excel.richvalue+xml"/>
<Override PartName="/xl/richData/richValueRel.xml"
          ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
<Override PartName="/xl/richData/richValueTypes.xml"
          ContentType="application/vnd.ms-excel.richvaluetypes+xml"/>
<Override PartName="/xl/richData/richValueStructure.xml"
          ContentType="application/vnd.ms-excel.richvaluestructure+xml"/>
```

##### Note: `xl/cellimages.xml` is optional (and may not appear in “Place in Cell” files)

Some online discussions reference `xl/cellimages.xml` (sometimes `xl/cellImages.xml`) for in-cell
pictures.

In the “Place in Cell” fixtures we inspected, in-cell images were represented using
`xl/metadata.xml` + `xl/richData/*` + `xl/media/*` and **no `cellImages` part** (e.g.
`fixtures/xlsx/basic/image-in-cell.xlsx`).

However, other real Excel workbooks do include a `cellimages` store part in addition to RichData
(e.g. `fixtures/xlsx/rich-data/images-in-cell.xlsx` contains `xl/cellimages.xml` +
`xl/_rels/cellimages.xml.rels`). Other producers and synthetic fixtures/tests may also include a
standalone `cellImages` part.

So for round-trip safety we should treat `xl/cellimages*.xml` / `xl/cellImages*.xml` as optional and
preserve it byte-for-byte if present.

### Linked data types / Rich values (Stocks, Geography, etc.)

Modern Excel supports **linked data types** (Stocks, Geography, Organization, Power BI, etc.) where a cell has a normal displayed value *plus* an attached structured payload (used for the “card” UI and for field extraction like `=A1.Price`).

At the worksheet level, this shows up as additional attributes on the `<c>` (cell) element:

- `vm="…"`, the **value metadata index**
- `cm="…"`, the **cell metadata index**

These are integer indices into metadata tables stored at the workbook level. In this repo we observe
both **1-based** (real Excel fixtures) and **0-based** (synthetic fixtures) `vm` indexing, so treat
`vm`/`cm` as opaque and preserve them exactly.

Example cell carrying rich/linked metadata:

```xml
<row r="2">
  <!-- The visible value is a shared string, but rich metadata is attached. -->
  <c r="A2" t="s" s="1" vm="1" cm="7">
    <v>0</v>
  </c>
</row>
```

Workbook-level metadata is stored in `xl/metadata.xml`. The exact contents vary by Excel version and feature usage, but common top-level tables include:

- `<metadataTypes>`: declares the metadata “type” records referenced by the other tables.
- `<valueMetadata>`: indexed by cell `vm`.
- `<cellMetadata>`: indexed by cell `cm`.
- `<futureMetadata>`: an extension container often present for newer features (including rich values).

Minimal sketch of `xl/metadata.xml` structure:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000"/>
  </metadataTypes>

  <valueMetadata count="2">
    <bk><!-- ... --></bk>
    <bk><!-- ... --></bk>
  </valueMetadata>

  <cellMetadata count="1">
    <bk><!-- ... --></bk>
  </cellMetadata>

  <futureMetadata name="XLRICHVALUE" count="1">
    <!-- ... -->
  </futureMetadata>
</metadata>
```

The structured payloads referenced by this metadata live under `xl/richData/`. Excel has been observed to
use multiple naming schemes, including:

- `rdrichvalue.xml` + `rdRichValueTypes.xml` / `rdrichvaluestructure.xml` (prefixed)
- `richValue.xml` + `richValueTypes.xml` / `richValueStructure.xml` (unprefixed)

In both cases, the image payload is commonly referenced indirectly via `richValueRel.xml` (and its
`xl/richData/_rels/richValueRel.xml.rels` relationships).

These pieces are connected via OPC relationships:

- `xl/_rels/workbook.xml.rels` typically has a relationship from the workbook to `xl/metadata.xml`.
- `xl/_rels/metadata.xml.rels` may have relationships from `xl/metadata.xml` to `xl/richData/*` parts.
- Some workbooks instead include direct relationships from `xl/_rels/workbook.xml.rels` to the richData tables
  (e.g. `richData/richValue.xml` and `richData/richValueRel.xml`) using Microsoft-specific relationship type URIs.

Simplified relationship sketch (Excel uses multiple relationship type URIs for workbook → metadata; richData linkage types vary across builds and may be linked either directly from `workbook.xml.rels` or indirectly via `metadata.xml.rels`):

```xml
<!-- xl/_rels/workbook.xml.rels -->
<Relationship Id="rIdMeta"
              <!-- Observed as either `.../metadata` or `.../sheetMetadata` depending on producer/version. -->
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"
              <!-- also observed: http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata -->
              Target="metadata.xml"/>

<!-- xl/_rels/metadata.xml.rels -->
<Relationship Id="rIdRich1"
              Type="…/relationships/richData"
              Target="richData/richValueTypes.xml"/>
<Relationship Id="rIdRich2"
              Type="…/relationships/richData"
              Target="richData/richValue.xml"/>
```

Packaging note:
- Workbooks that include these parts also typically add `[Content_Types].xml` `<Override>` entries for
  `/xl/metadata.xml` and the `xl/richData/*.xml` parts. For round-trip safety, preserve those
  declarations byte-for-byte when possible (Excel may warn/repair if content types are missing or
  mismatched).

**Formula’s current strategy:**

- Preserve `cm` and `<extLst>` on `<c>` elements when editing cell values.
- Preserve `vm` for untouched cells, but drop `vm` when overwriting a cell’s value/formula away from
  rich-value placeholder semantics (until Formula implements rich-value editing).
- Preserve `xl/metadata.xml`, `xl/richData/**`, and their relationship parts byte-for-byte whenever possible.
- Treat linked-data-type metadata as **opaque** until the calculation engine and data model grow first-class rich/linked value support.
- (Implementation sanity) This preservation behavior is covered by XLSX round-trip / patch tests (e.g. `crates/formula-xlsx/tests/sheetdata_row_col_attrs.rs`).

Further reading:
- [20-images-in-cells.md](./20-images-in-cells.md) (images-in-cell are implemented using the same rich-value + metadata mechanism).
- [20-images-in-cells-richdata.md](./20-images-in-cells-richdata.md) (additional detail on `xl/richData/*` part shapes and index indirection).

### Formula Storage
  
```xml
<!-- Simple formula -->
<c r="A1">
  <f>SUM(B1:B10)</f>
  <v>150</v>
</c>

<!-- Shared formula (for filled ranges) -->
<c r="A1">
  <f t="shared" ref="A1:A10" si="0">B1*2</f>
  <v>10</v>
</c>
<c r="A2">
  <f t="shared" si="0"/>  <!-- References shared formula -->
  <v>20</v>
</c>

<!-- Array formula (legacy CSE style) -->
<c r="A1">
  <f t="array" ref="A1:A5">TRANSPOSE(B1:F1)</f>
  <v>1</v>
</c>

<!-- Dynamic array formula (Excel 365) -->
<c r="A1">
  <f t="array" ref="A1:A5" aca="true">UNIQUE(B1:B100)</f>
  <v>First</v>
</c>

<!-- Formula with _xlfn. prefix for newer functions -->
<c r="A1">
  <f>_xlfn.XLOOKUP(D1,A1:A10,B1:B10)</f>
  <v>Result</v>
</c>
```

### Shared Strings

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
     count="100" uniqueCount="50">
  <si><t>Hello World</t></si>
  <si><t>Another String</t></si>
  <si>
    <r>  <!-- Rich text with formatting runs -->
      <rPr><b/><sz val="12"/></rPr>
      <t>Bold</t>
    </r>
    <r>
      <rPr><sz val="12"/></rPr>
      <t> Normal</t>
    </r>
  </si>
</sst>
```

### Styles

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
  </numFmts>
  
  <fonts count="2">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
    <font>
      <b/>
      <sz val="14"/>
      <color rgb="FF0000FF"/>
      <name val="Arial"/>
    </font>
  </fonts>
  
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  
  <borders count="2">
    <border><!-- empty border --></border>
    <border>
      <left style="thin"><color auto="1"/></left>
      <right style="thin"><color auto="1"/></right>
      <top style="thin"><color auto="1"/></top>
      <bottom style="thin"><color auto="1"/></bottom>
    </border>
  </borders>
  
  <cellXfs count="3">  <!-- Cell formats reference fonts/fills/borders by index -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="1" fillId="1" borderId="1" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>
```

### In-cell images (cellimages.xml)

Some workbooks (including real Excel workbooks in this repo) can store “images in cell” (pictures that
behave like cell content rather than floating drawing objects) in a dedicated workbook-level OPC part:

- Part: `xl/cellimages.xml` (casing varies; `xl/cellImages.xml` is also seen in the wild)
- Relationships: `xl/_rels/cellimages.xml.rels` (casing varies in lockstep with the XML part name)

Excel has been observed (in this repo) to use **both** encodings:

- RichData-only (no `xl/cellimages.xml` / `xl/cellImages.xml`): `fixtures/xlsx/basic/image-in-cell.xlsx`
- RichData **plus** `xl/cellimages.xml`: `fixtures/xlsx/rich-data/images-in-cell.xlsx`

This repo also includes the **synthetic** fixture `fixtures/xlsx/basic/cellimages.xlsx` (notes in
`fixtures/xlsx/basic/cellimages.md`), which *does*
contain a standalone `xl/cellimages.xml` part; treat it as an alternate store and preserve it when
present.

If a `cellImages` part is present, we should preserve it for round-trip safety.

From a **packaging / round-trip** perspective, the important thing is the relationship chain that connects this part to the actual image blobs under `xl/media/*`.

**Schema note:** `xl/cellimages.xml` is a Microsoft extension part; the root namespace / element vocabulary
varies across producers and Excel builds. In this repo, `…/2022/cellimages` is confirmed in a real
Excel fixture, while `…/2019/cellimages` is only observed in synthetic fixtures/tests so far. For
round-trip, treat the **part path** (including its original casing) as authoritative, not the root namespace.

#### How it’s usually connected

1. `xl/workbook.xml` (via `xl/_rels/workbook.xml.rels`) contains a relationship that targets `cellimages.xml` (or other casing variants):
   - The relationship **Type URI is a Microsoft extension** and has been observed to vary across Excel builds.
   - **Detection strategy**: treat any relationship whose `Target` resolves to either `xl/cellimages.xml` or `xl/cellImages.xml` (preserving the original casing) as authoritative, rather than hardcoding a single `Type` URI.
2. `xl/_rels/cellimages.xml.rels` contains relationships of type `…/relationships/image` pointing at `xl/media/*` files.
   - The relationship `Id` values (e.g. `rId1`) are referenced from within `xl/cellimages.xml` (either via `r:embed` on an `<a:blip>` or via `r:id`/`r:embed` on a `<cellImage>`), so they must be preserved (or updated consistently if rewriting).
   - Targets are typically relative paths like `media/image1.png` (resolving to `/xl/media/image1.png`), but should be preserved as-is.

#### `[Content_Types].xml` requirements

If `xl/cellimages.xml` is present, the package typically includes an override:

- `<Override PartName="/xl/cellimages.xml" ContentType="…"/>` (casing varies)

Excel uses a **Microsoft-specific** content type string for this part (the exact string may vary between versions).

Observed in this repo (see `crates/formula-xlsx/tests/cell_images.rs` and
`crates/formula-xlsx/tests/cellimages_preservation.rs`; do not hardcode):
- `application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml`
- `application/vnd.ms-excel.cellimages+xml`

**Preservation/detection strategy:**
- Treat any `[Content_Types].xml` `<Override>` whose `PartName` is `/xl/cellimages.xml` or `/xl/cellImages.xml` as authoritative.
- Preserve the `ContentType` value byte-for-byte on round-trip; **do not** hardcode a single MIME string in the writer.

#### Relationship type URIs

- `xl/_rels/cellimages.xml.rels` (casing varies) → `xl/media/*`:
  - **High confidence**: `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"`
- `xl/workbook.xml.rels` → `xl/cellimages.xml` / `xl/cellImages.xml`:
  - **Confirmed in the synthetic fixture** `fixtures/xlsx/basic/cellimages.xlsx`:
    - `http://schemas.microsoft.com/office/2022/relationships/cellImages`
  - Observed variants in tests/synthetic inputs:
    - `http://schemas.microsoft.com/office/2020/relationships/cellImages`
    - `http://schemas.microsoft.com/office/2020/07/relationships/cellImages`
  - Prefer detection by `Target`/part name when possible.

#### Minimal (non-normative) XML snippets

Workbook relationship entry (in `xl/_rels/workbook.xml.rels`):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <!-- fixtures/xlsx/basic/cellimages.xlsx -->
  <Relationship Id="rId3"
                Type="http://schemas.microsoft.com/office/2022/relationships/cellImages"
                Target="cellimages.xml"/>
</Relationships>
```

Cellimages-to-media relationship entry (in `xl/_rels/cellimages.xml.rels`):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png"/>
</Relationships>
```

Cellimages part referencing an image by relationship id (in `xl/cellimages.xml`):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage>
    <a:blip r:embed="rId1"/>
  </cellImage>
</cellImages>
```

---

## The Five Hardest Compatibility Problems

### 1. Conditional Formatting Version Divergence

Excel 2007 and Excel 2010+ use different XML schemas for the same visual features.

**Excel 2007 Data Bar:**
```xml
<conditionalFormatting sqref="A1:A10">
  <cfRule type="dataBar" priority="1">
    <dataBar>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FF638EC6"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>
```

**Excel 2010+ Data Bar (extended features):**
```xml
<x14:conditionalFormattings xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
    <x14:cfRule type="dataBar" id="{GUID}">
      <x14:dataBar minLength="0" maxLength="100" gradient="0" direction="leftToRight">
        <x14:cfvo type="autoMin"/>
        <x14:cfvo type="autoMax"/>
        <x14:negativeFillColor rgb="FFFF0000"/>
        <x14:axisColor rgb="FF000000"/>
      </x14:dataBar>
    </x14:cfRule>
    <xm:sqref>A1:A10</xm:sqref>
  </x14:conditionalFormatting>
</x14:conditionalFormattings>
```

**Strategy**: 
- Parse both schemas
- Convert internally to unified representation
- Write back preserving original schema version
- Use MC:AlternateContent for cross-version compatibility

### 2. Chart Fidelity (DrawingML + ChartEx)

Charts use DrawingML, a complex XML schema for vector graphics. Several newer Excel chart types (e.g. histogram, waterfall, treemap) are often stored using **ChartEx** (an “extended chart” schema) referenced from the drawing layer.

```xml
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$A$1</c:f></c:strRef></c:tx>
          <c:cat><!-- Categories --></c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$1:$B$10</c:f>
              <c:numCache>
                <c:ptCount val="10"/>
                <c:pt idx="0"><c:v>100</c:v></c:pt>
                <!-- ... -->
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
```

**Challenges:**
- Different applications render same XML differently
- Absolute positioning in EMUs (English Metric Units) anchored to sheet cells
- Complex inheritance of styles from theme
- Version-specific chart types (ChartEx / extension lists, Excel 2016+)
- Multiple related OPC parts (`drawing*.xml`, `chart*.xml`, `chartEx*.xml`, `style*.xml`, `colors*.xml`)

**Strategy:**
- Treat charts as a **lossless subsystem**: preserve chart-related parts byte-for-byte unless the user explicitly edits charts.
- Parse and render supported chart types incrementally; render a placeholder for unsupported chart types.
- Resolve theme-based colors and number formats using the same machinery as cells.
- Test with fixtures and visual regression against Excel output.

**Detail spec:** [17-charts.md](./17-charts.md)

### 3. Date Systems (The Lotus Bug)

Excel supports two date systems:

| System | Epoch | Day 1 |
|--------|-------|-------|
| 1900 (Windows default) | January 1, 1900 | Serial 1 |
| 1904 (Mac legacy) | January 1, 1904 | Serial 0 |

**The Lotus 1-2-3 Bug:**
Excel 1900 system incorrectly treats 1900 as a leap year (it wasn't). February 29, 1900 is serial 60, though this date never existed.

```
Serial 59 = February 28, 1900
Serial 60 = February 29, 1900  ← INVALID DATE
Serial 61 = March 1, 1900
```

**Implications:**
- Dates before March 1, 1900 are off by 1 day
- Mixing 1900 and 1904 workbooks creates 1,462 day differences
- We must emulate this bug for compatibility

**Strategy:**
```typescript
function serialToDate(serial: number, dateSystem: "1900" | "1904"): Date {
  if (dateSystem === "1900") {
    // Emulate Lotus bug
    if (serial < 60) {
      return addDays(new Date(1899, 11, 31), serial);
    } else if (serial === 60) {
      // Invalid date - Feb 29, 1900 didn't exist
      return new Date(1900, 1, 29);  // Represent as-if
    } else {
      // After the bug, off by one
      return addDays(new Date(1899, 11, 30), serial);
    }
  } else {
    return addDays(new Date(1904, 0, 1), serial);
  }
}
```

### 4. Dynamic Array Function Prefixes

Excel 365 introduced dynamic array functions that require `_xlfn.` prefix in file storage:

```xml
<!-- Stored in file -->
<f>_xlfn.UNIQUE(_xlfn.FILTER(A1:A100,B1:B100>0))</f>

<!-- Displayed to user -->
=UNIQUE(FILTER(A1:A100,B1:B100>0))
```

**Functions requiring prefix:**
- UNIQUE, FILTER, SORT, SORTBY, SEQUENCE
- XLOOKUP, XMATCH
- RANDARRAY
- LET, LAMBDA, MAP, REDUCE, SCAN, MAKEARRAY
- Many others added post-2010

**Opening in older Excel:**
```
=_xlfn.XLOOKUP(...)  ← Shown as formula text, #NAME? error
```

**Strategy:**
- Strip prefix on parse for display
- Add prefix on save for file compatibility
- Maintain list of all prefixed functions with version introduced

### 5. VBA Macro Preservation

VBA is stored as binary (`vbaProject.bin`) following OLE compound document format:

```
vbaProject.bin (OLE container)
├── VBA/
│   ├── _VBA_PROJECT     # VBA metadata
│   ├── dir              # Module directory (compressed)
│   ├── Module1          # Module source (compressed)
│   └── ThisWorkbook     # Workbook module
├── PROJECT              # Project properties
└── PROJECTwm            # Project web module
```

**Challenges:**
- Binary format with compression
- Digital signatures must be preserved (and, for “signed-only” macro trust modes, validated/bound to
  the MS-OVBA Contents Hash (v1/v2) / `ContentsHashV3` (v3) binding digest)
- No standard library for creation (only preservation)
- Security implications of execution

**Strategy:**
- Preserve `vbaProject.bin` byte-for-byte on round-trip
- Parse for display/inspection (MS-OVBA specification)
- Defer execution to Phase 2 or via optional component
- Offer migration path to Python/TypeScript

See also: [`vba-digital-signatures.md`](./vba-digital-signatures.md) (signature stream location,
payload variants, and MS-OVBA digest binding plan).

---

## Parsing Libraries and Tools

### By Platform

| Library | Platform | Formulas | Charts | VBA | Pivot |
|---------|----------|----------|--------|-----|-------|
| Open XML SDK | .NET | R/W | Full | Preserve | Partial |
| Apache POI | Java | Eval+R/W | Limited | Preserve | Limited |
| openpyxl | Python | R/W | Good | No | Preserve |
| xlrd/xlwt | Python | Read only | No | No | No |
| SheetJS | JavaScript | Read | Pro only | No | Pro only |
| calamine | Rust | Read | No | No | No |
| rust_xlsxwriter | Rust | Write | Partial | No | No |
| libxlsxwriter | C | Write | Good | No | No |

### Recommended Approach

1. **Reading**: Start with calamine (Rust, fast) for data extraction
2. **Writing**: rust_xlsxwriter for basic files
3. **Full fidelity**: Custom implementation following ECMA-376
4. **Reference**: Apache POI for behavior verification

---

## Round-Trip Preservation Strategy

### Principle: Preserve What We Don't Understand

In the current Rust implementation, there are multiple layers depending on the workload:

- For **OPC-level round-trip** without inflating every ZIP entry into memory, use
  `formula_xlsx::XlsxLazyPackage` (lazy reads + streaming ZIP rewrite on save).
- For algorithms that need whole-package random access, use `formula_xlsx::XlsxPackage` (fully
  materialized part map; writing generally re-packs the ZIP).

Conceptually, the round-trip contract looks like:

```typescript
interface XlsxDocument {
  // Parsed workbook model (data + formulas + modeled metadata).
  workbook: Workbook;

  // Preserved OPC parts as raw (uncompressed) bytes for round-trip fidelity.
  // This includes both XML parts we don't understand and binary parts like `xl/vbaProject.bin`.
  preservedParts: Map<PartPath, Uint8Array>;
}
```

**Terminology note:** throughout this document, “byte-for-byte” preservation typically refers to the
**OPC part payload bytes** (i.e. the uncompressed bytes you get after inflating a ZIP entry). When
using the streaming save path, **untouched ZIP entries** are often copied with
`zip::ZipWriter::raw_copy_file(...)`, which also preserves the original compressed bytes for those
entries.

Forward-compat note: worksheet `<sheetProtection>` elements may include modern hashing attributes
(`algorithmName`, `hashValue`, `saltValue`, `spinCount`). Even if the model only exposes the legacy
`password` hash, these attributes must be preserved on round-trip when protection settings are
unchanged (see Feature preservation rules: Worksheet protection, and the patch-in-place writer in
`crates/formula-xlsx/src/write/mod.rs`).

### Relationship ID Preservation

```xml
<!-- Original -->
<Relationship Id="rId1" Type="...worksheet" Target="worksheets/sheet1.xml"/>

<!-- WRONG: Regenerated IDs break internal references -->
<Relationship Id="rId5" Type="...worksheet" Target="worksheets/sheet1.xml"/>

<!-- CORRECT: Preserve original IDs -->
<Relationship Id="rId1" Type="...worksheet" Target="worksheets/sheet1.xml"/>
```

### Markup Compatibility (MC) Namespace

For forward compatibility with features we don't support:

```xml
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice Requires="x14">
    <!-- Excel 2010+ specific content -->
    <x14:sparklineGroups>...</x14:sparklineGroups>
  </mc:Choice>
  <mc:Fallback>
    <!-- Fallback for applications that don't support x14 -->
  </mc:Fallback>
</mc:AlternateContent>
```

**Strategy**: Preserve AlternateContent blocks, process Choice if we support the namespace.

---

## XLSB (Binary Format)

XLSB uses the same ZIP structure but binary records instead of XML:

```
Benefits:
- 2-3x faster to open/save
- 50% smaller file size
- Same feature support as XLSX

Structure:
- Same ZIP layout
- .bin files instead of .xml
- Records: [type: u16][size: u32][data: bytes]
```

### Binary Record Format

```
Record Structure:
┌──────────┬──────────┬────────────────┐
│ Type (2) │ Size (4) │ Data (variable)│
└──────────┴──────────┴────────────────┘

Example - Cell Value Record:
Type: 0x0002 (BrtCellReal)
Size: 8
Data: IEEE 754 double

Example - Formula Record:
Type: 0x0006 (BrtCellFmla)
Size: variable
Data: [value][flags][formula_bytes]
```

**Strategy**: Support XLSB reading for performance, focus XLSX for primary format.

---

## Testing Strategy

### Compatibility Test Suite

1. **Unit tests**: Each file component (cells, formulas, styles, etc.)
2. **Integration tests**: Complex workbooks with multiple features
3. **Round-trip tests**: Load → Save → Load, compare
4. **Cross-application tests**: Save from us, open in Excel; save from Excel, open in us
5. **Real-world corpus**: Test against collection of user-submitted files

### Test File Categories

| Category | Examples | Focus Areas |
|----------|----------|-------------|
| Basic | Simple data, formulas | Core functionality |
| Styling | Rich formatting, themes | Visual fidelity |
| Charts | All chart types | DrawingML rendering |
| Pivots | Complex pivot tables | Pivot cache, definitions |
| External | Links, queries | Connection handling |
| Large | 1M+ rows | Performance, memory |
| Legacy | Excel 97-2003 | .xls conversion |
| Complex | Financial models | Everything together |

### Automated Comparison

```typescript
interface ComparisonResult {
  identical: boolean;
  differences: Difference[];
}

interface Difference {
  path: string;  // e.g., "xl/worksheets/sheet1.xml/row[5]/c[3]/v"
  type: "missing" | "added" | "changed";
  original?: string;
  modified?: string;
  severity: "critical" | "warning" | "info";
}

async function compareWorkbooks(
  original: XlsxFile,
  roundTripped: XlsxFile
): Promise<ComparisonResult> {
  // Compare structure
  // Compare XML content with normalization
  // Compare binary parts byte-for-byte
  // Report all differences with severity
}
```

### Implemented Round-Trip Harness (xlsx-diff)

This repository includes a small XLSX fixture corpus and a part-level diff tool to
validate load → save → diff round-trips:

- Fixtures live under `fixtures/xlsx/**` (kept intentionally small).
- Encrypted/password-protected OOXML workbooks (e.g. `.xlsx`, `.xlsm`, `.xlsb` “Encrypt with Password”) are **OLE/CFB containers**, not ZIP/OPC packages.
  They live under `fixtures/encrypted/` (see `fixtures/encrypted/ooxml/` for the vendored encrypted `.xlsx`/`.xlsm` corpus) and are intentionally excluded from the `xlsx-diff::collect_fixture_paths` round-trip corpus.
- The diff tool is implemented in Rust: `crates/xlsx-diff`.

Run a diff locally:

```bash
bash scripts/cargo_agent.sh run -p xlsx-diff --bin xlsx_diff -- original.xlsx roundtripped.xlsx
```

Run the fixture harness (used by CI):

```bash
bash scripts/cargo_agent.sh test -p xlsx-diff --test roundtrip_fixtures
```

The harness performs a real load → save using the in-memory
`formula-xlsx::XlsxPackage` repacker (OPC-level package handling) and then diffs the original vs
written output.

Note: the **product save path** for “edit an existing `.xlsx`/`.xlsm`” is typically the streaming
rewrite pipeline (see [Performance Considerations](#performance-considerations)), which preserves
unmodified ZIP entries via `raw_copy_file` and avoids materializing the entire package.

Current normalization rules (to reduce false positives):

- Ignore whitespace-only XML text nodes unless `xml:space="preserve"` is set.
- Sort XML attributes (namespace declarations are ignored; resolved URIs are used instead).
- Sort `<Relationships>` entries by `(Id, Type, Target)`.
- Sort `[Content_Types].xml` entries by `(Default.Extension | Override.PartName)`.
- Sort worksheet `<sheetData>` rows/cells by their `r` attributes.

Current severity policy (subject to refinement as the writer matures):

- **critical**: missing parts, changes in `[Content_Types].xml`, changes in `*.rels`, any binary part diffs.
- **warning**: non-essential parts like themes / calcChain, extra parts.
- **info**: metadata-only changes under `docProps/*`.

---

## Performance Considerations

### Streaming ZIP rewrite (round-trip preservation)

When we edit an existing `.xlsx` / `.xlsm` and want to **preserve everything we don’t understand**
(charts, pivots, `customXml/`, VBA, media blobs, etc.), we avoid rebuilding the whole ZIP from
scratch.

Instead, we write a new archive using a **streaming ZIP rewriter**:

- The input ZIP is read once; the output ZIP is written once.
- **Untouched entries are copied byte-for-byte** using `zip::ZipWriter::raw_copy_file(...)` (no
  inflate + re-deflate), which preserves existing compression and avoids loading large binaries into
  memory.
- Only the small set of **modified** parts are inflated and rewritten (typically worksheets,
  `xl/sharedStrings.xml`, `[Content_Types].xml`, relevant `*.rels`, etc).

This is the core of `.xlsx/.xlsm` round-trip compatibility in the “edit existing workbook” save
path.

### Lazy package API: `formula_xlsx::XlsxLazyPackage`

`XlsxLazyPackage` is a thin wrapper around an existing OPC ZIP container (a file path on disk, or
an owned byte buffer) plus an in-memory map of explicit part overrides:

- Open without inflating all parts:
  - `XlsxLazyPackage::open(path)` / `XlsxLazyPackage::from_file(...)` (native)
  - `XlsxLazyPackage::from_bytes(...)` / `XlsxLazyPackage::from_vec(...)` (in-memory)
- Read only what you need: `read_part("xl/workbook.xml")?` inflates just that part.
- Override only what you change: `set_part("xl/workbook.xml", bytes)`.
- Save using the streaming rewriter: `write_to(...)` / `write_to_bytes()` (raw-copy for untouched
  ZIP entries).
- Note: today `XlsxLazyPackage::write_to` is only available on native targets (it is not supported
  on `wasm32`).

Use `XlsxLazyPackage` when you need OPC-level round-trip preservation but want to keep memory usage
low for large workbooks.

### In-memory package API: `formula_xlsx::XlsxPackage`

`XlsxPackage::from_bytes(...)` inflates the entire ZIP into a `BTreeMap<part_name, Vec<u8>>`. This
is convenient for algorithms that need random access to many parts at once, but it has real memory
cost (and enforces safety limits to avoid ZIP bombs).

Unlike the streaming pipeline, writing an `XlsxPackage` generally **re-packs** the ZIP (recompresses
entries). The *decompressed* part bytes are preserved, but the on-disk ZIP representation (compression
method/levels, ordering, etc.) may change.

Full materialization happens when:

- you construct an `XlsxPackage` (e.g. `XlsxPackage::from_bytes` / `from_bytes_limited`), and/or
- you operate on the full part map (`parts_map()` / `parts_map_mut()`).

Prefer the lazy/streaming path unless you truly need whole-package inspection or mutation.

### Streaming parsing (worksheet-scale)

Even with a streaming ZIP wrapper, individual parts like `xl/worksheets/sheetN.xml` can be large
when inflated. Avoid DOM-parsing whole worksheets when you only need a subset of information.

General guidance:

- Prefer streaming XML readers (SAX-style) for worksheet-scale tasks.
- Keep cell-level parsing incremental (row-by-row) so memory usage is proportional to the
  “working set”, not the full file.

### Lazy loading (parse only what you touch)

Many workloads only need a small subset of parts:

- workbook metadata (`xl/workbook.xml`, `*.rels`)
- styles (`xl/styles.xml`)
- shared strings (`xl/sharedStrings.xml`)
- specific worksheets

Avoid parsing “everything” up-front; defer until the caller needs it.

### Parallelism (best-effort)

Some parts are independent and can be parsed concurrently (styles, shared strings, workbook
metadata). Worksheets typically depend on those, but then can often be processed in parallel if the
caller has a multi-core budget.

---

## Future Considerations

1. **Excel for Web compatibility**: Some features differ in web version
2. **Google Sheets export**: Import/export from Google's format
3. **Numbers compatibility**: Apple's format for Mac users
4. **OpenDocument (ODS)**: LibreOffice compatibility
5. **New Excel features**: Monitor Excel updates for new XML schemas
6. **Images in Cell** (Place in Cell / `IMAGE()`): packaging + schema notes in [20-images-in-cells.md](./20-images-in-cells.md)
