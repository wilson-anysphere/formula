# Excel “Images in Cells” (`IMAGE()` / “Place in Cell”) — OOXML storage + Formula plan

## Scope

This doc is an **internal compatibility spec** for how Excel stores “images in cells” in `.xlsx` (OOXML)
and what Formula must parse/preserve to support:

- `IMAGE()` function results
- Insert → Pictures → **Place in Cell** (as a cell value)

**Out of scope:** UI rendering, layout/sizing behavior, image decoding, network fetch, caching policies.
This doc is strictly about **file parts + relationships + worksheet references** needed for correct
load/save round-trips.

## Background: “floating” images vs “images in cells”

Excel has (at least) two distinct storage mechanisms:

1. **Floating images / shapes** anchored to cells via DrawingML
   - Stored under `xl/drawings/*` with image binaries under `xl/media/*`.
   - Already covered by the general DrawingML preservation strategy.
2. **Images in cells** (newer Excel / Microsoft 365)
   - Primarily stored via the workbook-level **Rich Data** system (`xl/metadata.xml` + `xl/richData/*`).
   - Some workbooks (including real Excel fixtures in this repo) additionally include a dedicated **cell
     image store** part (`xl/cellimages.xml` / `xl/cellImages.xml`). Treat this as an alternate/legacy
     store and preserve **both** graphs when present.
   - The rest of this document focuses on this second mechanism.

## Expected OOXML parts

Workbooks using images-in-cells are expected to include some/all of the following parts:

```
xl/
├── cellimages.xml                # Optional (some files use richData-only wiring; preserve if present)
├── cellimages1.xml               # Optional; allow numeric suffixes like other indexed XLSX parts
├── media/
│   └── image*.{png,jpg,gif,...}
├── metadata.xml
├── _rels/
│   ├── cellimages.xml.rels       # Optional (only if a cellimages.xml part exists)
│   └── metadata.xml.rels (commonly present when `metadata.xml` references `xl/richData/*`)
└── richData/
    # Observed rich value stores (see notes below):
    ├── richValue.xml                 # legacy/2017 variant rich values (some producers use `richValues.xml`)
    ├── richValueTypes.xml            # optional legacy types table
    ├── richValueStructure.xml        # optional legacy structure table
    ├── rdrichvalue.xml               # modern “Place in Cell” rich values (rdRichValue naming)
    ├── rdrichvaluestructure.xml      # rdRichValue structure table (defines <v> positions)
    ├── rdRichValueTypes.xml          # rdRichValue types table
    # Shared relationship-slot table (used by both variants):
    ├── richValueRel.xml              # integer slot -> r:id mapping (name may be numbered/custom in some producers/tests)
    └── _rels/
        └── richValueRel.xml.rels     # `.rels` for the slot-table part (e.g. `richValueRel.xml.rels` or `richValueRel1.xml.rels`)
```

Notes:

- **Part-name casing:** The real Excel fixture in this repo uses `xl/cellimages.xml` (all lowercase),
  but some producers (and some synthetic fixtures/tests) use `xl/cellImages.xml` (camel-case). OPC part
  names are case-sensitive inside the ZIP, so:
  - readers should handle both variants, and
  - writers should preserve the original casing when round-tripping an existing file.
- **`cellimages.xml` may be absent:** Some workbooks store image-in-cell values entirely via
  `xl/metadata.xml` + `xl/richData/*` (especially the `richValueRel*` relationship-slot table → `.rels` → `xl/media/*`) without a
  separate `xl/cellimages.xml` / `xl/cellImages.xml` part.
  - Observed in the real Excel fixture [`fixtures/xlsx/basic/image-in-cell.xlsx`](../fixtures/xlsx/basic/image-in-cell.xlsx) (notes in
    [`fixtures/xlsx/basic/image-in-cell.md`](../fixtures/xlsx/basic/image-in-cell.md)).
  - Also observed in the synthetic minimal fixture `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` (notes in
    `fixtures/xlsx/basic/image-in-cell-richdata.md`) and `crates/formula-xlsx/tests/rich_data_roundtrip.rs`.
  - If a `cellImages` store part exists, preserve it and its relationship graph.
- **`cellimages.xml` may be present:** Other Excel-generated workbooks include a dedicated cell image store
  part in addition to `xl/metadata.xml` + `xl/richData/*`.
  - Observed in [`fixtures/xlsx/rich-data/images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx) (notes in
    [`fixtures/xlsx/rich-data/images-in-cell.md`](../fixtures/xlsx/rich-data/images-in-cell.md)).
  - If present, preserve it and its relationship graph.
- `xl/media/*` contains the actual image bytes (usually `.png`, but Excel may use other formats).
- The exact `xl/richData/*` file set can vary across Excel builds; the part names shown above include the
  two **observed** naming schemes in this repo:
  - legacy: `richValue*.xml` (some producers/tests also use the plural `richValues*.xml`)
  - modern “Place in Cell”: `rdrichvalue*.xml` + supporting structure/types tables
  Formula should preserve the entire `xl/richData/` directory byte-for-byte (OPC part payload bytes)
  unless we explicitly implement rich-value editing.
- `xl/metadata.xml` and the per-cell `c/@vm` + `c/@cm` attributes connect worksheet cells to the rich
  value system.
- When present, `xl/_rels/metadata.xml.rels` typically connects `xl/metadata.xml` → `xl/richData/*` parts.
  Formula should preserve these relationships byte-for-byte (OPC part payload bytes) for safe
  round-trips.

See also: [20-images-in-cells-richdata.md](./20-images-in-cells-richdata.md) for a deeper (still
best-effort) description of the `richValue*`/**`rdRichValue*`** part sets and how the `richValueRel*`
relationship-slot table is used to resolve media relationships.

For a concrete, fixture-backed “Place in Cell” schema walkthrough (including the `rdrichvalue*` keys
`_rvRel:LocalImageIdentifier` and `CalcOrigin`), see:

- [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md)
  - The Excel-produced fixture [`fixtures/xlsx/basic/image-in-cell.xlsx`](../fixtures/xlsx/basic/image-in-cell.xlsx) uses the same richData-only wiring
    (and does **not** use `xl/cellimages.xml`), with notes in [`fixtures/xlsx/basic/image-in-cell.md`](../fixtures/xlsx/basic/image-in-cell.md).

## In-repo fixtures (cell image store part)
This repo includes a few small fixtures that exercise the workbook-level `cellImages` part. These are
useful for confirming namespace + relationship `Type` URI variants, independent of the Rich Data wiring
used by modern Excel “Place in Cell”.

### Fixture: `fixtures/xlsx/basic/cell-images.xlsx` (camel-case part name)

Fixture workbook: [`fixtures/xlsx/basic/cell-images.xlsx`](../fixtures/xlsx/basic/cell-images.xlsx) (notes in
[`fixtures/xlsx/basic/cell-images.md`](../fixtures/xlsx/basic/cell-images.md)).

Note:

- `fixtures/xlsx/basic/cell-images.xlsx` and `fixtures/xlsx/basic/cellimages.xlsx` are **synthetic**
  (generated by Formula’s fixture tooling; `docProps/app.xml` reports `Application="Formula Fixtures"`).
  They are useful for parser/preservation tests, but their namespaces/relationship types should not be
  treated as Excel ground truth.
- `fixtures/xlsx/rich-data/images-in-cell.xlsx` is a **real Excel** fixture and should be treated as the
  ground truth for modern Excel’s `xl/cellimages.xml` wiring.

Observed values:

- Part paths:
  - `xl/cellImages.xml`
  - `xl/_rels/cellImages.xml.rels`
  - `xl/media/image1.png`
- `xl/cellImages.xml` root namespace:
  - `http://schemas.microsoft.com/office/spreadsheetml/2023/02/main`
- Workbook → `cellImages.xml` relationship `Type` URI (in `xl/_rels/workbook.xml.rels`):
  - `http://schemas.microsoft.com/office/2023/02/relationships/cellImage`

### Fixture: `fixtures/xlsx/basic/cellimages.xlsx` (lowercase part name)

Fixture workbook: [`fixtures/xlsx/basic/cellimages.xlsx`](../fixtures/xlsx/basic/cellimages.xlsx) (notes in
[`fixtures/xlsx/basic/cellimages.md`](../fixtures/xlsx/basic/cellimages.md)).

Observed values:

- Part paths:
  - `xl/cellimages.xml`
  - `xl/_rels/cellimages.xml.rels`
  - `xl/media/image1.png`
- `xl/cellimages.xml` root namespace:
  - `http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages`
- Workbook → `cellimages.xml` relationship `Type` URI (in `xl/_rels/workbook.xml.rels`):
  - `http://schemas.microsoft.com/office/2022/relationships/cellImages`

### Fixture: `fixtures/xlsx/rich-data/images-in-cell.xlsx` (real Excel)

Fixture workbook: [`fixtures/xlsx/rich-data/images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx) (notes in [`fixtures/xlsx/rich-data/images-in-cell.md`](../fixtures/xlsx/rich-data/images-in-cell.md)).

Confirmed values from this fixture:

- Part paths:
  - `xl/cellimages.xml` (note lowercase path)
  - `xl/_rels/cellimages.xml.rels`
  - `xl/media/image1.png`
- `xl/cellimages.xml` root namespace:
  - `http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages`
- `[Content_Types].xml` override for the part:
  - `application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml`
- Workbook → `cellimages.xml` relationship `Type` URI (in `xl/_rels/workbook.xml.rels`):
  - `http://schemas.microsoft.com/office/2019/relationships/cellimages`

### Fixture: `fixtures/xlsx/images-in-cells/image-in-cell.xlsx` (real Excel; Place in Cell + `IMAGE()`)

Fixture workbook: [`fixtures/xlsx/images-in-cells/image-in-cell.xlsx`](../fixtures/xlsx/images-in-cells/image-in-cell.xlsx).
Notes: [`fixtures/xlsx/images-in-cells/image-in-cell.md`](../fixtures/xlsx/images-in-cells/image-in-cell.md).

This fixture is intended as a ground-truth-ish Excel workbook that contains **both**:

- Insert → Pictures → **Place in Cell** (cell `A1`)
- a formula using `_xlfn.IMAGE(...)` (cell `B1`)

Confirmed values from this fixture:

- Part paths:
  - `xl/cellimages.xml`
  - `xl/_rels/cellimages.xml.rels`
  - `xl/metadata.xml`
  - `xl/richData/richValue.xml`
  - `xl/richData/richValueRel.xml`
  - `xl/richData/richValueTypes.xml`
  - `xl/richData/richValueStructure.xml`
  - `xl/richData/_rels/richValueRel.xml.rels`
  - `xl/media/image1.png`
- Worksheet metadata pointers:
  - `xl/worksheets/sheet1.xml` contains `c/@vm` on both `A1` and `B1`
  - `B1` contains a formula with `_xlfn.IMAGE("https://example.com/image.png")`
- `[Content_Types].xml` overrides:
  - `/xl/cellimages.xml`: `application/vnd.ms-excel.cellimages+xml`
  - `/xl/metadata.xml`: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml`
  - `/xl/richData/richValue.xml`: `application/vnd.ms-excel.richvalue+xml`
  - `/xl/richData/richValueRel.xml`: `application/vnd.ms-excel.richvaluerel+xml`
  - `/xl/richData/richValueTypes.xml`: `application/vnd.ms-excel.richvaluetypes+xml`
  - `/xl/richData/richValueStructure.xml`: `application/vnd.ms-excel.richvaluestructure+xml`
- Workbook-level relationship `Type` URIs (in `xl/_rels/workbook.xml.rels`):
  - workbook → metadata:
    - `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"`
    - `Target="metadata.xml"`
  - workbook → cell images store:
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/cellImages"`
    - `Target="cellimages.xml"`
  - workbook → richData parts:
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue"`
      - `Target="richData/richValue.xml"`
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel"`
      - `Target="richData/richValueRel.xml"`
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueTypes"`
      - `Target="richData/richValueTypes.xml"`
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueStructure"`
      - `Target="richData/richValueStructure.xml"`
- Image relationship usage:
  - `xl/_rels/cellimages.xml.rels` contains `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"`
    pointing to `Target="media/image1.png"`.
  - `xl/richData/_rels/richValueRel.xml.rels` contains the same image `Type`, pointing to
    `Target="../media/image1.png"`.

### Quick reference: `cellImages` part graph (OPC + XML)

This section is a “what to look for” summary for the core **cell image store** parts. Details and
variant shapes are documented further below.

#### Confirmed vs unconfirmed

**Confirmed (from in-repo fixtures/tests):**

- A workbook can contain a dedicated `cellImages` part (seen in tests as `xl/cellimages.xml` and
  `xl/cellImages.xml`) plus a matching relationship part at `xl/_rels/<part>.rels`.
- The **synthetic** fixture `fixtures/xlsx/basic/cell-images.xlsx` contains `xl/cellImages.xml` with root namespace:
  - `http://schemas.microsoft.com/office/spreadsheetml/2023/02/main`
- The **synthetic** fixture `fixtures/xlsx/basic/cellimages.xlsx` contains `xl/cellimages.xml` with root namespace:
  - `http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages`
- The real Excel fixture [`fixtures/xlsx/rich-data/images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx) contains `xl/cellimages.xml` with namespace:
  - `http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages`
- These fixtures contain `<cellImage>` entries that reference images via:
  - `<a:blip r:embed="rIdX"/>`
- The `cellImages` XML can reference binary images via DrawingML-style `r:embed="rIdX"` references.
- `rIdX` is resolved through the `*.rels` part to an image under `xl/media/*`.
- Image relationship type is the standard OOXML one:
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`

**Unconfirmed / still needs additional real Excel samples (`IMAGE()` and version variants):**

- Whether Excel ever uses multiple numbered parts like `cellimages1.xml` / `cellImages1.xml` (only a single
  `cellimages.xml` / `cellImages.xml` part has been observed in this repo so far).
- The full set of namespaces used by real Excel builds for `cellImages` across versions.
  - Confirmed in the real Excel fixture `fixtures/xlsx/rich-data/images-in-cell.xlsx`:
    - `http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages`
  - Also observed in synthetic fixtures/tests (treat as opaque and preserve):
    - `http://schemas.microsoft.com/office/spreadsheetml/2023/02/main`
    - `http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages`
    - `http://schemas.microsoft.com/office/spreadsheetml/2020/07/main`
    - `http://schemas.microsoft.com/office/spreadsheetml/2019/11/main`
- Exact schema shape(s) emitted by real Excel across versions.
  - Confirmed in the real Excel fixture `fixtures/xlsx/rich-data/images-in-cell.xlsx`: `<cellImage>`
    contains a full DrawingML `<xdr:pic>` subtree with an `<a:blip r:embed="…"/>` reference.
  - It is still unknown whether Excel ever emits the lightweight `<cellImage><a:blip .../></cellImage>`
    form (that shape is currently only observed in synthetic fixtures/tests in this repo).
- Whether Excel consistently uses a single relationship `Type` URI (and whether the relationship is
  always on `xl/_rels/workbook.xml.rels` vs sometimes worksheet-level).
- The exact “cell → image” mapping mechanism across **all** Excel scenarios.
  - Confirmed for a rust_xlsxwriter-generated **“Place in Cell”** workbook (used for schema verification in this repo):
    - worksheet cell is `t="e"` with cached `#VALUE!` and `vm="1"`
    - image bytes are resolved via `xl/metadata.xml` + `xl/richData/rd*` + `xl/richData/richValueRel*.xml(.rels)` → `xl/media/*`
    - no `xl/cellimages.xml`/`xl/cellImages.xml` part is used in that case
    - see: [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md)
  - Confirmed for a real Excel-generated fixture workbook checked into this repo:
    - [`fixtures/xlsx/basic/image-in-cell.xlsx`](../fixtures/xlsx/basic/image-in-cell.xlsx) (see [`fixtures/xlsx/basic/image-in-cell.md`](../fixtures/xlsx/basic/image-in-cell.md))
    - same pattern: error cell `t="e"`/`#VALUE!` + `vm="…"` → `xl/metadata.xml` + `xl/richData/rd*` + `richValueRel` → `xl/media/*`
    - no `xl/cellimages.xml`/`xl/cellImages.xml` part is used in that fixture
  - Confirmed for a real Excel-generated fixture workbook that includes **both** `xl/cellimages.xml` and `xl/richData/*`:
    - [`fixtures/xlsx/rich-data/images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx) (notes in [`fixtures/xlsx/rich-data/images-in-cell.md`](../fixtures/xlsx/rich-data/images-in-cell.md))
    - cell `A1` uses `vm="1" cm="1"` with a numeric cached `<v>0</v>` (not `t="e"`/`#VALUE!`)
  - Still an open question for real Excel-generated `IMAGE()` results and other producers.

#### Parts

- `xl/cellimages.xml` (**preferred**, but casing can vary; see note above)
- `xl/_rels/cellimages.xml.rels`
- image binaries: `xl/media/imageN.<ext>`

#### XML namespace + structure (observed)

- Root element local name: `<cellImages>`
- Confirmed in the real Excel fixture `fixtures/xlsx/rich-data/images-in-cell.xlsx`:
  - namespace: `http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages`
- Also observed in synthetic fixtures/tests (treat as opaque and preserve):
  - `http://schemas.microsoft.com/office/spreadsheetml/2023/02/main` (synthetic fixture `fixtures/xlsx/basic/cell-images.xlsx`)
  - `http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages` (synthetic fixture `fixtures/xlsx/basic/cellimages.xlsx`)
  - `http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages`
  - `http://schemas.microsoft.com/office/spreadsheetml/2020/07/main`
  - `http://schemas.microsoft.com/office/spreadsheetml/2019/11/main`
- The root contains one or more `<cellImage>` entries. Some schemas embed a full DrawingML picture
  subtree (e.g. `<xdr:pic>`; observed in the real Excel fixture `fixtures/xlsx/rich-data/images-in-cell.xlsx`),
  but the in-repo synthetic fixtures/tests also use a lightweight:
  - `<cellImage><a:blip r:embed="rIdX"/></cellImage>`
- `r:embed="rIdX"` is resolved via `xl/_rels/cellimages.xml.rels` to a `Target` under `xl/media/*`.

#### Content types (observed)

- `[Content_Types].xml` override (often present; preserve whatever is in the source workbook). Observed in this repo:

```xml
<!-- Real Excel fixture (`fixtures/xlsx/rich-data/images-in-cell.xlsx`) -->
<Override PartName="/xl/cellimages.xml"
          ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml"/>

<!-- Real Excel fixture (`fixtures/xlsx/images-in-cells/image-in-cell.xlsx`) -->
<Override PartName="/xl/cellimages.xml"
          ContentType="application/vnd.ms-excel.cellimages+xml"/>
```

#### Relationship types

- Image relationship type (standard OOXML):
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`
- Relationship type for “workbook → `cellImages*.xml` / `cellimages*.xml`” discovery (Microsoft extension; **variable**):
  - Confirmed in the real Excel fixture [`fixtures/xlsx/rich-data/images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx):
    - `http://schemas.microsoft.com/office/2019/relationships/cellimages`
  - Observed in synthetic fixtures/tests:
    - `http://schemas.microsoft.com/office/2023/02/relationships/cellImage`
    - `http://schemas.microsoft.com/office/2022/relationships/cellImages`
  - Observed variants in tests/synthetic inputs:
    - `http://schemas.microsoft.com/office/2020/relationships/cellImages`
    - `http://schemas.microsoft.com/office/2020/07/relationships/cellImages`
  - **Detection rule:** prefer identifying the relationship by resolved `Target` part name
    (`cellImages*.xml` / `cellimages*.xml`) rather than hardcoding a single `Type` URI.

#### Minimal example (`xl/cellimages.xml`) (synthetic)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages"
            xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage>
    <xdr:pic>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
      </xdr:blipFill>
    </xdr:pic>
  </cellImage>
</cellImages>
```

#### Minimal example (`xl/_rels/cellimages.xml.rels`) (synthetic)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png"/>
</Relationships>
```

#### Minimal example (`xl/_rels/workbook.xml.rels` entry) (fixtures; **Type URI observed to vary**)

Some files link `xl/workbook.xml` → `xl/cellImages*.xml` via an OPC relationship in
`xl/_rels/workbook.xml.rels`. The relationship `Type` is Microsoft-specific and has been observed to
vary; prefer detecting by `Target` when possible.

```xml
<!-- fixtures/xlsx/basic/cell-images.xlsx -->
<Relationship Id="rId3"
              Type="http://schemas.microsoft.com/office/2023/02/relationships/cellImage"
              Target="cellImages.xml"/>

<!-- fixtures/xlsx/basic/cellimages.xlsx -->
<Relationship Id="rId3"
              Type="http://schemas.microsoft.com/office/2022/relationships/cellImages"
              Target="cellimages.xml"/>
```

## Worksheet cell references (`c/@vm`, `c/@cm`, `<extLst>`)

SpreadsheetML’s `<c>` (cell) element can carry metadata indices:

- `c/@vm` — **value metadata index** (used to associate a cell’s *value* with a record in `xl/metadata.xml`)
- `c/@cm` — **cell metadata index** (used to associate the *cell* with a record in `xl/metadata.xml`)

For round-trip safety, Formula must preserve these attributes (and the `<extLst>` subtree) for **untouched**
cells when round-tripping, because they can “point” to image/rich-value structures elsewhere in the
package.

However, if Formula **overwrites a cell’s value/formula**, it cannot currently update the corresponding
rich-value tables, so it drops `vm="…"` (value metadata) to avoid leaving a dangling pointer. In contrast,
`cm="…"` and `<extLst>` are preserved (see `crates/formula-xlsx/tests/cell_metadata_preservation.rs`).

### Minimal examples (from existing Formula fixtures/tests)

The repository already has fixtures/tests exercising preservation of these attributes:

```xml
<!-- `vm` attribute example (synthetic fixture `fixtures/xlsx/metadata/rich-values-vm.xlsx`) -->
<row r="1">
  <c r="A1" vm="1"><v>1</v></c>
</row>
```

```xml
<!-- `cm` attribute + extLst subtree preservation (crates/formula-xlsx/tests/cell_metadata_preservation.rs) -->
<c r="A1" s="5" cm="7" customAttr="x">
  <v>1</v>
  <extLst>
    <ext uri="{123}">
      <test xmlns="http://example.com">ok</test>
    </ext>
  </extLst>
</c>
```

### Images-in-cells cell shapes (observed)

Cells containing an image-in-cell typically signal that fact via metadata pointers (`c/@vm` and sometimes
`c/@cm`). The cached `<v>` value does **not** directly reference the image bytes; in the fixtures in this
repo it is an internal placeholder/cache value (e.g. `#VALUE!` or `0`).

Observed shapes in **real Excel fixtures** in this repo:

- **“Place in Cell” (embedded local image; richData-only / `rdRichValue*` mapping):** cells are
  error-typed (`t="e"`) with cached `#VALUE!` and a `vm` attribute:

  ```xml
  <c r="B2" t="e" vm="1"><v>#VALUE!</v></c>
  ```

  (see `fixtures/xlsx/basic/image-in-cell.xlsx`)

- **In-cell image with a `cellimages.xml` store part:** cells may be plain numeric with `vm` + `cm` and a
  numeric cached `<v>`:

  ```xml
  <c r="A1" vm="1" cm="1"><v>0</v></c>
  ```

  (see `fixtures/xlsx/rich-data/images-in-cell.xlsx`)

The exact cell shape for **`IMAGE()` formula results** still needs more real Excel samples; it may include
a formula `<f>_xlfn.IMAGE(...)</f>` and/or an `<extLst>` payload. Treat all cell children and attributes as
opaque and preserve them for untouched cells.

**Round-trip rule:** preserve `cm` and `<extLst>` verbatim. Preserve `vm` for untouched cells, but drop
`vm` when overwriting a cell’s value/formula unless we explicitly implement full rich-value editing.

### How `vm` maps to `xl/metadata.xml` and the rich value store (`xl/richData/richValue*.xml` / `xl/richData/rdrichvalue.xml`)

Formula’s current understanding (implemented in `crates/formula-xlsx/src/rich_data/metadata.rs`) is:

1. Worksheet cells reference a *value metadata record* via `c/@vm`.
   - Excel commonly emits `vm` as **1-based**, but **0-based** values are also observed in the wild (and in our
     test-only richData fixtures). Treat `vm` as opaque and preserve it.
2. `xl/metadata.xml` contains `<valueMetadata>` with a list of `<bk>` records; `vm` selects a `<bk>`.
3. That `<bk>` contains `<rc t="…" v="…"/>` where:
    - `t` selects `"XLRICHVALUE"` inside `<metadataTypes>`.
      - In the Excel-generated fixtures in this repo, `t` is **1-based** (`t="1"` when
        `<metadataTypes>` has a single `<metadataType name="XLRICHVALUE"/>`).
      - Other workbooks/tests have been observed to use **0-based** indexing here; treat `t` as ambiguous
        and resolve best-effort.
    - `v` is **0-based** (in this schema it indexes into `<futureMetadata name="XLRICHVALUE">`’s `<bk>` list; other schemas may use `v` differently).
4. That future-metadata `<bk>` contains an extension element (commonly `xlrd:rvb`) with an `i="…"`
    attribute, which is a **0-based** index into the rich value store part (either `xl/richData/richValue*.xml`
    or `xl/richData/rdrichvalue.xml`, depending on file/Excel build).

Representative snippet (from the unit tests in `crates/formula-xlsx/src/rich_data/metadata.rs`):

```xml
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="2">
    <metadataType name="SOMEOTHERTYPE"/>
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>

  <futureMetadata name="XLRICHVALUE" count="2">
    <bk>
      <extLst>
        <ext uri="{...}">
          <xlrd:rvb i="5"/>
        </ext>
      </extLst>
    </bk>
    <bk>
      <extLst>
        <ext uri="{...}">
          <xlrd:rvb i="42"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>

  <valueMetadata count="2">
    <bk><rc t="2" v="0"/></bk> <!-- vm="1" -> rv index 5 -->
    <bk><rc t="2" v="1"/></bk> <!-- vm="2" -> rv index 42 -->
  </valueMetadata>
</metadata>
```

Other observed `xl/metadata.xml` shapes exist. For example, the **synthetic** fixture
`fixtures/xlsx/basic/image-in-cell-richdata.xlsx` uses a `<futureMetadata name="XLRICHVALUE">` table with a
single `<xlrd:rvb i="0"/>` mapping, but its sheet cells use `vm="0"` (0-based).

In that fixture:

- Worksheet cells use `vm="0"` (0-based).
- `<rc t="1" v="0"/>` selects the first `<futureMetadata name="XLRICHVALUE">` `<bk>`.
- `<xlrd:rvb i="0"/>` provides the 0-based rich value index into `xl/richData/richValue.xml`.

See [`docs/20-images-in-cells-richdata.md`](./20-images-in-cells-richdata.md) for the exact XML and
namespaces observed in that fixture. Formula currently treats these schemas as opaque and focuses on
round-trip preservation (with best-effort extraction utilities in `crates/formula-xlsx/src/rich_data/mod.rs`).

## `xl/cellimages.xml` (a.k.a. `xl/cellImages.xml`)

`xl/cellimages.xml` is the workbook-level “cell image store” part. It is expected to contain a list of
image entries that can be referenced (directly or indirectly) by rich values.

Note: not all images-in-cell workbooks include this part (some use `xl/metadata.xml` + `xl/richData/*`
only). When present, it must be preserved along with its `.rels` and referenced media.

The part embeds **SpreadsheetDrawing / DrawingML** `<xdr:pic>` payloads and uses
`<a:blip r:embed="rId…">` to reference an image relationship in `xl/_rels/cellimages.xml.rels`.

Observed root namespaces (from in-repo fixtures/tests; only `.../2022/cellimages` is confirmed in a real
Excel file in this repo so far):

- `http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages` (real Excel + synthetic fixtures)
- `http://schemas.microsoft.com/office/spreadsheetml/2023/02/main`
- `http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages`
- `http://schemas.microsoft.com/office/spreadsheetml/2020/07/main`
- `http://schemas.microsoft.com/office/spreadsheetml/2019/11/main`

Namespace prefixes vary (`cx`, `etc`, or none). Parsers should match by **local-name** (e.g.
`cellImages`, `cellImage`, `pic`, `blip`) and by relationship-namespace attributes, not by prefix.

### Example from synthetic fixture (`fixtures/xlsx/basic/cellimages.xlsx`)

`xl/cellimages.xml`:

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

Representative example (from `crates/formula-xlsx/tests/cell_images.rs`; non-normative):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:cellImages xmlns:cx="http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages"
               xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:pic>
    <xdr:blipFill>
      <a:blip r:embed="rId1"/>
    </xdr:blipFill>
  </xdr:pic>
</cx:cellImages>
```

Another observed shape (from `crates/formula-xlsx/tests/cellimages_preservation.rs`) wraps the `<xdr:pic>`
in a `cellImage` container element:

```xml
<etc:cellImages xmlns:etc="http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <etc:cellImage>
    <xdr:pic>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
      </xdr:blipFill>
    </xdr:pic>
  </etc:cellImage>
</etc:cellImages>
```

Some producers emit a more lightweight schema where the relationship ID is stored directly on a
`<cellImage>` element (rather than within a DrawingML `<pic>` subtree). Formula’s `cell_images` parser
has explicit support for `r:id` on `<cellImage>` (and some variants use `r:embed` instead):

```xml
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:id="rId1"/>
</cellImages>
```

```xml
<cellImage r:embed="rId1"/>
```

### `xl/_rels/cellimages.xml.rels` (a.k.a. `xl/_rels/cellImages.xml.rels`)

`xl/_rels/cellimages.xml.rels` contains OPC relationships from `cellimages.xml` to the binary image parts
under `xl/media/*`.

This relationships file is standard OPC, and the **image relationship type URI is known** (it is the
standard OOXML image relationship type, and is confirmed in the real Excel fixture
`fixtures/xlsx/rich-data/images-in-cell.xlsx`; it is also used in synthetic fixtures/tests like
`fixtures/xlsx/basic/cellimages.xlsx`):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png"/>
</Relationships>
```

Targets are usually relative paths and may appear as `media/imageN.png` or `../media/imageN.png`
(preserve the original `Target` exactly).

**Parser resilience (Formula):** the `cell_images` parser uses a best-effort resolver that tries:

1. standard OPC resolution relative to the source part (`xl/cellimages*.xml`)
2. a fallback relative to the `.rels` part
3. a fallback that re-roots under `xl/` if the path escaped via `..`

See `crates/formula-xlsx/src/cell_images/mod.rs` (`resolve_target_best_effort`).

**Round-trip rules:**

- Preserve `Relationship/@Id` values.
- Preserve `Target` paths and file names (Excel reuses these paths across features).
- Preserve the referenced `xl/media/*` bytes byte-for-byte (OPC part payload bytes).

## `xl/metadata.xml`

`xl/metadata.xml` is the workbook-level part that backs the `c/@vm` and `c/@cm` indices.

At a minimum, Formula must:

- parse and preserve the part itself, and
- preserve any `vm`/`cm` attributes in worksheets that point into it.

Representative skeleton (SpreadsheetML namespace is expected, but element details are fixture-dependent):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <!-- Defines metadata types/strings and the valueMetadata/cellMetadata tables -->
  <!-- ... -->
</metadata>
```

## `xl/richData/*` (rich values)

Excel stores non-primitive “rich” cell values using a set of XML parts under `xl/richData/`.
For images-in-cells, these rich values ultimately resolve to an **image binary** under `xl/media/*`,
but there are multiple schemas and packaging patterns in the ecosystem.

### Two observed rich value stores (richValue* vs rdRichValue*)

This repo currently has fixtures/tests covering two rich value store families:

1. **Legacy / unprefixed store:** `xl/richData/richValue*.xml` (often with optional `richValueTypes.xml` /
   `richValueStructure.xml`).
2. **Modern embedded-image (“Place in Cell”) store:** `xl/richData/rdrichvalue*.xml` +
   `xl/richData/rdrichvaluestructure.xml` + `xl/richData/rdRichValueTypes.xml`.

Both stores use `xl/metadata.xml` + worksheet `c/@vm` to map cells to rich value indices, and both resolve
media via the shared relationship-slot table:

- `xl/richData/richValueRel.xml` → `xl/richData/_rels/richValueRel.xml.rels` → `xl/media/*`.

### Packaging patterns (how the media is wired)

1. **RichData → RichValueRel → media (no `cellimages.xml` part)**
   - Observed in this repo in both:
     - `fixtures/xlsx/basic/image-in-cell.xlsx` (real Excel “Place in Cell” fixture; notes in `fixtures/xlsx/basic/image-in-cell.md`), and
     - `crates/formula-xlsx/tests/embedded_images_place_in_cell_roundtrip.rs`
       (generated with `rust_xlsxwriter::Worksheet::embed_image_with_format`).
    - The image bytes are resolved via:
      - `xl/richData/richValueRel*.xml` → `xl/richData/_rels/richValueRel*.xml.rels` → `xl/media/*`
2. **RichData → RichValueRel → media (and a `cellimages.xml` store part is also present)**
   - Observed in the real Excel fixture:
     - [`fixtures/xlsx/rich-data/images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx) (notes in
       [`fixtures/xlsx/rich-data/images-in-cell.md`](../fixtures/xlsx/rich-data/images-in-cell.md))
    - In this file:
      - `xl/richData/richValueRel*.xml` → `xl/richData/_rels/richValueRel*.xml.rels` → `xl/media/*` resolves the image bytes, and
      - `xl/cellimages.xml` → `xl/_rels/cellimages.xml.rels` → `xl/media/*` also references the same image bytes.
   - The exact semantic role of `xl/cellimages.xml` in this variant (whether it is a cache, a parallel lookup
     table, or is referenced by rich value payloads) is not fully characterized. Treat the `cellimages` part as
     **opaque** and preserve it byte-for-byte (OPC part payload bytes) for safe round-trips.
3. **`cellimages.xml` / `cellImages.xml` “cell image store” → media (standalone)**
   - Observed in this repo via synthetic fixtures/tests like `crates/formula-xlsx/tests/cell_images.rs`.
   - The image bytes are resolved via:
     - `xl/cellimages.xml` → `xl/_rels/cellimages.xml.rels` → `xl/media/*`

**Round-trip rule:** preserve `xl/cellimages*.xml` / `xl/cellImages*.xml` and its `.rels` + referenced media
whenever it is present, even if we do not fully understand how the cell/rich-value mapping uses it.

At minimum, a rich value store is expected to exist when `xl/metadata.xml` indicates the `XLRICHVALUE`
metadata type (the exact mapping schema varies by producer/Excel build).

Depending on the producer, the rich value store may be named:

- `xl/richData/richValue*.xml` (Excel-like naming), or
- `xl/richData/rdrichvalue*.xml` / `xl/richData/rdRichValueTypes.xml` (rdRichValue naming; observed in this
  repo in both real Excel fixtures and `rust_xlsxwriter` output).

See also:

- [`docs/20-images-in-cells-richdata.md`](./20-images-in-cells-richdata.md) — concrete notes on the
  `richValue*` parts (types/structures/values/relationship indirection) and index bases used by Excel.

Because the exact file set and schemas vary across Excel builds, Formula’s short-term strategy is:

- **preserve all `xl/richData/*` parts and their `*.rels`**, and
- treat them as an **atomic bundle** with `xl/metadata.xml` (and `xl/cellimages.xml` if present) during round-trip.

Common file names (Excel version-dependent; treat as “expected shape”, not a strict schema):

- `xl/richData/richValue.xml`
- `xl/richData/richValueRel.xml` (relationship-slot table; may be numbered/custom) + its `.rels` under `xl/richData/_rels/`
- `xl/richData/richValueTypes.xml`
- `xl/richData/richValueStructure.xml`

## `[Content_Types].xml` requirements

Workbooks that include these parts may declare content types in `[Content_Types].xml`.

In this repo, the fixtures that include `xl/metadata.xml` and/or `xl/richData/*` include explicit
`<Override>` entries for those parts. However, in the wild some workbooks may rely on
`<Default Extension="xml" ContentType="application/xml"/>` instead. Implementations should preserve
whatever is in the source workbook.

Independently of overrides for `.xml` parts, image payloads under `xl/media/*` still require appropriate
image MIME defaults (e.g. `png` → `image/png`) for interoperability.

- **Override** entries for XML parts like `/xl/cellimages.xml` (or `/xl/cellImages.xml`), `/xl/metadata.xml`,
  and `xl/richData/*.xml`
- **Default** entries for image extensions used under `/xl/media/*` (`png`, `jpg`, `gif`, etc.)

### `xl/cellimages.xml` content type override

Observed values (preserve whatever is in the source workbook):

- `application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml`
  - observed in the real Excel fixture [`fixtures/xlsx/rich-data/images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx)
  - also used by `crates/formula-xlsx/tests/cell_images.rs`
- `application/vnd.ms-excel.cellimages+xml`
  - used by `crates/formula-xlsx/tests/cellimages_preservation.rs`

Excel uses Microsoft-specific content type strings for this part, and the exact string may vary across
versions/builds.

Note: MIME types are case-insensitive, but for round-trip safety we preserve the `ContentType` string
byte-for-byte (including its original capitalization).

**Round-trip rule:** treat any `<Override PartName="/xl/cellimages.xml" .../>` (or the camel-case variant) as
authoritative and preserve its `ContentType` value byte-for-byte.

If we ever need to synthesize this part from scratch, `application/vnd.ms-excel.cellimages+xml` is a
reasonable default (it matches Excel’s vendor-specific pattern like `...threadedcomments+xml` / `...person+xml`),
but we should still prefer the original file’s value when round-tripping.

### `xl/metadata.xml` content type override

Observed in the real Excel fixture `fixtures/xlsx/rich-data/images-in-cell.xlsx` and the synthetic
fixture `fixtures/xlsx/metadata/rich-values-vm.xlsx`:

- `application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml`

Also observed in tests:

- `application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml`
  - used by `crates/formula-xlsx/tests/metadata_rich_value_roundtrip.rs`

In this repo’s fixtures, `xl/metadata.xml` always has an explicit `<Override>` entry. For robustness,
parsers should not depend on it (locate parts via relationships/paths) and writers should preserve any
existing overrides byte-for-byte when round-tripping.

### `xl/richData/*` content types (observed + variants)

Content types for `xl/richData/*` vary across Excel/producers and across the two naming schemes
(`richValue*.xml` vs `rdRichValue*`).

Observed in the **synthetic** fixture [`fixtures/xlsx/basic/image-in-cell-richdata.xlsx`](../fixtures/xlsx/basic/image-in-cell-richdata.xlsx) (explicit overrides present):

- `/xl/richData/richValue.xml`: `application/vnd.ms-excel.richvalue+xml`
- `/xl/richData/richValueRel.xml`: `application/vnd.ms-excel.richvaluerel+xml`

Observed in [`fixtures/xlsx/basic/image-in-cell.xlsx`](../fixtures/xlsx/basic/image-in-cell.xlsx) (explicit overrides present):

- `/xl/richData/richValueRel.xml`: `application/vnd.ms-excel.richvaluerel+xml`
- `/xl/richData/rdrichvalue.xml`: `application/vnd.ms-excel.rdrichvalue+xml`
- `/xl/richData/rdrichvaluestructure.xml`: `application/vnd.ms-excel.rdrichvaluestructure+xml`
- `/xl/richData/rdRichValueTypes.xml`: `application/vnd.ms-excel.rdrichvaluetypes+xml`

Observed in [`fixtures/xlsx/rich-data/images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx)
(real Excel; explicit overrides present for the unprefixed `richValue*` naming scheme). The synthetic
fixture `fixtures/xlsx/rich-data/richdata-minimal.xlsx` (used by tests) uses the same set:

- `application/vnd.ms-excel.richvalue+xml` (for `/xl/richData/richValue.xml`)
- `application/vnd.ms-excel.richvaluerel+xml` (for `/xl/richData/richValueRel.xml`)
- `application/vnd.ms-excel.richvaluetypes+xml` (for `/xl/richData/richValueTypes.xml`)
- `application/vnd.ms-excel.richvaluestructure+xml` (for `/xl/richData/richValueStructure.xml`)

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- ... -->
  <Default Extension="png" ContentType="image/png"/>
  <!-- ... -->

  <!-- In this repo’s fixtures, explicit overrides are present for these parts. Other producers may omit
       some overrides and rely on the default XML content type; preserve whatever the source workbook uses. -->
  <Override PartName="/xl/cellimages.xml"
             ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml"/>

  <Override PartName="/xl/metadata.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>

  <!-- Unprefixed “richValue*” naming (observed in the synthetic fixture `fixtures/xlsx/rich-data/richdata-minimal.xlsx`
       and the real Excel fixture `fixtures/xlsx/rich-data/images-in-cell.xlsx`; other producers may omit explicit
       overrides even when parts exist). -->
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/vnd.ms-excel.richvaluetypes+xml"/>
  <Override PartName="/xl/richData/richValueStructure.xml" ContentType="application/vnd.ms-excel.richvaluestructure+xml"/>
  <Override PartName="/xl/richData/richValue.xml" ContentType="application/vnd.ms-excel.richvalue+xml"/>
  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
</Types>
```

TODO (fixture-driven): add additional Excel-generated workbooks covering:

- `IMAGE()` results across Excel versions and input types (URL vs local file).
  - We have an initial Excel-produced sample with an `_xlfn.IMAGE(...)` formula cell:
    - `fixtures/xlsx/images-in-cells/image-in-cell.xlsx`
- “Place in Cell” images across Excel versions

and update this section if Excel emits additional `xl/richData/*` part names or content types beyond those
already observed in our fixture corpus (including synthetic fixtures and real Excel fixtures), e.g.:

- `fixtures/xlsx/rich-data/richdata-minimal.xlsx` (synthetic Formula fixture used by tests)
- `fixtures/xlsx/basic/image-in-cell.xlsx` (Excel-generated)
- `fixtures/xlsx/images-in-cells/image-in-cell.xlsx` (Excel-generated; Place in Cell + `_xlfn.IMAGE()`)

and exercised by tests like
`crates/formula-xlsx/tests/richdata_preservation.rs`.

## Relationship type URIs (what we know vs TODO)

Known (stable, used across OOXML):

- Image relationships (used by DrawingML and expected to be used by `cellimages.xml`):
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`

Partially known (fixture-driven details still recommended):

- Workbook → `xl/metadata.xml` relationship:
  - Lives in `xl/_rels/workbook.xml.rels`.
  - Observed in the **synthetic** fixture `fixtures/xlsx/metadata/rich-values-vm.xlsx`:
    - `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"`
  - Also observed in [`fixtures/xlsx/rich-data/images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx):
    - `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"`
  - Also observed in [`fixtures/xlsx/images-in-cells/image-in-cell.xlsx`](../fixtures/xlsx/images-in-cells/image-in-cell.xlsx):
    - `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"`
  - Observed in the **synthetic** fixture [`fixtures/xlsx/basic/image-in-cell-richdata.xlsx`](../fixtures/xlsx/basic/image-in-cell-richdata.xlsx):
    - `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"`
  - Observed in [`fixtures/xlsx/basic/image-in-cell.xlsx`](../fixtures/xlsx/basic/image-in-cell.xlsx):
    - `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"`
  - Preservation is covered by `crates/formula-xlsx/tests/metadata_rich_values_vm_roundtrip.rs`.
- Workbook → richData parts (when stored directly in workbook relationships):
  - Observed in the **synthetic** fixture [`fixtures/xlsx/basic/image-in-cell-richdata.xlsx`](../fixtures/xlsx/basic/image-in-cell-richdata.xlsx):
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue"`
      - `Target="richData/richValue.xml"`
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel"`
      - `Target="richData/richValueRel.xml"`
  - Also observed in the real Excel fixture [`fixtures/xlsx/images-in-cells/image-in-cell.xlsx`](../fixtures/xlsx/images-in-cells/image-in-cell.xlsx):
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue"`
      - `Target="richData/richValue.xml"`
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel"`
      - `Target="richData/richValueRel.xml"`
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueTypes"`
      - `Target="richData/richValueTypes.xml"`
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueStructure"`
      - `Target="richData/richValueStructure.xml"`
- Metadata → richData parts (when linked indirectly via `xl/_rels/metadata.xml.rels`):
  - Observed in the synthetic fixture `fixtures/xlsx/rich-data/richdata-minimal.xlsx` and the real Excel fixture
    [`fixtures/xlsx/rich-data/images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx):
    - `Type="http://schemas.microsoft.com/office/2017/relationships/richValueTypes"`
      - `Target="richData/richValueTypes.xml"`
    - `Type="http://schemas.microsoft.com/office/2017/relationships/richValueStructure"`
      - `Target="richData/richValueStructure.xml"`
    - `Type="http://schemas.microsoft.com/office/2017/relationships/richValueRel"`
      - `Target="richData/richValueRel.xml"`
    - `Type="http://schemas.microsoft.com/office/2017/relationships/richValue"`
      - `Target="richData/richValue.xml"`
- Workbook → `xl/cellimages.xml` relationship:
  - Lives in `xl/_rels/workbook.xml.rels`.
  - Excel uses a Microsoft-extension relationship `Type` URI that has been observed to vary.
  - **Confirmed in fixtures:**
    - `Type="http://schemas.microsoft.com/office/2019/relationships/cellimages"`
      - real Excel fixture: [`fixtures/xlsx/rich-data/images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx)
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/cellImages"`
      - real Excel fixture: [`fixtures/xlsx/images-in-cells/image-in-cell.xlsx`](../fixtures/xlsx/images-in-cells/image-in-cell.xlsx)
    - `Type="http://schemas.microsoft.com/office/2023/02/relationships/cellImage"`
      - synthetic fixture: `fixtures/xlsx/basic/cell-images.xlsx`
    - `Type="http://schemas.microsoft.com/office/2022/relationships/cellImages"`
      - synthetic fixture: `fixtures/xlsx/basic/cellimages.xlsx`
  - Observed variants in tests/synthetic inputs:
    - `Type="http://schemas.microsoft.com/office/2020/relationships/cellImages"`
    - `Type="http://schemas.microsoft.com/office/2020/07/relationships/cellImages"`
  - **Round-trip / detection rule:** identify the relationship by resolved `Target`
    (`/xl/cellimages.xml` or `/xl/cellImages.xml`) rather than hardcoding a single `Type`.
- RichData relationship indirection (images referenced via the `richValueRel*` relationship-slot table):
  - `xl/richData/_rels/richValueRel*.xml.rels` is expected to contain standard image relationships:
    - `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"`
  - Observed in the **synthetic** fixture `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` and the unit test
    `crates/formula-xlsx/tests/rich_data_cell_images.rs`.
  - Workbook → richData relationships (Type URIs are Microsoft-specific and versioned). Observed in this repo:
    - `http://schemas.microsoft.com/office/2017/06/relationships/richValue`
      - synthetic fixture: `image-in-cell-richdata.xlsx`
      - real Excel fixture: `fixtures/xlsx/images-in-cells/image-in-cell.xlsx`
    - `http://schemas.microsoft.com/office/2017/06/relationships/richValueRel`
      - synthetic fixture: `image-in-cell-richdata.xlsx`
      - real Excel fixture: `fixtures/xlsx/images-in-cells/image-in-cell.xlsx`
    - `http://schemas.microsoft.com/office/2017/06/relationships/richValueTypes`
      - real Excel fixture: `fixtures/xlsx/images-in-cells/image-in-cell.xlsx`
    - `http://schemas.microsoft.com/office/2017/06/relationships/richValueStructure`
      - real Excel fixture: `fixtures/xlsx/images-in-cells/image-in-cell.xlsx`
    - `http://schemas.microsoft.com/office/2017/relationships/richValue` (fixture: [`images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx) via `xl/_rels/metadata.xml.rels`)
    - `http://schemas.microsoft.com/office/2017/relationships/richValueRel` (fixture: [`images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx) via `xl/_rels/metadata.xml.rels`)
    - `http://schemas.microsoft.com/office/2017/relationships/richValueTypes` (fixture: [`images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx) via `xl/_rels/metadata.xml.rels`)
    - `http://schemas.microsoft.com/office/2017/relationships/richValueStructure` (fixture: [`images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx) via `xl/_rels/metadata.xml.rels`)
    - `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue` (fixture: [`image-in-cell.xlsx`](../fixtures/xlsx/basic/image-in-cell.xlsx); also test: `embedded_images_place_in_cell_roundtrip.rs`)
    - `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure` (fixture: [`image-in-cell.xlsx`](../fixtures/xlsx/basic/image-in-cell.xlsx))
    - `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes` (fixture: [`image-in-cell.xlsx`](../fixtures/xlsx/basic/image-in-cell.xlsx))
    - `http://schemas.microsoft.com/office/2022/10/relationships/richValueRel` (fixture: [`image-in-cell.xlsx`](../fixtures/xlsx/basic/image-in-cell.xlsx); also test: `embedded_images_place_in_cell_roundtrip.rs`)
  - (Exact `richValue` schemas and relationship discovery still vary; preserve unknown relationships.)

## TODO: verify with real Excel sample

This doc is partially derived from **in-repo fixtures** (some synthetic, some Excel-generated) plus
best-effort reverse engineering.

Current status:

- **Confirmed** for real Excel-generated “Place in Cell” workbooks in this repo:
  - RichData-only wiring (no `xl/cellimages.xml`):
    - [`fixtures/xlsx/basic/image-in-cell.xlsx`](../fixtures/xlsx/basic/image-in-cell.xlsx) (notes in [`fixtures/xlsx/basic/image-in-cell.md`](../fixtures/xlsx/basic/image-in-cell.md))
    - error cell `t="e"` / cached `#VALUE!` + `vm="…"` → `xl/metadata.xml` → `xl/richData/rd*` → `richValueRel*` → `xl/media/*`
  - RichData + cell image store wiring (`xl/cellimages.xml` present):
    - [`fixtures/xlsx/rich-data/images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx) (notes in [`fixtures/xlsx/rich-data/images-in-cell.md`](../fixtures/xlsx/rich-data/images-in-cell.md))
    - cell `A1` uses `vm="1" cm="1"` with cached `<v>0</v>` (image binding still comes from `vm` + `xl/metadata.xml`)
- ✅ Real Excel-generated workbook that includes an `_xlfn.IMAGE(...)` formula cell:
  - [`fixtures/xlsx/images-in-cells/image-in-cell.xlsx`](../fixtures/xlsx/images-in-cells/image-in-cell.xlsx)
- Still recommended: additional real Excel-generated `IMAGE()` fixtures without external network dependencies
  (e.g. `file://`), to confirm whether Excel embeds/caches image bytes differently across builds.

Before we hardcode any remaining assumptions, validate them against a real Excel-generated workbook that uses both:

- Insert → Pictures → **Place in Cell**
- a formula cell containing `=IMAGE(...)`

Checklist (what we still want additional fixtures for):

1. Confirm whether Excel ever uses numbered parts like `cellImages1.xml` / `cellimages1.xml` (not yet observed in this repo).
2. Discover additional `cellImages` namespace variants across Excel versions:
   - observed: `http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages` (real Excel fixtures:
     `fixtures/xlsx/rich-data/images-in-cell.xlsx`, `fixtures/xlsx/images-in-cells/image-in-cell.xlsx`)
   - observed: `http://schemas.microsoft.com/office/spreadsheetml/2023/02/main` (synthetic fixture `cell-images.xlsx`)
3. Discover additional XML shapes:
   - observed: `<cellImages><cellImage><xdr:pic>...` (real Excel fixture `fixtures/xlsx/rich-data/images-in-cell.xlsx`)
   - observed: `<cellImages><xdr:pic>...` (real Excel fixture `fixtures/xlsx/images-in-cells/image-in-cell.xlsx`)
4. Confirm additional `[Content_Types].xml` override strings (observed to vary across producers/fixtures):
   - observed: `application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml` (real Excel fixture `fixtures/xlsx/rich-data/images-in-cell.xlsx`)
   - observed: `application/vnd.ms-excel.cellimages+xml` (real Excel fixture `fixtures/xlsx/images-in-cells/image-in-cell.xlsx`)
5. Discover additional relationship Type URIs and owning-part behaviors:
    - observed workbook → cellimages: `http://schemas.microsoft.com/office/2019/relationships/cellimages`
      - real Excel fixture: `fixtures/xlsx/rich-data/images-in-cell.xlsx`
    - observed workbook → cellimages: `http://schemas.microsoft.com/office/2017/06/relationships/cellImages`
      - real Excel fixture: `fixtures/xlsx/images-in-cells/image-in-cell.xlsx`
    - observed workbook → cellImages: `http://schemas.microsoft.com/office/2023/02/relationships/cellImage` (synthetic fixture)
     - still unknown: worksheet-level relationship variants (if any)
6. Confirm how `=IMAGE(...)` worksheet cells are encoded across Excel versions and input types (URL vs local file).
   - Observed in the real Excel fixture `fixtures/xlsx/images-in-cells/image-in-cell.xlsx`:
     - formula cell: `<f>_xlfn.IMAGE("...")</f>`
     - `c/@vm` is present and the cached value is numeric (`<v>0</v>`, not `t="e"` / `#VALUE!`).

## Round-trip constraints for Formula

Until Formula implements a full semantic model for images-in-cells, the compatibility requirement is:

1. **Preserve the parts (if present)**:
   - `xl/cellimages.xml` (or `xl/cellImages.xml`)
   - `xl/_rels/cellimages.xml.rels` (or `xl/_rels/cellImages.xml.rels`)
   - `xl/media/*` images referenced by those relationships
   - `xl/metadata.xml`
   - `xl/richData/*` (and `xl/richData/_rels/*`)
2. **Preserve worksheet references**:
   - `c/@vm` and `c/@cm`
   - any `<extLst>` content in cells/worksheets relevant to rich values
3. **Preserve `[Content_Types].xml` and `*.rels`** entries for all of the above.

This is the minimum needed so that:

- opening an Excel workbook with images-in-cells,
- editing unrelated values, and
- saving back to `.xlsx`

does not “orphan” images or break Excel’s internal references.

## Status in Formula

### Implemented / covered by tests today

- **`xl/cellimages.xml` / `xl/cellImages.xml` parsing (workbook-level) + media import**
  - Parser: `crates/formula-xlsx/src/cell_images/mod.rs`
  - Test: `crates/formula-xlsx/tests/cell_images.rs`
- **Best-effort image import during `XlsxDocument` load**
  - `crates/formula-xlsx/src/read/mod.rs` calls `CellImages::parse_from_parts(...)` (best-effort) to populate `workbook.images`.
- **Preservation of `xl/cellimages.xml` / `xl/cellImages.xml` + matching `.rels` + `xl/media/*` on cell edits**
  - Test: `crates/formula-xlsx/tests/cellimages_preservation.rs`
- **Round-trip preservation of richData-backed in-cell image parts**
  - Test: `crates/formula-xlsx/tests/rich_data_roundtrip.rs`
  - Fixture: `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` (synthetic)
- **Round-trip preservation of a real Excel in-cell image workbook with `xl/cellimages.xml` + full `richValue*` tables**
  - Test: `crates/formula-xlsx/tests/real_excel_images_in_cell.rs`
  - Fixture: `fixtures/xlsx/rich-data/images-in-cell.xlsx`
- **Preservation of richData parts when related from `xl/metadata.xml` (via `xl/_rels/metadata.xml.rels`)**
  - Test: `crates/formula-xlsx/tests/richdata_preservation.rs`
  - Fixture: `fixtures/xlsx/rich-data/richdata-minimal.xlsx` (synthetic)
- **Preservation of RichData “Place in Cell” parts (`xl/metadata.xml` + `xl/richData/*` + `xl/media/*`) on edits**
  - Test: `crates/formula-xlsx/tests/embedded_images_place_in_cell_roundtrip.rs`
- **Best-effort extraction of richData-backed in-cell images (cell → bytes)**
  - API: `crates/formula-xlsx/src/rich_data/mod.rs` (`extract_rich_cell_images`)
  - Test: `crates/formula-xlsx/tests/rich_data_cell_images.rs`
- **`vm` attribute preservation** for untouched cells is covered by:
  - `crates/formula-xlsx/tests/sheetdata_row_col_attrs.rs` (`editing_a_cell_does_not_strip_unrelated_row_col_or_cell_attrs`)
  - `crates/formula-xlsx/tests/metadata_rich_values_vm_roundtrip.rs` (also asserts `xl/metadata.xml` is preserved and the workbook relationship to `metadata.xml` remains)
- **`vm` attribute dropping** (when patching a cell value/formula away from rich-value placeholder semantics) is covered by:
  - `crates/formula-xlsx/tests/cell_metadata_preservation.rs`
  - `crates/formula-xlsx/tests/metadata_rich_value_roundtrip.rs`
- **`cm` + `<extLst>` preservation** during cell patching is covered by:
  - `crates/formula-xlsx/tests/cell_metadata_preservation.rs`
- **Best-effort `xl/metadata.xml` parsing for rich values (`vm` -> richValue index)**
  - `crates/formula-xlsx/src/rich_data/metadata.rs`
- **SpreadsheetML `xl/metadata.xml` parser (opaque-preserving)**
  - `crates/formula-xlsx/src/metadata.rs` parses the core `<metadataTypes>` + `<cellMetadata>` / `<valueMetadata>`
    `<rc>` records and preserves `<futureMetadata>` `<bk>` payloads as raw inner XML for inspection/debugging.
- **`_xlfn.` prefix handling** exists in:
  - `crates/formula-xlsx/src/formula_text.rs`
  - includes an explicit `IMAGE()` round-trip test (`xlfn_roundtrip_preserves_image_function`)

Limitations (current Formula behavior):

- Formula can **load** the image bytes referenced by the `cellImages` part (`xl/cellimages.xml` / `xl/cellImages.xml`) into `workbook.images`
  during `XlsxDocument` load.
- For richData-backed images, Formula has best-effort extractors (see above), but the main `formula-model` cell value
  layer does not yet represent an image-in-cell value as a first-class `CellValue` variant (and this doc intentionally
  does not cover UI rendering).

### TODO work (required for images-in-cells)

- Fixture coverage status:
  - ✅ Real Excel “Place in Cell” fixture (richData-backed, no `xl/cellimages*.xml`):
    - [`fixtures/xlsx/basic/image-in-cell.xlsx`](../fixtures/xlsx/basic/image-in-cell.xlsx)
  - ✅ Real Excel “Place in Cell” fixture (includes `xl/cellimages.xml` + full `richValue*` tables):
    - [`fixtures/xlsx/rich-data/images-in-cell.xlsx`](../fixtures/xlsx/rich-data/images-in-cell.xlsx)
  - ✅ Real Excel fixture that includes an `_xlfn.IMAGE(...)` cell:
    - [`fixtures/xlsx/images-in-cells/image-in-cell.xlsx`](../fixtures/xlsx/images-in-cells/image-in-cell.xlsx)
  - **Confirm and document remaining relationship/content-type variants** from additional real Excel fixtures (especially `=IMAGE(...)`):
    - `[Content_Types].xml` overrides for:
      - `/xl/metadata.xml`
      - `/xl/richData/*.xml` (especially `/xl/richData/richValue.xml`)
  - the relationship Type URIs (if any) that connect the workbook/worksheets to:
    - `xl/cellimages.xml` / `xl/cellImages.xml`
    - `xl/metadata.xml`
    - `xl/richData/*`
- **Rich-value semantics (beyond preservation)**:
  - parse enough of `xl/richData/richValue.xml` to connect `vm` → richValue record → cell image entry
  - (still out of scope here: UI rendering)
