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
   - Stored via a workbook-level **cell image store** plus the **rich value / metadata** system.
   - The rest of this document focuses on this second mechanism.

## Expected OOXML parts

Workbooks using images-in-cells are expected to include some/all of the following parts:

```
xl/
├── cellimages.xml
├── _rels/
│   └── cellimages.xml.rels
├── media/
│   └── image*.{png,jpg,gif,...}
├── metadata.xml
└── richData/
    ├── richValue.xml
    ├── richValueRel.xml
    ├── richValueTypes.xml
    ├── richValueStructure.xml
    └── _rels/
        └── richValueRel.xml.rels
```

Notes:

- `xl/media/*` contains the actual image bytes (usually `.png`, but Excel may use other formats).
- The exact `xl/richData/*` file set can vary across Excel builds; the `richValue*` names shown above are
  common, but Formula should preserve the entire `xl/richData/` directory byte-for-byte unless we
  explicitly implement rich-value editing.
- `xl/metadata.xml` and the per-cell `c/@vm` + `c/@cm` attributes connect worksheet cells to the rich
  value system.

See also: [20-xlsx-richdata-images-in-cell.md](./20-xlsx-richdata-images-in-cell.md) for a deeper (still
best-effort) description of the `richValue*` part set and how `richValueRel.xml` is used to resolve
media relationships.

## Worksheet cell references (`c/@vm`, `c/@cm`, `<extLst>`)

SpreadsheetML’s `<c>` (cell) element can carry metadata indices:

- `c/@vm` — **value metadata index** (used to associate a cell’s *value* with a record in `xl/metadata.xml`)
- `c/@cm` — **cell metadata index** (used to associate the *cell* with a record in `xl/metadata.xml`)

For round-trip safety, Formula must preserve these attributes even when the value/formula changes,
because they can “point” to image/rich-value structures elsewhere in the package.

### Minimal examples (from existing Formula fixtures/tests)

The repository already has fixtures/tests exercising preservation of these attributes:

```xml
<!-- `vm` attribute example (fixtures/xlsx/basic/row-col-attrs.xlsx) -->
<row r="2" spans="1:1" ht="20" customHeight="1">
  <c r="A2" vm="1"><v>2</v></c>
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

### Images-in-cells cell shape (representative; confirm with fixture)

Cells containing an image-in-cell (either via `=IMAGE(...)` or a placed-in-cell picture) are expected to
use `vm`/`cm` and/or an `<extLst>` to reference workbook-level rich value/image tables.

Representative shape (exact details TBD; do not treat this as authoritative until a real Excel fixture is
checked in):

```xml
<c r="A1" vm="1" cm="7">
  <f>_xlfn.IMAGE("https://example.com/cat.png")</f>
  <v>0</v>
  <extLst>...</extLst>
</c>
```

**Round-trip rule:** `vm`, `cm`, and the entire `<extLst>` subtree must be preserved verbatim unless we
explicitly implement full rich-value editing.

### How `vm` maps to `xl/metadata.xml` and `xl/richData/richValue.xml`

Formula’s current understanding (implemented in `crates/formula-xlsx/src/rich_data/metadata.rs`) is:

1. Worksheet cells reference a *value metadata record* via `c/@vm` (**1-based**).
2. `xl/metadata.xml` contains `<valueMetadata>` with a list of `<bk>` records; `vm` selects a `<bk>`.
3. That `<bk>` contains `<rc t="…" v="…"/>` where:
   - `t` is the **1-based** index of `"XLRICHVALUE"` inside `<metadataTypes>`.
   - `v` is a **0-based** index into `<futureMetadata name="XLRICHVALUE">`’s `<bk>` list.
4. That future-metadata `<bk>` contains an extension element (commonly `xlrd:rvb`) with an `i="…"`
   attribute, which is a **0-based** index into `xl/richData/richValue.xml`.

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

## `xl/cellimages.xml`

`xl/cellimages.xml` is the workbook-level “cell image store” part. It is expected to contain a list of
image entries that can be referenced (directly or indirectly) by rich values.

The part embeds **SpreadsheetDrawing / DrawingML** `<xdr:pic>` payloads and uses
`<a:blip r:embed="rId…">` to reference an image relationship in `xl/_rels/cellimages.xml.rels`.

Observed root namespaces (from in-repo tests; Excel versions may vary):

- `http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages`
- `http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages`

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

### `xl/_rels/cellimages.xml.rels`

`xl/_rels/cellimages.xml.rels` contains OPC relationships from `cellimages.xml` to the binary image parts
under `xl/media/*`.

This relationships file is standard OPC, and the **image relationship type URI is known**:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png"/>
</Relationships>
```

**Round-trip rules:**

- Preserve `Relationship/@Id` values.
- Preserve `Target` paths and file names (Excel reuses these paths across features).
- Preserve the referenced `xl/media/*` bytes byte-for-byte.

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
For images-in-cells, these rich values are expected to contain (directly or indirectly) a reference to:

- an entry in `xl/cellimages.xml`, and therefore
- an image binary in `xl/media/*`.

At minimum, `xl/richData/richValue.xml` is expected to exist when `xl/metadata.xml` contains
`<futureMetadata name="XLRICHVALUE">` entries with `xlrd:rvb i="…"` references (see mapping details above).

Because the exact file set and schemas vary across Excel builds, Formula’s short-term strategy is:

- **preserve all `xl/richData/*` parts and their `*.rels`**, and
- treat them as an **atomic bundle** with `xl/metadata.xml` + `xl/cellimages.xml` during round-trip.

Common file names (Excel version-dependent; treat as “expected shape”, not a strict schema):

- `xl/richData/richValue.xml`
- `xl/richData/richValueRel.xml` (+ `xl/richData/_rels/richValueRel.xml.rels`)
- `xl/richData/richValueTypes.xml`
- `xl/richData/richValueStructure.xml`

## `[Content_Types].xml` requirements

Workbooks that include these parts must also declare content types in `[Content_Types].xml`:

- **Override** entries for XML parts like `/xl/cellimages.xml`, `/xl/metadata.xml`, and `xl/richData/*.xml`
- **Default** entries for image extensions used under `/xl/media/*` (`png`, `jpg`, `gif`, etc.)

### `xl/cellimages.xml` content type override

Observed values (from in-repo tests; preserve whatever is in the source workbook):

- `application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml`
  - used by `crates/formula-xlsx/tests/cell_images.rs`
- `application/vnd.ms-excel.cellimages+xml`
  - used by `crates/formula-xlsx/tests/cellimages_preservation.rs`

Excel uses Microsoft-specific content type strings for this part, and the exact string may vary across
versions/builds.

**Round-trip rule:** treat any `<Override PartName="/xl/cellimages.xml" .../>` as authoritative and
preserve its `ContentType` value byte-for-byte (do not hardcode a single MIME type in the writer).

### Other content types (TODO: fixture-driven)

Content types for `xl/metadata.xml` and `xl/richData/*` still need confirmation from a real Excel-exported
workbook (and corresponding fixture + tests).

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- ... -->
  <Default Extension="png" ContentType="image/png"/>
  <!-- ... -->

  <Override PartName="/xl/cellimages.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml"/>

  <!-- TODO: confirm these from an Excel fixture -->
  <Override PartName="/xl/metadata.xml" ContentType="TODO"/>
  <Override PartName="/xl/richData/richValue.xml" ContentType="TODO"/>
</Types>
```

**TODO (fixture-driven):** add an Excel-generated workbook using real images-in-cells, then update this
section with the exact `ContentType="..."` strings for `metadata.xml` and `richData/*`.

## Relationship type URIs (what we know vs TODO)

Known (stable, used across OOXML):

- Image relationships (used by DrawingML and expected to be used by `cellimages.xml`):
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`

Partially known (fixture-driven details still recommended):

- Workbook → `xl/cellimages.xml` relationship:
  - Lives in `xl/_rels/workbook.xml.rels`.
  - Excel uses a Microsoft-extension relationship `Type` URI that has been observed to vary.
  - **Round-trip / detection rule:** identify the relationship by resolved `Target` (`/xl/cellimages.xml`)
    rather than hardcoding a single `Type`.

TODO (confirm via real Excel fixture, then harden parsers/writers):

- Relationship type(s) connecting `xl/workbook.xml` (or other workbook-level parts) to:
  - `xl/metadata.xml`
  - `xl/richData/*`
  - `xl/cellimages.xml`

Until confirmed, Formula must preserve any such relationships byte-for-byte rather than regenerating.

## Round-trip constraints for Formula

Until Formula implements a full semantic model for images-in-cells, the compatibility requirement is:

1. **Preserve the parts**:
   - `xl/cellimages.xml`
   - `xl/_rels/cellimages.xml.rels`
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

- **`xl/cellimages.xml` parsing (workbook-level) + media import**
  - Parser: `crates/formula-xlsx/src/cell_images/mod.rs`
  - Test: `crates/formula-xlsx/tests/cell_images.rs`
- **Best-effort image import during `XlsxDocument` load**
  - `crates/formula-xlsx/src/read/mod.rs` calls `load_cell_images_from_parts(...)` to populate `workbook.images`.
- **Preservation of `xl/cellimages.xml` + `xl/_rels/cellimages.xml.rels` + `xl/media/*` on cell edits**
  - Test: `crates/formula-xlsx/tests/cellimages_preservation.rs`
- **`vm` attribute preservation** on edit is covered by:
  - `crates/formula-xlsx/tests/sheetdata_row_col_attrs.rs` (`editing_a_cell_does_not_strip_unrelated_row_col_or_cell_attrs`)
- **`cm` + `<extLst>` preservation** during cell patching is covered by:
  - `crates/formula-xlsx/tests/cell_metadata_preservation.rs`
- **Best-effort `xl/metadata.xml` parsing for rich values (`vm` -> richValue index)**
  - `crates/formula-xlsx/src/rich_data/metadata.rs`
- **`_xlfn.` prefix handling** exists in:
  - `crates/formula-xlsx/src/formula_text.rs`
  - includes an explicit `IMAGE()` round-trip test (`xlfn_roundtrip_preserves_image_function`)

### TODO work (required for images-in-cells)

- **Add a real Excel-generated fixture workbook** covering:
  - a “Place in Cell” inserted image
  - a formula cell containing `=IMAGE(...)`
  - and the accompanying `xl/metadata.xml` + `xl/richData/*` parts
- **Confirm and document the remaining relationship/content-type details** from that fixture:
  - `[Content_Types].xml` overrides for:
    - `/xl/metadata.xml`
    - `/xl/richData/*.xml` (especially `/xl/richData/richValue.xml`)
  - the relationship Type URIs (if any) that connect the workbook/worksheets to:
    - `xl/cellimages.xml`
    - `xl/metadata.xml`
    - `xl/richData/*`
- **Rich-value semantics (beyond preservation)**:
  - parse enough of `xl/richData/richValue.xml` to connect `vm` → richValue record → cell image entry
  - (still out of scope here: UI rendering)
