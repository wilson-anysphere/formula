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
    ├── *.xml
    └── _rels/
        └── *.rels
```

Notes:

- `xl/media/*` contains the actual image bytes (usually `.png`, but Excel may use other formats).
- `xl/richData/*` contains “rich value” tables used by Excel for non-primitive cell values (data types,
  entities, and images-in-cells).
- `xl/metadata.xml` and the per-cell `c/@vm` + `c/@cm` attributes connect worksheet cells to the rich
  value system.

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

## `xl/cellimages.xml`

`xl/cellimages.xml` is the workbook-level “cell image store” part. It is expected to contain a list of
image entries that can be referenced (directly or indirectly) by rich values.

**Schema note:** the exact XML namespace and element vocabulary is Excel-specific and must be validated
against a real fixture before we implement a semantic model.

Representative skeleton:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="TODO:excel-cellimages-namespace"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:id="rId1"/>
  <!-- ... -->
</cellImages>
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

Because the exact file set and schemas vary across Excel builds, Formula’s short-term strategy is:

- **preserve all `xl/richData/*` parts and their `*.rels`**, and
- treat them as an **atomic bundle** with `xl/metadata.xml` + `xl/cellimages.xml` during round-trip.

Example file names seen in the ecosystem (verify with a checked-in fixture before hard-coding):

- `xl/richData/richValueTypes.xml`
- `xl/richData/richValueStructures.xml`
- `xl/richData/richValueData.xml`
- plus potential supporting parts and relationship files under `xl/richData/_rels/`

## `[Content_Types].xml` requirements

Workbooks that include these parts must also declare content types in `[Content_Types].xml`:

- **Override** entries for XML parts like `/xl/cellimages.xml`, `/xl/metadata.xml`, and `xl/richData/*.xml`
- **Default** entries for image extensions used under `/xl/media/*` (`png`, `jpg`, `gif`, etc.)

Representative (content types for these newer parts are **TODO** until we have an Excel fixture):

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- ... -->
  <Default Extension="png" ContentType="image/png"/>
  <!-- ... -->

  <Override PartName="/xl/cellimages.xml" ContentType="TODO"/>
  <Override PartName="/xl/metadata.xml" ContentType="TODO"/>
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="TODO"/>
  <Override PartName="/xl/richData/richValueStructures.xml" ContentType="TODO"/>
  <Override PartName="/xl/richData/richValueData.xml" ContentType="TODO"/>
</Types>
```

**TODO (fixture-driven):** confirm the exact `ContentType="..."` strings produced by current Excel builds
and add a fixture + assertion coverage in `crates/formula-xlsx`.

## Relationship type URIs (what we know vs TODO)

Known (stable, used across OOXML):

- Image relationships (used by DrawingML and expected to be used by `cellimages.xml`):
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`

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

- **`vm` attribute preservation** on edit is covered by:
  - `crates/formula-xlsx/tests/sheetdata_row_col_attrs.rs` (`editing_a_cell_does_not_strip_unrelated_row_col_or_cell_attrs`)
- **`cm` + `<extLst>` preservation** during cell patching is covered by:
  - `crates/formula-xlsx/tests/cell_metadata_preservation.rs`
- **`_xlfn.` prefix handling** exists in:
  - `crates/formula-xlsx/src/formula_text.rs`

### TODO work (required for images-in-cells)

- **Add a `cellimages.xml` parser/preserver task**
  - Parse and preserve `xl/cellimages.xml` + `xl/_rels/cellimages.xml.rels` + referenced `xl/media/*`.
  - Add a fixture workbook that includes at least one placed-in-cell image and one `IMAGE()` result.
  - Add round-trip assertions for:
    - relationship IDs
    - `[Content_Types].xml` overrides
    - worksheet `vm`/`cm` attributes
- **Add `_xlfn.IMAGE` to the `_xlfn.` prefix-required function set**
  - Excel stores newer functions with `_xlfn.` in OOXML; `IMAGE` must be included for correct save
    compatibility.
  - Update `XL_FN_REQUIRED_FUNCTIONS` in `crates/formula-xlsx/src/formula_text.rs` once validated.

