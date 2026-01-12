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
   - Stored via the workbook-level **rich value / metadata** system, and may additionally use a
     dedicated **cell image store** part (`xl/cellImages.xml`) depending on producer/version.
   - The rest of this document focuses on this second mechanism.

## Expected OOXML parts

Workbooks using images-in-cells are expected to include some/all of the following parts:

```
xl/
├── cellImages.xml                # Optional (some files use richData-only wiring; preserve if present)
├── cellImages1.xml               # Optional; allow numeric suffixes like other indexed XLSX parts
├── media/
│   └── image*.{png,jpg,gif,...}
├── metadata.xml
├── _rels/
│   ├── cellImages.xml.rels       # Optional (only if a cellImages.xml part exists)
│   └── metadata.xml.rels (commonly present when `metadata.xml` references `xl/richData/*`)
└── richData/
    ├── richValue.xml
    ├── richValueRel.xml
    ├── richValueTypes.xml
    ├── richValueStructure.xml
    └── _rels/
        └── richValueRel.xml.rels
```

Notes:

- **Part-name casing:** Some producers (and our current fixtures/tests) use `xl/cellimages.xml` (all
  lowercase) instead of `xl/cellImages.xml`. OPC part names are case-sensitive inside the ZIP, so:
  - readers should handle both variants, and
  - writers should preserve the original casing when round-tripping an existing file.
- **`cellImages.xml` may be absent:** Some workbooks store image-in-cell values entirely via
  `xl/metadata.xml` + `xl/richData/*` (especially `richValueRel.xml` → `.rels` → `xl/media/*`) without a
  separate `xl/cellImages.xml` part. See `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` and
  `crates/formula-xlsx/tests/rich_data_roundtrip.rs`.
  - If `cellImages.xml` exists, preserve it and its relationship graph.
- `xl/media/*` contains the actual image bytes (usually `.png`, but Excel may use other formats).
- The exact `xl/richData/*` file set can vary across Excel builds; the `richValue*` names shown above are
  common, but Formula should preserve the entire `xl/richData/` directory byte-for-byte unless we
  explicitly implement rich-value editing.
- `xl/metadata.xml` and the per-cell `c/@vm` + `c/@cm` attributes connect worksheet cells to the rich
  value system.
- When present, `xl/_rels/metadata.xml.rels` typically connects `xl/metadata.xml` → `xl/richData/*` parts.
  Formula should preserve these relationships byte-for-byte for safe round-trips.

See also: [20-images-in-cells-richdata.md](./20-images-in-cells-richdata.md) for a deeper (still
best-effort) description of the `richValue*` part set and how `richValueRel.xml` is used to resolve
media relationships.

For a concrete, fixture-backed “Place in Cell” schema walkthrough (including the `rdrichvalue*` keys
`_rvRel:LocalImageIdentifier` and `CalcOrigin`), see:

- [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md)
  - The Excel-produced fixture `fixtures/xlsx/basic/image-in-cell.xlsx` uses the same richData-only wiring
    (and does **not** use `xl/cellImages.xml`), with notes in `fixtures/xlsx/basic/image-in-cell.md`.

## In-repo fixture (cell image store part)

Fixture workbook: `fixtures/xlsx/basic/cell-images.xlsx`

Confirmed values from this fixture:

- Part paths:
  - `xl/cellImages.xml`
  - `xl/_rels/cellImages.xml.rels`
  - `xl/media/image1.png`
- `xl/cellImages.xml` root namespace:
  - `http://schemas.microsoft.com/office/spreadsheetml/2023/02/main`
- `[Content_Types].xml` override for the part:
  - `application/vnd.ms-excel.cellimages+xml`
- Workbook → `cellImages.xml` relationship `Type` URI (in `xl/_rels/workbook.xml.rels`):
  - `http://schemas.microsoft.com/office/2023/02/relationships/cellImage`

### Quick reference: `cellImages` part graph (OPC + XML)

This section is a “what to look for” summary for the core **cell image store** parts. Details and
variant shapes are documented further below.

#### Confirmed vs unconfirmed

**Confirmed (from in-repo fixtures/tests):**

- A workbook can contain a dedicated `cellImages` part (seen in tests as `xl/cellimages.xml` and
  `xl/cellImages.xml`) plus a matching relationship part at `xl/_rels/<part>.rels`.
- `fixtures/xlsx/basic/cell-images.xlsx` contains `xl/cellImages.xml` with namespace:
  - `http://schemas.microsoft.com/office/spreadsheetml/2023/02/main`
- The `cellImages` XML can reference binary images via DrawingML-style `r:embed="rIdX"` references.
- `rIdX` is resolved through the `*.rels` part to an image under `xl/media/*`.
- Image relationship type is the standard OOXML one:
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`

**Unconfirmed / needs real Excel sample (Place in Cell / `IMAGE()`):**

- Exact part naming + casing used by current Excel builds (and whether multiple numbered parts like
  `cellImages1.xml` are used).
- Exact root namespace used by Excel for `cellImages` today (we *expect* `.../2019/cellimages`, but
  have seen other variants in synthetic fixtures).
- Exact schema shape (e.g. whether `<cellImage>` always contains a full `<xdr:pic>` subtree or can be
  a lightweight reference-only element).
- Whether Excel consistently uses a single relationship `Type` URI (and whether the relationship is
  always on `xl/workbook.xml.rels` vs sometimes worksheet-level).
- The exact “cell → image” mapping mechanism across **all** Excel scenarios.
  - Confirmed for a rust_xlsxwriter-generated **“Place in Cell”** workbook (used for schema verification in this repo):
    - worksheet cell is `t="e"` with cached `#VALUE!` and `vm="1"`
    - image bytes are resolved via `xl/metadata.xml` + `xl/richData/rd*` + `xl/richData/richValueRel.xml(.rels)` → `xl/media/*`
    - no `xl/cellImages.xml`/`xl/cellimages.xml` part is used in that case
    - see: [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md)
  - Still an open question for real Excel-generated `IMAGE()` results and other producers (where `xl/cellImages*.xml` may appear).

#### Parts

- `xl/cellImages.xml` (**preferred**, but casing can vary; see note above)
- `xl/_rels/cellImages.xml.rels`
- image binaries: `xl/media/imageN.<ext>`

#### XML namespace + structure (likely)

- Root element local name: `<cellImages>` in namespace:
  - `http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages`
- Some files may use newer versions like:
  - `http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages`
- Some in-repo synthetic fixtures also use a more generic Microsoft SpreadsheetML namespace:
  - `http://schemas.microsoft.com/office/spreadsheetml/2020/07/main` (unverified vs real Excel)
- `fixtures/xlsx/basic/cell-images.xlsx` uses:
  - `http://schemas.microsoft.com/office/spreadsheetml/2023/02/main`
- The root contains one or more `<cellImage>` entries, each containing a DrawingML picture subtree
  (typically `<xdr:pic>`) and a blip like:
  - `<a:blip r:embed="rIdX"/>`
- `r:embed="rIdX"` is resolved via `xl/_rels/cellImages.xml.rels` to a `Target` under `xl/media/*`.

#### Content types (expected)

- `[Content_Types].xml` override (expected, but verify against real Excel):

```xml
<Override PartName="/xl/cellImages.xml"
          ContentType="application/vnd.ms-excel.cellimages+xml"/>
```

#### Relationship types

- Image relationship type (standard OOXML):
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`
- Relationship type for “workbook/worksheet → `cellImages.xml`” discovery:
  - Confirmed in `fixtures/xlsx/basic/cell-images.xlsx`:
    - `http://schemas.microsoft.com/office/2023/02/relationships/cellImage`
  - Other candidates observed in synthetic fixtures / tooling:
    - `http://schemas.microsoft.com/office/2020/07/relationships/cellImages`
    - `http://schemas.microsoft.com/office/2022/relationships/cellImages`
  - Candidate observed in a synthetic round-trip test:
    - `http://schemas.microsoft.com/office/2020/07/relationships/cellImages`
  - Candidate observed in synthetic fixtures / corpus tooling:
    - `http://schemas.microsoft.com/office/2022/relationships/cellImages`

#### Minimal example (`xl/cellImages.xml`) (synthetic)

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

#### Minimal example (`xl/_rels/cellImages.xml.rels`) (synthetic)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png"/>
</Relationships>
```

#### Minimal example (`xl/_rels/workbook.xml.rels` entry) (fixture)

Some files link `xl/workbook.xml` → `xl/cellImages.xml` via an OPC relationship in
`xl/_rels/workbook.xml.rels` using a Microsoft-specific relationship `Type`.

```xml
<Relationship Id="rId3"
              Type="http://schemas.microsoft.com/office/2023/02/relationships/cellImage"
              Target="cellImages.xml"/>
```

## Worksheet cell references (`c/@vm`, `c/@cm`, `<extLst>`)

SpreadsheetML’s `<c>` (cell) element can carry metadata indices:

- `c/@vm` — **value metadata index** (used to associate a cell’s *value* with a record in `xl/metadata.xml`)
- `c/@cm` — **cell metadata index** (used to associate the *cell* with a record in `xl/metadata.xml`)

For round-trip safety, Formula must preserve these attributes even when the value/formula changes,
because they can “point” to image/rich-value structures elsewhere in the package.

### Minimal examples (from existing Formula fixtures/tests)

The repository already has fixtures/tests exercising preservation of these attributes:

```xml
<!-- `vm` attribute example (fixtures/xlsx/metadata/rich-values-vm.xlsx) -->
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

1. Worksheet cells reference a *value metadata record* via `c/@vm`.
   - Excel commonly emits `vm` as **1-based**, but **0-based** values are also observed in the wild (and in our
     test-only richData fixtures). Treat `vm` as opaque and preserve it.
2. `xl/metadata.xml` contains `<valueMetadata>` with a list of `<bk>` records; `vm` selects a `<bk>`.
3. That `<bk>` contains `<rc t="…" v="…"/>` where:
    - `t` is the **1-based** index of `"XLRICHVALUE"` inside `<metadataTypes>`.
    - `v` is **0-based** (in this schema it indexes into `<futureMetadata name="XLRICHVALUE">`’s `<bk>` list; other schemas may use `v` differently).
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

Other observed `xl/metadata.xml` shapes exist. For example, `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`
contains a `<valueMetadata>` table but **no** `<futureMetadata>` block; Formula currently treats these
schemas as opaque and focuses on round-trip preservation (with best-effort extraction utilities in
`crates/formula-xlsx/src/rich_data/mod.rs`).

## `xl/cellImages.xml` (a.k.a. `xl/cellimages.xml`)

`xl/cellImages.xml` is the workbook-level “cell image store” part. It is expected to contain a list of
image entries that can be referenced (directly or indirectly) by rich values.

Note: not all images-in-cell workbooks include this part (some use `xl/metadata.xml` + `xl/richData/*`
only). When present, it must be preserved along with its `.rels` and referenced media.

The part embeds **SpreadsheetDrawing / DrawingML** `<xdr:pic>` payloads and uses
`<a:blip r:embed="rId…">` to reference an image relationship in `xl/_rels/cellImages.xml.rels`.

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
has explicit support for `r:id` on `<cellImage>`:

```xml
<etc:cellImages xmlns:etc="http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <etc:cellImage r:id="rId1"/>
</etc:cellImages>
```

### `xl/_rels/cellImages.xml.rels` (a.k.a. `xl/_rels/cellimages.xml.rels`)

`xl/_rels/cellImages.xml.rels` contains OPC relationships from `cellImages.xml` to the binary image parts
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

Targets are usually relative paths and may appear as `media/image1.png` or `../media/image1.png`
(preserve the original `Target` exactly).

**Parser resilience (Formula):** the `cell_images` parser uses a best-effort resolver that tries:

1. standard OPC resolution relative to the source part (`xl/cellimages*.xml`)
2. a fallback relative to the `.rels` part
3. a fallback that re-roots under `xl/` if the path escaped via `..`

See `crates/formula-xlsx/src/cell_images/mod.rs` (`resolve_target_best_effort`).

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
For images-in-cells, these rich values ultimately resolve to an **image binary** under `xl/media/*`,
but there appear to be multiple packaging patterns in the ecosystem:

1. **RichData → RichValueRel → media (no `cellImages.xml` part)**
   - Observed in this repo via `crates/formula-xlsx/tests/embedded_images_place_in_cell_roundtrip.rs`
     (generated with `rust_xlsxwriter::Worksheet::embed_image_with_format`).
   - The image bytes are resolved via:
     - `xl/richData/richValueRel.xml` → `xl/richData/_rels/richValueRel.xml.rels` → `xl/media/*`
2. **`cellImages.xml` “cell image store” → media**
   - Observed in this repo via `crates/formula-xlsx/tests/cell_images.rs` and related preservation tests.
   - The image bytes are resolved via:
     - `xl/cellImages.xml` → `xl/_rels/cellImages.xml.rels` → `xl/media/*`

The exact way that a worksheet cell points at a `cellImages.xml` entry (if that part is present) is
still not fully verified against a real Excel-generated “Place in Cell” workbook; treat that linkage
as **opaque** and preserve all related parts for safe round-trip.

At minimum, a rich value store is expected to exist when `xl/metadata.xml` indicates the `XLRICHVALUE`
metadata type (the exact mapping schema varies by producer/Excel build).

Depending on the producer, the rich value store may be named:

- `xl/richData/richValue*.xml` (Excel-like naming), or
- `xl/richData/rdrichvalue*.xml` / `xl/richData/rdRichValueTypes.xml` (rdRichValue naming; observed in this
  repo from `rust_xlsxwriter` output).

See also:

- [`docs/20-images-in-cells-richdata.md`](./20-images-in-cells-richdata.md) — concrete notes on the
  `richValue*` parts (types/structures/values/relationship indirection) and index bases used by Excel.

Because the exact file set and schemas vary across Excel builds, Formula’s short-term strategy is:

- **preserve all `xl/richData/*` parts and their `*.rels`**, and
- treat them as an **atomic bundle** with `xl/metadata.xml` (and `xl/cellImages.xml` if present) during round-trip.

Common file names (Excel version-dependent; treat as “expected shape”, not a strict schema):

- `xl/richData/richValue.xml`
- `xl/richData/richValueRel.xml` (+ `xl/richData/_rels/richValueRel.xml.rels`)
- `xl/richData/richValueTypes.xml`
- `xl/richData/richValueStructure.xml`

## `[Content_Types].xml` requirements

Workbooks that include these parts may declare content types in `[Content_Types].xml`.

In this repo, some fixtures rely on `<Default Extension="xml" ContentType="application/xml"/>` for
`xl/metadata.xml` and `xl/richData/*` (no explicit overrides), while others include explicit overrides.
Implementations should preserve whatever is in the source workbook.

Independently of overrides for `.xml` parts, image payloads under `xl/media/*` still require appropriate
image MIME defaults (e.g. `png` → `image/png`) for interoperability.

- **Override** entries for XML parts like `/xl/cellImages.xml` (or `/xl/cellimages.xml`), `/xl/metadata.xml`,
  and `xl/richData/*.xml`
- **Default** entries for image extensions used under `/xl/media/*` (`png`, `jpg`, `gif`, etc.)

### `xl/cellImages.xml` content type override

Observed values (from in-repo tests; preserve whatever is in the source workbook):

- `application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml`
  - used by `crates/formula-xlsx/tests/cell_images.rs`
- `application/vnd.ms-excel.cellimages+xml`
  - used by `crates/formula-xlsx/tests/cellimages_preservation.rs`

Excel uses Microsoft-specific content type strings for this part, and the exact string may vary across
versions/builds.

Note: MIME types are case-insensitive, but for round-trip safety we preserve the `ContentType` string
byte-for-byte (including its original capitalization).

**Round-trip rule:** treat any `<Override PartName="/xl/cellImages.xml" .../>` (or the lowercase variant) as
authoritative and preserve its `ContentType` value byte-for-byte.

If we ever need to synthesize this part from scratch, `application/vnd.ms-excel.cellimages+xml` is a
reasonable default (it matches Excel’s vendor-specific pattern like `...threadedcomments+xml` / `...person+xml`),
but we should still prefer the original file’s value when round-tripping.

### `xl/metadata.xml` content type override

Observed in `fixtures/xlsx/metadata/rich-values-vm.xlsx`:

- `application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml`

Also observed in tests:

- `application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml`
  - used by `crates/formula-xlsx/tests/metadata_rich_value_roundtrip.rs`

And observed that **no override** may be present (default `application/xml`), in:

- `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`
Note: some workbooks omit the override entirely and rely on the package default
`<Default Extension="xml" ContentType="application/xml"/>` (e.g. `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`).

### `xl/richData/*` content types (observed + TODO)

Content types for `xl/richData/*` vary across Excel/producers and across the two naming schemes
(`richValue*.xml` vs `rdRichValue*`).

Observed in `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`: no explicit `[Content_Types].xml` overrides
for `xl/richData/*` (falls back to the package default `application/xml`).

Observed in `fixtures/xlsx/basic/image-in-cell.xlsx` (explicit overrides present):

- `/xl/richData/richValueRel.xml`: `application/vnd.ms-excel.richvaluerel+xml`
- `/xl/richData/rdrichvalue.xml`: `application/vnd.ms-excel.rdrichvalue+xml`
- `/xl/richData/rdrichvaluestructure.xml`: `application/vnd.ms-excel.rdrichvaluestructure+xml`
- `/xl/richData/rdRichValueTypes.xml`: `application/vnd.ms-excel.rdrichvaluetypes+xml`

Likely patterns seen in the ecosystem for the *unprefixed* `richValue.xml` / `richValueTypes.xml` /
`richValueStructure.xml` parts (unverified; do not hardcode without a real Excel fixture in `fixtures/xlsx/**`):

- `application/vnd.ms-excel.richvalue+xml` (for `/xl/richData/richValue.xml`)
- `application/vnd.ms-excel.richvaluerel+xml` (for `/xl/richData/richValueRel.xml`)
- `application/vnd.ms-excel.richvaluetypes+xml` (for `/xl/richData/richValueTypes.xml`)
- `application/vnd.ms-excel.richvaluestructure+xml` (for `/xl/richData/richValueStructure.xml`)

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- ... -->
  <Default Extension="png" ContentType="image/png"/>
  <!-- ... -->

  <!-- These overrides may be absent; Excel sometimes relies on the default XML content type. -->
  <Override PartName="/xl/cellImages.xml"
             ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml"/>

  <Override PartName="/xl/metadata.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>

  <!-- TODO: confirm these from an Excel fixture -->
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="TODO"/>
  <Override PartName="/xl/richData/richValueStructure.xml" ContentType="TODO"/>
  <Override PartName="/xl/richData/richValue.xml" ContentType="TODO"/>
  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
</Types>
```

**TODO (fixture-driven):** add an Excel-generated workbook using real images-in-cells, then update this
section with the exact `ContentType="..."` strings for `richData/*`.

## Relationship type URIs (what we know vs TODO)

Known (stable, used across OOXML):

- Image relationships (used by DrawingML and expected to be used by `cellImages.xml`):
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`

Partially known (fixture-driven details still recommended):

- Workbook → `xl/metadata.xml` relationship:
  - Lives in `xl/_rels/workbook.xml.rels`.
  - Observed in `fixtures/xlsx/metadata/rich-values-vm.xlsx`:
    - `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"`
  - Observed in `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`:
    - `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"`
  - Preservation is covered by `crates/formula-xlsx/tests/metadata_rich_values_vm_roundtrip.rs`.
- Workbook → richData parts (when stored directly in workbook relationships):
  - Observed in `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`:
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue"`
      - `Target="richData/richValue.xml"`
    - `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel"`
      - `Target="richData/richValueRel.xml"`
- Workbook → `xl/cellImages.xml` relationship:
  - Lives in `xl/_rels/workbook.xml.rels`.
  - Excel uses a Microsoft-extension relationship `Type` URI that has been observed to vary.
  - Candidate observed in `crates/formula-xlsx/tests/cellimages_roundtrip_preserves_parts.rs`:
    - `Type="http://schemas.microsoft.com/office/2020/07/relationships/cellImages"`
  - Candidate observed in synthetic fixtures / corpus tooling:
    - `Type="http://schemas.microsoft.com/office/2022/relationships/cellImages"`
  - **Round-trip / detection rule:** identify the relationship by resolved `Target`
    (`/xl/cellImages.xml` or `/xl/cellimages.xml`) rather than hardcoding a single `Type`.
- RichData relationship indirection (images referenced via `richValueRel.xml`):
  - `xl/richData/_rels/richValueRel.xml.rels` is expected to contain standard image relationships:
    - `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"`
  - Observed in `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` and the unit test
    `crates/formula-xlsx/tests/rich_data_cell_images.rs`.
  - Workbook → richData relationships (Type URIs are Microsoft-specific and versioned). Observed in this repo:
    - `http://schemas.microsoft.com/office/2017/06/relationships/richValue` (fixture: `image-in-cell-richdata.xlsx`)
    - `http://schemas.microsoft.com/office/2017/06/relationships/richValueRel` (fixture: `image-in-cell-richdata.xlsx`)
    - `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue` (test: `embedded_images_place_in_cell_roundtrip.rs`)
    - `http://schemas.microsoft.com/office/2022/10/relationships/richValueRel` (test: `embedded_images_place_in_cell_roundtrip.rs`)
  - (Exact `richValue` schemas and relationship discovery still vary; preserve unknown relationships.)

TODO (confirm via real Excel fixture, then harden parsers/writers):

- Relationship type(s) connecting `xl/workbook.xml` (or other workbook-level parts) to:
  - `xl/cellImages.xml`
  - richData parts when linked indirectly via `xl/_rels/metadata.xml.rels` (instead of directly from `workbook.xml.rels`)

Until confirmed, Formula must preserve any such relationships byte-for-byte rather than regenerating.

## TODO: verify with real Excel sample

This doc is partially derived from **synthetic fixtures** in this repository plus best-effort reverse
engineering. Before we hardcode any remaining assumptions, validate them against a real Excel-generated
workbook that uses both:

- Insert → Pictures → **Place in Cell**
- a formula cell containing `=IMAGE(...)`

Checklist:

1. Confirm the canonical part name(s) and casing:
   - `xl/cellImages.xml` vs `xl/cellimages.xml` and whether numbered parts (`cellImages1.xml`) are used.
2. Confirm the `cellImages` namespace versions used by current Excel builds:
   - `.../2019/cellimages` vs `.../2022/cellimages`
3. Confirm the exact `cellImages` XML shape:
   - whether `<cellImage>` wrappers are always present
   - required attributes/elements on `<cellImage>` (if any)
4. Confirm `[Content_Types].xml` override(s) used by real Excel:
   - whether it is consistently `application/vnd.ms-excel.cellimages+xml` or varies.
5. Discover the workbook/worksheet relationship to `cellImages.xml`:
   - owning part (`xl/workbook.xml` vs per-sheet)
   - relationship `Type` URI (likely Microsoft-specific)
6. Confirm how the worksheet cell references an image:
   - the exact `vm` / `metadata.xml` / `richData/*` path and any `<extLst>` hooks used in `sheetN.xml`.

## Round-trip constraints for Formula

Until Formula implements a full semantic model for images-in-cells, the compatibility requirement is:

1. **Preserve the parts (if present)**:
   - `xl/cellImages.xml` (or `xl/cellimages.xml`)
   - `xl/_rels/cellImages.xml.rels` (or `xl/_rels/cellimages.xml.rels`)
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

- **`xl/cellImages.xml` / `xl/cellimages.xml` parsing (workbook-level) + media import**
  - Parser: `crates/formula-xlsx/src/cell_images/mod.rs`
  - Test: `crates/formula-xlsx/tests/cell_images.rs`
- **Best-effort image import during `XlsxDocument` load**
  - `crates/formula-xlsx/src/read/mod.rs` calls `load_cell_images_from_parts(...)` to populate `workbook.images`.
- **Preservation of `xl/cellImages.xml` / `xl/cellimages.xml` + matching `.rels` + `xl/media/*` on cell edits**
  - Test: `crates/formula-xlsx/tests/cellimages_preservation.rs`
- **Round-trip preservation of richData-backed in-cell image parts**
  - Test: `crates/formula-xlsx/tests/rich_data_roundtrip.rs`
  - Fixture: `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`
- **Preservation of RichData “Place in Cell” parts (`xl/metadata.xml` + `xl/richData/*` + `xl/media/*`) on edits**
  - Test: `crates/formula-xlsx/tests/embedded_images_place_in_cell_roundtrip.rs`
- **Best-effort extraction of richData-backed in-cell images (cell → bytes)**
  - API: `crates/formula-xlsx/src/rich_data/mod.rs` (`extract_rich_cell_images`)
  - Test: `crates/formula-xlsx/tests/rich_data_cell_images.rs`
- **`vm` attribute preservation** on edit is covered by:
  - `crates/formula-xlsx/tests/sheetdata_row_col_attrs.rs` (`editing_a_cell_does_not_strip_unrelated_row_col_or_cell_attrs`)
  - `crates/formula-xlsx/tests/metadata_rich_values_vm_roundtrip.rs` (also asserts `xl/metadata.xml` is preserved and the workbook relationship to `metadata.xml` remains)
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

- Formula can **load** the image bytes referenced by the `cellImages` part (`xl/cellImages.xml` / `xl/cellimages.xml`) into `workbook.images`
  during `XlsxDocument` load.
- For richData-backed images, Formula has best-effort extractors (see above), but the main `formula-model` cell value
  layer does not yet represent an image-in-cell value as a first-class `CellValue` variant (and this doc intentionally
  does not cover UI rendering).

### TODO work (required for images-in-cells)

- **Add a real Excel-generated fixture workbook** covering:
  - a “Place in Cell” inserted image that uses `xl/cellImages.xml` (if present in modern Excel builds)
  - a richData-backed image-in-cell (e.g. from `=IMAGE(...)`)
- **Confirm and document the remaining relationship/content-type details** from that fixture:
  - `[Content_Types].xml` overrides for:
    - `/xl/metadata.xml`
    - `/xl/richData/*.xml` (especially `/xl/richData/richValue.xml`)
  - the relationship Type URIs (if any) that connect the workbook/worksheets to:
    - `xl/cellImages.xml` / `xl/cellimages.xml`
    - `xl/metadata.xml`
    - `xl/richData/*`
- **Rich-value semantics (beyond preservation)**:
  - parse enough of `xl/richData/richValue.xml` to connect `vm` → richValue record → cell image entry
  - (still out of scope here: UI rendering)
