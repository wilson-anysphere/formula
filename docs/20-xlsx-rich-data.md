# XLSX Rich Data (`richData`) and “Image in Cell” Storage

## Overview

Excel historically stores images via the **drawing layer** (`xl/drawings/*`, anchored/floating shapes). Newer Excel builds (Microsoft 365) also support **“Place in Cell” / “Image in Cell”**, where an image behaves like a *cell value*.

For the dedicated “images in cells” packaging overview (including the optional `xl/cellImages.xml` part, sometimes `xl/cellimages.xml`), see:

- [`docs/20-images-in-cells.md`](./20-images-in-cells.md)
- [`docs/20-images-in-cells-richdata.md`](./20-images-in-cells-richdata.md)
- [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md) (confirmed “Place in Cell” rich-value wiring + exact URIs)
  - See also the Excel-produced fixture `fixtures/xlsx/basic/image-in-cell.xlsx` (notes in `fixtures/xlsx/basic/image-in-cell.md`).

In OOXML this is implemented using **Rich Values** (“richData”) plus **cell value metadata**:

- The worksheet cell points to a **value-metadata record** via the cell attribute `vm="…"`.
- That value-metadata record binds the cell to a **rich value index**. The exact mapping schema varies:
  - Some files use a `futureMetadata name="XLRICHVALUE"` table containing `<xlrd:rvb i="…"/>` entries.
  - Other variants may omit `futureMetadata`/`rvb` and may use `rc/@v` directly as the rich value index
    (not currently observed in the `image-in-cell*.xlsx` fixtures in this repo).
- The rich value data lives in either:
  - `xl/richData/richValue*.xml` (a list of `<rv>` entries), or
  - the **`rdRichValue*`** variant (`xl/richData/rdrichvalue.xml` + `xl/richData/rdrichvaluestructure.xml`)
    where positional `<v>` fields are interpreted via a structure table.
- Images are referenced indirectly via an integer **relationship-slot index** into
  `xl/richData/richValueRel.xml`, which in turn resolves to an OPC relationship in
  `xl/richData/_rels/richValueRel.xml.rels`, pointing at an `xl/media/*` part.

This document captures the part relationships and (most importantly) the **index mappings** needed to implement full support (reader → model → writer) and to avoid compatibility regressions when round-tripping unknown rich-data content.

> Note: The exact element/type names inside `<rv>` for image payloads vary by Excel version and are not
> fully specified in the public ECMA-376 base schema. This repo contains **real Excel fixtures** showing
> concrete shapes/namespaces for:
> - the unprefixed **`richValue*`** variant (see `fixtures/xlsx/rich-data/images-in-cell.xlsx`), and
> - the **`rdRichValue*`** variant (see `fixtures/xlsx/basic/image-in-cell.xlsx`),
> plus minimal synthetic fixtures for regression tests. See
> [Observed in fixtures](#observed-in-fixtures-in-repo) and [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md).

---

## Parts involved (minimum set for images-in-cell)

```
xl/
├── worksheets/
│   └── sheetN.xml                      # Cell has vm="…"
├── metadata.xml                        # Value metadata indexed by vm
└── richData/
    ├── richValue.xml / richValue1.xml  # Rich values (<rv>), indexed by metadata (schema varies)
    ├── richValueTypes.xml              # Optional (type id -> structure id)
    ├── richValueStructure.xml          # Optional (field layout)
    ├── rdrichvalue.xml                 # rdRichValue variant rich values (<rv> + positional <v> fields)
    ├── rdrichvaluestructure.xml        # rdRichValue variant structure table (ordered keys for <v> positions)
    ├── rdRichValueTypes.xml            # rdRichValue variant type/key flags (often present)
    ├── richValueRel.xml                # Relationship-slot table (root/name/namespace varies)
    └── _rels/
        └── richValueRel.xml.rels       # OPC relationships: rId -> ../media/image*.png
xl/media/
└── image*.{png,jpg,...}                # Binary image payload
```

The package also needs the usual bookkeeping:

- In this repo’s fixture corpus, workbooks that include `xl/metadata.xml` and/or `xl/richData/*` also include
  explicit `[Content_Types].xml` overrides for those parts. Preserve whatever the source workbook uses and do
  not hardcode a single required set.
- The workbook part (`xl/workbook.xml` + `xl/_rels/workbook.xml.rels`) typically contains a relationship to `xl/metadata.xml` and often relates directly to the rich value parts.

Observed (in this repo) relationship type URIs:

- Workbook → metadata:
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata`
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata`
- Workbook → rich value store (Microsoft-specific, versioned):
  - `http://schemas.microsoft.com/office/2017/06/relationships/richValue`
  - `http://schemas.microsoft.com/office/2017/06/relationships/richValueRel`
  - (variant, when the richData parts are related from `xl/metadata.xml` via `xl/_rels/metadata.xml.rels`)
    - `http://schemas.microsoft.com/office/2017/relationships/richValue`
    - `http://schemas.microsoft.com/office/2017/relationships/richValueRel`
    - `http://schemas.microsoft.com/office/2017/relationships/richValueTypes`
    - `http://schemas.microsoft.com/office/2017/relationships/richValueStructure`
  - `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue` (rdRichValue naming)
  - `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure` (rdRichValue naming)
  - `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes` (rdRichValue naming)
  - `http://schemas.microsoft.com/office/2022/10/relationships/richValueRel`

For details and fixture references, see:

- [`docs/20-images-in-cells-richdata.md`](./20-images-in-cells-richdata.md) (general richValue* / rdRichValue* notes)
- [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md) (concrete “Place in Cell” mapping chain)

This doc focuses on the **parts listed above** because they form the minimal chain to go from a cell → image bytes.

---

## The index chain (cell → metadata → rich value → relationship slot → media)

At a high level:

```
sheetN.xml: <c vm="VM_INDEX">…</c>
  └─> xl/metadata.xml: valueMetadata[VM_INDEX]   (vm can be 0-based or 1-based; treat as opaque)
         └─> RV_INDEX (resolved best-effort; in this repo’s fixtures: rc/@v -> <futureMetadata name="XLRICHVALUE"> -> <xlrd:rvb i="..."/>)
                └─> xl/richData/richValue*.xml (or xl/richData/rdrichvalue.xml): <rv> entry at RV_INDEX
                       └─> (image payload references REL_SLOT_INDEX)
                             └─> xl/richData/richValueRel.xml: <rel> at REL_SLOT_INDEX => r:id="rIdX"
                                   └─> xl/richData/_rels/richValueRel.xml.rels: Relationship Id="rIdX"
                                        └─> Target="../media/imageY.png" => xl/media/imageY.png bytes
```

### Indexing notes (practical assumptions)

- **`vm` is ambiguous (0-based or 1-based).** Both appear in this repo:
  - 1-based: `fixtures/xlsx/metadata/rich-values-vm.xlsx` (synthetic Formula fixture; see `crates/formula-xlsx/tests/metadata_rich_value_roundtrip.rs`)
  - 1-based: `fixtures/xlsx/basic/image-in-cell.xlsx` (Excel-generated “Place in Cell” fixture; `vm="1"`/`vm="2"`)
  - 0-based: `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`
- Rich value indices are **0-based**.
- `xl/metadata.xml` schemas vary:
  - In all images-in-cell fixtures in this repo, we observed the `futureMetadata`/`rvb` variant where
    `rc/@v` indexes into the `futureMetadata` table and `xlrd:rvb/@i` is the rich value index.
  - Other schemas may exist in the wild; treat metadata as opaque and preserve unknown tables/attributes.
- `richValue*.xml` can be **split across multiple parts** (`richValue.xml`, `richValue1.xml`, …). The `RV_INDEX` should be interpreted as a **global index across the concatenated `<rv>` streams** in part order.
  - Formula interprets part order using **numeric-suffix ordering** (`richValue.xml` then `richValue1.xml`, `richValue2.xml`, …), which is enforced by
    `crates/formula-xlsx/tests/rich_value_part_numeric_suffix_order.rs`.
  - Open question (Excel): the exact ordering rules Excel uses when multiple parts exist (lexicographic vs numeric suffix). Treat the writer as
    “preserve existing parts” and avoid renumbering unless the mapping is rebuilt holistically.
- Image references inside `<rv>` appear to use an **integer relationship-slot index** (not an `rId` string directly). That slot index points into the ordered `<rel>` list in `richValueRel.xml`.

---

## Observed in fixtures (in-repo)

These are fixture XLSX files in this repository used for parser/round-trip testing (some Excel-generated,
some synthetic). Prefer the Excel-generated fixtures as the primary ground truth for namespaces/root
element names.

### Fixture: `fixtures/xlsx/rich-data/images-in-cell.xlsx` (Excel `richValue*` + `cellimages.xml`)

See also: [`fixtures/xlsx/rich-data/images-in-cell.md`](../fixtures/xlsx/rich-data/images-in-cell.md) (walkthrough).

This fixture contains:

* `xl/cellimages.xml` + `xl/_rels/cellimages.xml.rels` (cell image store)
* `xl/metadata.xml` + `xl/_rels/metadata.xml.rels` (value/cell metadata; `futureMetadata`/`xlrd:rvb` mapping)
* full `xl/richData/richValue*.xml` table set:
  * `richValue.xml`
  * `richValueRel.xml`
  * `richValueTypes.xml`
  * `richValueStructure.xml`

Worksheet-level note (important for parsing):

* In `xl/worksheets/sheet1.xml`, the in-cell image at `A1` is encoded as a plain numeric cell:
  * `<c r="A1" vm="1" cm="1"><v>0</v></c>`
* In contrast, other real Excel “Place in Cell” fixtures in this repo use an error cell encoding
  (`t="e"` with cached `#VALUE!`). Do not treat the cached `<v>` (or `t="e"`) as authoritative for image
  binding; use `vm`/`cm` + `xl/metadata.xml` + `xl/richData/*`.

Notable shape differences vs the minimal/synthetic `image-in-cell-richdata.xlsx` fixture:

* `xl/richData/richValue.xml` uses a `<values>` wrapper and a `type="…"` attribute:
  * `<rv type="0"><v kind="rel">0</v></rv>`
* `xl/richData/richValueRel.xml` uses root `<rvRel>` (namespace `…/2017/richdata`) and wraps entries in `<rels>`.

- `fixtures/xlsx/basic/image-in-cell.xlsx` is **real Excel-generated** (see `docProps/app.xml`).
- `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` is **synthetic** (tagged `Application=Formula Fixtures`).

Treat Excel-generated fixtures as the primary ground truth for Excel’s current on-disk schema. Synthetic
fixtures are still useful for exercising edge cases and ensuring we preserve unknown parts/attributes.

### Fixture: `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` (synthetic; `richValue.xml` + `richValueRel.xml` 2017 variant)

See also: [`fixtures/xlsx/basic/image-in-cell-richdata.md`](../fixtures/xlsx/basic/image-in-cell-richdata.md) (walkthrough of this fixture).

Note: this fixture is intentionally minimal and is tagged in `docProps/app.xml` as `Application=Formula Fixtures`
(not Excel). It is still useful for testing round-trip preservation and basic parsing.

**Parts present** (complete inventory from `unzip -l`):

```
[Content_Types].xml
_rels/.rels
docProps/core.xml
docProps/app.xml
xl/workbook.xml
xl/_rels/workbook.xml.rels
xl/worksheets/sheet1.xml
xl/styles.xml
xl/metadata.xml
xl/richData/richValue.xml
xl/richData/richValueRel.xml
xl/richData/_rels/richValueRel.xml.rels
xl/media/image1.png
```

**`xl/metadata.xml` (`futureMetadata name="XLRICHVALUE"` + `xlrd:rvb i="..."`)**

```xml
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000" maxSupportedVersion="120000"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{3E2803F5-59A4-4A43-8C86-93BA0C219F4F}">
          <xlrd:rvb i="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk>
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
```

Notes:

* In this fixture, worksheet `c/@vm` is **0-based** (`vm="0"` selects the first `valueMetadata` record).
* `rc/@v` is a **0-based** index into the `<futureMetadata name="XLRICHVALUE">` `<bk>` list.
* `xlrd:rvb/@i` is the **0-based rich value index** into `xl/richData/richValue.xml`.

**`xl/richData/richValue.xml` + `xl/richData/richValueRel.xml` namespaces**

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv s="0" t="image"><v kind="rel">0</v></rv>
</rvData>
```

```xml
<richValueRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>
```

* Namespace is **`…/2017/richdata`** for `richValue.xml` and **`…/2017/richdata2`** for `richValueRel.xml`.
* The image payload is a single integer `<v>` (in this fixture: `<v kind="rel">…</v>`) and it is the
  0-based index into the `<rel>` list.
  * Shape: `<rv t="image"><v kind="rel">REL_SLOT</v></rv>`
  * In this fixture: `REL_SLOT = 0`

### Fixture: `fixtures/xlsx/basic/image-in-cell.xlsx` (`rdrichvalue.xml` + `richValueRel.xml` 2022 variant)

See also: [`fixtures/xlsx/basic/image-in-cell.md`](../fixtures/xlsx/basic/image-in-cell.md) (walkthrough of this fixture).

**Parts present** (complete inventory from `unzip -l`):

```
[Content_Types].xml
_rels/.rels
docProps/core.xml
docProps/app.xml
xl/workbook.xml
xl/_rels/workbook.xml.rels
xl/theme/theme1.xml
xl/styles.xml
xl/sharedStrings.xml
xl/worksheets/sheet1.xml
xl/worksheets/_rels/sheet1.xml.rels
xl/metadata.xml
xl/richData/richValueRel.xml
xl/richData/rdrichvalue.xml
xl/richData/rdrichvaluestructure.xml
xl/richData/rdRichValueTypes.xml
xl/richData/_rels/richValueRel.xml.rels
xl/media/image1.png
xl/media/image2.png
xl/printerSettings/printerSettings1.bin
```

**`xl/metadata.xml` (`futureMetadata name="XLRICHVALUE"` + `xlrd:rvb i="..."`)**

```xml
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="2">
    <bk>
      <extLst>
        <ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}">
          <xlrd:rvb i="0"/>
        </ext>
      </extLst>
    </bk>
    <bk>
      <extLst>
        <ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}">
          <xlrd:rvb i="1"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="2">
    <bk><rc t="1" v="0"/></bk>
    <bk><rc t="1" v="1"/></bk>
  </valueMetadata>
</metadata>
```

Notes:

* `rc/@v` is a 0-based index into the `<futureMetadata name="XLRICHVALUE">` `<bk>` list.
* `xlrd:rvb/@i` provides the 0-based rich value index into `xl/richData/rdrichvalue.xml`.

**`xl/richData/richValueRel.xml` root + namespace (2022 variant)**

```xml
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
  <rel r:id="rId2"/>
</richValueRels>
```

**`xl/richData/rdrichvalue.xml` + `rdrichvaluestructure.xml` (positional `<v>` fields)**

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="2">
  <rv s="0"><v>0</v><v>5</v></rv>
  <rv s="0"><v>1</v><v>5</v></rv>
</rvData>
```

```xml
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <s t="_localImage">
    <k n="_rvRel:LocalImageIdentifier" t="i"/>
    <k n="CalcOrigin" t="i"/>
  </s>
</rvStructures>
```

Notes:

* The `<k>` list defines the meaning of each `<v>` position.
* The relationship-slot index is stored in the field named **`_rvRel:LocalImageIdentifier`** (not
  necessarily “the first `<v>`”).
  * In this specific fixture, the `_localImage` key order is:
    1) `_rvRel:LocalImageIdentifier`
    2) `CalcOrigin`
    so `<rv><v>0</v><v>5</v></rv>` means: relationship slot `0`, `CalcOrigin = 5`.
* `xl/richData/_rels/richValueRel.xml.rels` is an unordered map; do not assume its `<Relationship>` order
  matches the `<rel>` order in `richValueRel.xml`. Resolve the slot index using the `<rel>` list order,
  then look up that `r:id` in the `.rels` file to find the `Target`.

**`xl/richData/rdRichValueTypes.xml` root + namespace**

This fixture also includes `xl/richData/rdRichValueTypes.xml`, with root/local-name `rvTypesInfo` in the
`http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2` namespace.

---

## Where this is implemented in Formula (code pointers)

* Rich-value chain extractor:
  * [`crates/formula-xlsx/src/rich_data/mod.rs`](../crates/formula-xlsx/src/rich_data/mod.rs)
* `metadata.xml` parsing used for `XlsxDocument::rich_value_index()`:
  * [`crates/formula-xlsx/src/read/mod.rs`](../crates/formula-xlsx/src/read/mod.rs) (`MetadataPart`)

---

## Synthetic end-to-end example

The following example is *synthetic* but demonstrates the mapping.

### 1) Worksheet cell (`xl/worksheets/sheet1.xml`)

```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <!-- vm selects a valueMetadata record in xl/metadata.xml (often 1-based, but 0-based is also observed) -->
      <c r="B2" vm="1">
        <!-- The cell's plain <v> is not the image payload.
             The image binding is driven by vm + metadata.xml. -->
        <v>0</v>
      </c>
    </row>
  </sheetData>
</worksheet>
```

### 2) Value metadata (`xl/metadata.xml`)

This is where Excel binds a cell’s `vm` index to a rich value index via the richData extension element
`<xlrd:rvb>` (in the `futureMetadata` variant). Other files omit `futureMetadata`/`rvb` entirely.

```xml
<metadata
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">

  <!-- The list of metadata “types”. Records refer to this by `rc/@t`, which may be 1-based or 0-based
       depending on the producer/Excel build (1-based is observed in the Excel fixtures in this repo). -->
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>

  <!-- Future metadata table: rc/@v is a 0-based index into this bk list. -->
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <!-- ext/@uri GUID varies by producer/version; preserve it byte-for-byte when round-tripping. -->
        <!-- Observed in `fixtures/xlsx/basic/image-in-cell.xlsx`: {3e2802c4-a4d2-4d8b-9148-e3be6c30e623} -->
        <ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}">
          <!-- i="5" => rich value #5 (0-based) in the rich value table
               (e.g. richValue*.xml or rdrichvalue.xml depending on the naming scheme) -->
          <xlrd:rvb i="5"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>

  <!-- vm="1" points at the first (1st) value-metadata record in 1-based files (0-based is also observed). -->
  <valueMetadata count="1">
    <bk>
      <!-- Record 1: type t="1" (metadataTypes[0]) and v="0" (futureMetadata bk[0]) -->
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
```

### 3) Rich values (`xl/richData/richValue*.xml`)

`richValue*.xml` contains an ordered list of `<rv>` entries (often under an `<rvData>` root). The rich value index selects an `<rv>` by position (unless an explicit global index attribute is present).

For images-in-cell, the `<rv>` payload includes (at minimum) some reference to an **image relationship slot index**. The exact field name is Excel-version dependent; the important part is that it’s an **integer slot index** into `richValueRel.xml`.

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <!-- ... rv[0] .. rv[4] ... -->

  <!-- rv[5] (selected by xlrd:rvb i="5") -->
  <rv t="image">
    <!-- Synthetic payload: relationship-slot index = 0 -->
    <v kind="rel">0</v>
  </rv>
</rvData>
```

### 4) Relationship slot table (`xl/richData/richValueRel.xml`)

`richValueRel.xml` is an **ordered table** mapping an integer slot index to an `r:id`.

```xml
<richValueRel
  xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

  <!-- Slot 0 (0-based) -->
  <rel r:id="rId1"/>

  <!-- Slot 1 -->
  <rel r:id="rId2"/>
</richValueRel>
```

### 5) OPC relationships (`xl/richData/_rels/richValueRel.xml.rels`)

This is standard OPC. The `r:id` from `richValueRel.xml` is resolved here to a concrete target part.

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship
    Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image1.png"/>

  <Relationship
    Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image2.png"/>
</Relationships>
```

### 6) Image payload (`xl/media/image1.png`)

The media part is raw bytes; the relationship target above points to it.

```
xl/media/image1.png   # PNG binary payload for cell B2
```

---

## Writer / model implications (what must be preserved)

Even before full rich-data editing is implemented, round-trip compatibility needs:

- **Preserve `vm="…"` attributes** on `<c>` elements (even if unknown).
- Preserve `xl/metadata.xml` **byte-for-byte** if we don’t understand/edit it (same strategy as charts).
- Preserve all `xl/richData/*` parts (including additional, not-yet-modeled parts Excel may emit).
- Preserve **relationship ordering** inside `richValueRel.xml` because rich values appear to reference relationships by **slot index**, not by `rId` string.

---

## Known gaps / uncertainties (remaining)

1. **More `<rv>` payload variants**
   - This repo confirms two concrete encodings:
      - `richValue.xml` variant: `<rv t="image"><v kind="rel">0</v></rv>` (0 = relationship-slot index) with
        `richValueRel.xml` root `<richValueRel xmlns="…/2017/richdata2">`.
      - `rdRichValue*` variant: `rdrichvalue.xml` `<rv><v>…</v><v>…</v></rv>` interpreted via
        `rdrichvaluestructure.xml`, with the relationship slot index in the key
        `_rvRel:LocalImageIdentifier`, and `richValueRel.xml` root
        `<richValueRels xmlns="…/2022/richvaluerel">`.
   - Other Excel builds may emit additional fields/structures; preserve unknown subtrees.
2. **Multi-part `richValue*.xml` behavior**
   - When does Excel split into `richValue1.xml`, `richValue2.xml`, etc.?
   - Are indices global across all parts? (Assumed yes.)
3. **Namespace URIs / extension GUIDs**
   - Real fixtures use the `xlrd` prefix with namespace `http://schemas.microsoft.com/office/spreadsheetml/2017/richdata`
     and an `ext/@uri` GUID (treat GUIDs/extension URIs as opaque; preserve them).
4. **Workbook-level relationships**
   - We now have concrete relationship Type URIs for both variants (see the fixture inventories above),
     but additional Excel builds may introduce more versioned Microsoft relationship types.

If you add fixtures for this feature, document them under `fixtures/xlsx/**` and update this doc with the observed exact XML. Also update:

- [`docs/20-images-in-cells-richdata.md`](./20-images-in-cells-richdata.md) (the detailed richValue* spec note)

---

## Existing regression coverage (do not duplicate)

While planning richer `richData` support, multiple preservation/regression tests were proposed. The
`crates/formula-xlsx/tests/` suite has since grown broader coverage. Before adding new tests for
“preserve X on round-trip”, check the existing suite first.

### `vm` / `cm` preservation

- Capture on read (unit test):
  - `crates/formula-xlsx/src/read/mod.rs` (`reads_cell_cm_and_vm_attributes_into_cell_meta`)
- Preservation during cell patching:
  - `crates/formula-xlsx/tests/cell_metadata_preservation.rs`
  - `crates/formula-xlsx/tests/streaming_cell_metadata_preservation.rs`

### `xl/metadata.xml` preservation

- Document round-trip / part preservation:
  - `crates/formula-xlsx/tests/metadata_and_richdata_preservation.rs`
  - `crates/formula-xlsx/tests/preserve_rich_data_parts.rs`
- Preservation through calcChain removal / recalc policy changes:
  - `crates/formula-xlsx/tests/recalc_policy_preserves_metadata.rs`
- Preservation when the writer needs to synthesize `sharedStrings.xml`:
  - `crates/formula-xlsx/tests/richdata_preserved_on_save.rs`

### `xl/richData/*` parts preservation

- Package patching:
  - `crates/formula-xlsx/tests/rich_data_streaming_patch_preservation.rs`
- Document / streaming round-trip preservation:
  - `crates/formula-xlsx/tests/metadata_and_richdata_preservation.rs`
  - `crates/formula-xlsx/tests/streaming_preserve_rich_data_parts.rs`

### Recalc-policy / calcChain removal (preserve metadata)

- Ensures `calcChain.xml` can be dropped without also dropping `metadata.xml` (and their rels / content
  types):
  - `crates/formula-xlsx/tests/recalc_policy_preserves_metadata.rs`

### SharedStrings synthesis preservation

- Ensures rich-data parts are not rewritten/dropped when `sharedStrings.xml` is synthesized:
  - `crates/formula-xlsx/tests/richdata_preserved_on_save.rs`
- Ensures writer respects the sharedStrings target from `workbook.xml.rels` (and does not synthesize a
  second `xl/sharedStrings.xml`):
  - `crates/formula-xlsx/tests/shared_strings_target_resolution.rs`

---

## Internal task tracker (rich-data preservation)

The following tasks are now redundant (already covered by repo tests), and should not be re-implemented
as new standalone regression tests:

| Task | Status | Covered by |
|------|--------|------------|
| 190 | Covered / redundant | `crates/formula-xlsx/tests/rich_data_streaming_patch_preservation.rs` |
| 219 | Covered / redundant | `crates/formula-xlsx/src/read/mod.rs` + `crates/formula-xlsx/tests/cell_metadata_preservation.rs` |
| 232 / 271 | Covered / redundant | `crates/formula-xlsx/tests/recalc_policy_preserves_metadata.rs` |
| 238 | Covered / redundant | `crates/formula-xlsx/tests/streaming_cell_metadata_preservation.rs` |

Remaining gaps (still worth doing) are tracked elsewhere and include:

- Task 188 / 247: writer vm/cm emission
- Task 194: `metadata.xml.rels` discovery
- Task 344: real linked data types fixture (non-synthetic)
- Task 372: module naming cleanup
