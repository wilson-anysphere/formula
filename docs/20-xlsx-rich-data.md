# XLSX Rich Data (`richData`) and “Image in Cell” Storage

## Overview

Excel historically stores images via the **drawing layer** (`xl/drawings/*`, anchored/floating shapes). Newer Excel builds (Microsoft 365) also support **“Place in Cell” / “Image in Cell”**, where an image behaves like a *cell value*.

For the dedicated “images in cells” packaging overview (including the optional `xl/cellImages.xml` part), see:

- [`docs/20-images-in-cells.md`](./20-images-in-cells.md)
- [`docs/20-images-in-cells-richdata.md`](./20-images-in-cells-richdata.md)
- [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md) (confirmed “Place in Cell” rich-value wiring + exact URIs)
  - See also the Excel-produced fixture `fixtures/xlsx/basic/image-in-cell.xlsx` (notes in `fixtures/xlsx/basic/image-in-cell.md`).

In OOXML this is implemented using **Rich Values** (“richData”) plus **cell value metadata**:

- The worksheet cell points to a **value-metadata record** via the cell attribute `vm="…"`.
- That value-metadata record binds the cell to a **rich value index**. The exact mapping schema varies:
  - Some files use a `futureMetadata name="XLRICHVALUE"` table containing `<xlrd:rvb i="…"/>` entries.
  - Other files omit `futureMetadata`/`rvb` and appear to use `rc/@v` directly as the rich value index.
- The rich value data lives in either:
  - `xl/richData/richValue*.xml` (a list of `<rv>` entries), or
  - the **`rdRichValue*`** variant (`xl/richData/rdrichvalue.xml` + `xl/richData/rdrichvaluestructure.xml`)
    where positional `<v>` fields are interpreted via a structure table.
- Images are referenced indirectly via an integer **relationship-slot index** into
  `xl/richData/richValueRel.xml`, which in turn resolves to an OPC relationship in
  `xl/richData/_rels/richValueRel.xml.rels`, pointing at an `xl/media/*` part.

This document captures the part relationships and (most importantly) the **index mappings** needed to implement full support (reader → model → writer) and to avoid compatibility regressions when round-tripping unknown rich-data content.

> Note: The exact element/type names inside `<rv>` for image payloads vary by Excel version and are not
> fully specified in the public ECMA-376 base schema. This repo contains **two real fixtures** showing
> concrete shapes/namespaces for both the `richValue.xml` and `rdRichValue*` variants — see
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

- `[Content_Types].xml` may include overrides for `metadata.xml` and the `xl/richData/*` parts (some files rely on the default `application/xml`).
- The workbook part (`xl/workbook.xml` + `xl/_rels/workbook.xml.rels`) typically contains a relationship to `xl/metadata.xml` and often relates directly to the rich value parts.

This doc focuses on the **parts listed above** because they form the minimal chain to go from a cell → image bytes.

---

## The index chain (cell → metadata → rich value → relationship slot → media)

At a high level:

```
sheetN.xml: <c vm="VM_INDEX">…</c>
  └─> xl/metadata.xml: valueMetadata[VM_INDEX]   (vm can be 0-based or 1-based; treat as opaque)
         └─> RV_INDEX (resolved best-effort; either directly from rc/@v, or indirectly via an <xlrd:rvb i="..."/> table)
               └─> xl/richData/richValue*.xml (or xl/richData/rdrichvalue.xml): <rv> entry at RV_INDEX
                      └─> (image payload references REL_SLOT_INDEX)
                            └─> xl/richData/richValueRel.xml: <rel> at REL_SLOT_INDEX => r:id="rIdX"
                                  └─> xl/richData/_rels/richValueRel.xml.rels: Relationship Id="rIdX"
                                        └─> Target="../media/imageY.png" => xl/media/imageY.png bytes
```

### Indexing notes (practical assumptions)

- **`vm` is ambiguous (0-based or 1-based).** Both appear in this repo:
  - 1-based: `fixtures/xlsx/metadata/rich-values-vm.xlsx` (see `crates/formula-xlsx/tests/metadata_rich_value_roundtrip.rs`)
  - 1-based: `fixtures/xlsx/basic/image-in-cell.xlsx` (Excel-generated “Place in Cell” fixture; `vm="1"`/`vm="2"`)
  - 0-based: `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`
- Rich value indices are **0-based**.
- `xl/metadata.xml` schemas vary:
  - In the `futureMetadata`/`rvb` variant, `rc/@v` indexes into the `futureMetadata` table, and `xlrd:rvb/@i` is the rich value index.
  - In other variants, `rc/@v` appears to directly be the rich value index.
- `richValue*.xml` can be **split across multiple parts** (`richValue.xml`, `richValue1.xml`, …). The `RV_INDEX` should be interpreted as a **global index across the concatenated `<rv>` streams** in part order.
  - Open question: the exact ordering rules Excel uses when multiple parts exist (lexicographic vs numeric suffix). Use numeric suffix ordering (`richValue.xml` then `richValue1.xml`, `richValue2.xml`, …) and verify with fixtures.
- Image references inside `<rv>` appear to use an **integer relationship-slot index** (not an `rId` string directly). That slot index points into the ordered `<rel>` list in `richValueRel.xml`.

---

## Observed in fixtures (in-repo)

These are **real Excel-generated** XLSX files in this repository, and should be treated as the primary
ground truth for namespaces/root element names.

### Fixture: `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` (`richValue.xml` + `richValueRel.xml` 2017 variant)

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

**`xl/metadata.xml` shape (no `futureMetadata`)**

```xml
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="0" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>
  </metadataTypes>
  <valueMetadata count="1">
    <bk>
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
```

Notes:

* In this fixture, worksheet `c/@vm` is **0-based** (`vm="0"` selects the first `valueMetadata` record).
* With no `futureMetadata` table present, `rc/@v` **appears to be the rich value index** (0-based).

**`xl/richData/richValue.xml` + `xl/richData/richValueRel.xml` namespaces**

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv s="0" t="image"><v>REL_SLOT</v></rv>
</rvData>
```

```xml
<richValueRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>
```

* Namespace is **`…/2017/richdata`** for `richValue.xml` and **`…/2017/richdata2`** for `richValueRel.xml`.
* The image payload is a single integer `<v>`: `REL_SLOT` is the 0-based index into the `<rel>` list.

### Fixture: `fixtures/xlsx/basic/image-in-cell.xlsx` (`rdrichvalue.xml` + `richValueRel.xml` 2022 variant)

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
      <!-- vm="1" => the first valueMetadata record in xl/metadata.xml -->
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

  <!-- The list of metadata “types”. Records refer to this by 1-based index (t="…"). -->
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

  <!-- vm="1" points at the first (1st) value-metadata record. -->
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
    <v>0</v>
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
     - `richValue.xml` variant: `<rv t="image"><v>REL_SLOT</v></rv>` with
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
