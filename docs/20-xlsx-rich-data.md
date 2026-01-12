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
- The rich value data lives in `xl/richData/richValue*.xml` as an `<rv>` entry (or an “rdRichValue” variant; see below).
- Images are referenced indirectly via an index into `xl/richData/richValueRel.xml`, which in turn resolves to an OPC relationship in `xl/richData/_rels/richValueRel.xml.rels`, pointing at a `xl/media/*` part.

This document captures the part relationships and (most importantly) the **index mappings** needed to implement full support (reader → model → writer) and to avoid compatibility regressions when round-tripping unknown rich-data content.

> Note: The exact element/type names inside `<rv>` for image payloads vary by Excel version and are not
> fully specified in the public ECMA-376 base schema. This doc focuses on the *stable wiring* (OPC parts
> + index indirections) that we must preserve.
>
> One concrete schema is now confirmed for “Place in Cell” images produced by `rust_xlsxwriter`:
> worksheet cell `t="e"`/`#VALUE!` + `metadata.xml` + `xl/richData/rdrichvalue*.xml` + `richValueRel` →
> `xl/media/*`. See [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md).

---

## Parts involved (minimum set for images-in-cell)

```
xl/
├── worksheets/
│   └── sheetN.xml                      # Cell has vm="…"
├── metadata.xml                        # Value metadata indexed by vm
└── richData/
    ├── richValue.xml / richValue1.xml        # Rich value instances (<rv>), indexed by metadata (schema varies)
    ├── richValueTypes.xml                    # Optional: type table (type id -> structure id)
    ├── richValueStructure.xml                # Optional: structure table (field layout)
    ├── rdrichvalue.xml                       # Alternate naming (observed in rust_xlsxwriter output)
    ├── rdrichvaluestructure.xml              # Alternate naming
    ├── rdRichValueTypes.xml                  # Alternate naming (note casing)
    ├── richValueRel.xml                      # Ordered <rel r:id="…"> list (relationship slots)
    └── _rels/
        └── richValueRel.xml.rels       # OPC relationships: rId -> ../media/image*.png
xl/media/
└── image*.{png,jpg,...}                # Binary image payload
```

The package also needs the usual bookkeeping:

- `[Content_Types].xml` may include overrides for `metadata.xml` and the `xl/richData/*` parts (some files rely on the default `application/xml`).
- The workbook part (`xl/workbook.xml` + `xl/_rels/workbook.xml.rels`) typically contains a relationship to `xl/metadata.xml` and often relates directly to the rich value parts.

Observed (in this repo) relationship type URIs:

- Workbook → metadata:
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata`
  - `http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata`
- Workbook → rich value store (Microsoft-specific, versioned):
  - `http://schemas.microsoft.com/office/2017/06/relationships/richValue`
  - `http://schemas.microsoft.com/office/2017/06/relationships/richValueRel`
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
         └─> RV_INDEX (resolved best-effort; either directly from rc/@v, or indirectly via an <xlrd:rvb i="..."/> table)
               └─> xl/richData/richValue*.xml: <rv> entry at RV_INDEX
                      └─> (image payload references REL_SLOT_INDEX)
                            └─> xl/richData/richValueRel.xml: <rel> at REL_SLOT_INDEX => r:id="rIdX"
                                  └─> xl/richData/_rels/richValueRel.xml.rels: Relationship Id="rIdX"
                                        └─> Target="../media/imageY.png" => xl/media/imageY.png bytes
```

### Indexing notes (practical assumptions)

- **`vm` is ambiguous (0-based or 1-based).** Both appear in this repo:
  - 1-based: `fixtures/xlsx/metadata/rich-values-vm.xlsx` (see `crates/formula-xlsx/tests/metadata_rich_value_roundtrip.rs`)
  - 0-based: `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`
- Rich value indices are **0-based**.
- `xl/metadata.xml` schemas vary:
  - In the `futureMetadata`/`rvb` variant, `rc/@v` indexes into the `futureMetadata` table, and `xlrd:rvb/@i` is the rich value index.
  - In other variants, `rc/@v` appears to directly be the rich value index.
- `richValue*.xml` can be **split across multiple parts** (`richValue.xml`, `richValue1.xml`, …). The `RV_INDEX` should be interpreted as a **global index across the concatenated `<rv>` streams** in part order.
  - Open question: the exact ordering rules Excel uses when multiple parts exist (lexicographic vs numeric suffix). Use numeric suffix ordering (`richValue.xml` then `richValue1.xml`, `richValue2.xml`, …) and verify with fixtures.
- Image references inside `<rv>` appear to use an **integer relationship-slot index** (not an `rId` string directly). That slot index points into the ordered `<rel>` list in `richValueRel.xml`.

---

## Synthetic end-to-end example

The following example is *synthetic* but demonstrates the mapping.

### 1) Worksheet cell (`xl/worksheets/sheet1.xml`)

Cells that reference rich values always carry `vm="…"` to select a record in `xl/metadata.xml`, but the
*underlying cell value representation varies*:

* Some files store a normal `<v>` payload (often numeric) alongside `vm="…"`.
* “Place in Cell” embedded images (confirmed for rust_xlsxwriter) store the cell as an error:
  `t="e"` with cached `#VALUE!` and `vm="1"`. See
  [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md).

```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <!-- vm="1" => the first valueMetadata record in xl/metadata.xml -->
      <c r="B2" t="e" vm="1">
        <!-- The cached <v> (here `#VALUE!`) is not the image payload.
             The image binding is driven by vm + metadata.xml. -->
        <v>#VALUE!</v>
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
        <ext uri="{BDBB8CDC-FA1E-496E-A857-3C3F30B4D73F}">
          <!-- i="5" => rich value #5 (0-based) in richValue*.xml -->
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

## Known gaps / uncertainties (needs real Excel fixtures)

1. **Exact `<rv>` payload shape for images**
   - Confirmed for the `rdRichValue` / `_localImage` variant emitted by `rust_xlsxwriter`:
     `_rvRel:LocalImageIdentifier` + `CalcOrigin` (positional values) with `richValueRel` indirection.
     See: [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md).
   - Still an open question for other Excel/producers and for the `richValue.xml` (non-`rd*`) variants.
2. **Multi-part `richValue*.xml` behavior**
   - When does Excel split into `richValue1.xml`, `richValue2.xml`, etc.?
   - Are indices global across all parts? (Assumed yes.)
3. **Namespace URIs / extension GUIDs**
   - This doc uses the commonly observed `xlrd` prefix and a placeholder `ext/@uri` GUID. Confirm and treat as opaque (preserve even if unknown).
4. **Workbook-level relationships**
   - Confirm how `xl/metadata.xml` and `xl/richData/*` are linked from the workbook part in various Excel versions (relationship Type URIs are versioned Microsoft extensions).

If you add fixtures for this feature, document them under `fixtures/xlsx/**` and update this doc with the observed exact XML. Also update:

- [`docs/20-images-in-cells-richdata.md`](./20-images-in-cells-richdata.md) (the detailed richValue* spec note)
