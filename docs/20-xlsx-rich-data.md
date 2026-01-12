# XLSX Rich Data (`richData`) and “Image in Cell” Storage

## Overview

Excel historically stores images via the **drawing layer** (`xl/drawings/*`, anchored/floating shapes). Newer Excel builds (Microsoft 365) also support **“Place in Cell” / “Image in Cell”**, where an image behaves like a *cell value*.

In OOXML this is implemented using **Rich Values** (“richData”) plus **cell value metadata**:

- The worksheet cell points to a **value-metadata record** via the cell attribute `vm="…"`.
- That value-metadata record binds the cell to a **rich value** via the richData extension element `<xlrd:rvb i="…"/>`.
- The rich value data lives in `xl/richData/richValue*.xml` as an `<rv>` entry.
- Images are referenced indirectly via an index into `xl/richData/richValueRel.xml`, which in turn resolves to an OPC relationship in `xl/richData/_rels/richValueRel.xml.rels`, pointing at a `xl/media/*` part.

This document captures the part relationships and (most importantly) the **index mappings** needed to implement full support (reader → model → writer) and to avoid compatibility regressions when round-tripping unknown rich-data content.

> Note: The exact element/type names inside `<rv>` for image payloads vary by Excel version and are not fully specified in the public ECMA-376 base schema. Real Excel fixtures are required to validate the precise XML shape. This doc focuses on the *stable wiring* (OPC parts + index indirections) that we must preserve.

---

## Parts involved (minimum set for images-in-cell)

```
xl/
├── worksheets/
│   └── sheetN.xml                      # Cell has vm="…"
├── metadata.xml                        # Value metadata indexed by vm
└── richData/
    ├── richValue.xml / richValue1.xml  # Rich values (<rv>) indexed by xlrd:rvb/@i
    ├── richValueRel.xml                # Ordered <rel r:id="…"> list (relationship slots)
    └── _rels/
        └── richValueRel.xml.rels       # OPC relationships: rId -> ../media/image*.png
xl/media/
└── image*.{png,jpg,...}                # Binary image payload
```

The package also needs the usual bookkeeping:

- `[Content_Types].xml` must include overrides for `metadata.xml` and the `xl/richData/*` parts, plus image MIME types.
- The workbook part (`xl/workbook.xml` + `xl/_rels/workbook.xml.rels`) typically contains a relationship to `xl/metadata.xml` (and may also relate to richData parts depending on Excel version).

This doc focuses on the **parts listed above** because they form the minimal chain to go from a cell → image bytes.

---

## The index chain (cell → metadata → rich value → relationship slot → media)

At a high level:

```
sheetN.xml: <c vm="VM_INDEX">…</c>
  └─> xl/metadata.xml: valueMetadata[VM_INDEX]
        └─> <xlrd:rvb i="RV_INDEX"/>
              └─> xl/richData/richValue*.xml: <rv> entry at RV_INDEX
                    └─> (image payload references REL_SLOT_INDEX)
                          └─> xl/richData/richValueRel.xml: <rel> at REL_SLOT_INDEX => r:id="rIdX"
                                └─> xl/richData/_rels/richValueRel.xml.rels: Relationship Id="rIdX"
                                      └─> Target="../media/imageY.png" => xl/media/imageY.png bytes
```

### Indexing notes (practical assumptions)

- **`vm` and `i` are treated as 0-based indices** (consistent with most SpreadsheetML index fields like shared strings and style IDs). Validate against real Excel fixtures before hard-coding.
- `richValue*.xml` can be **split across multiple parts** (`richValue.xml`, `richValue1.xml`, …). The `RV_INDEX` should be interpreted as a **global index across the concatenated `<rv>` streams** in part order.
  - Open question: the exact ordering rules Excel uses when multiple parts exist (lexicographic vs numeric suffix). Use numeric suffix ordering (`richValue.xml` then `richValue1.xml`, `richValue2.xml`, …) and verify with fixtures.
- Image references inside `<rv>` appear to use an **integer relationship-slot index** (not an `rId` string directly). That slot index points into the ordered `<rel>` list in `richValueRel.xml`.

---

## Synthetic end-to-end example

The following example is *synthetic* but demonstrates the mapping.

### 1) Worksheet cell (`xl/worksheets/sheet1.xml`)

```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <!-- vm="0" => valueMetadata record #0 in xl/metadata.xml -->
      <c r="B2" vm="0">
        <!-- The cell's plain <v> is not the image payload.
             The image binding is driven by vm + metadata.xml. -->
        <v>0</v>
      </c>
    </row>
  </sheetData>
</worksheet>
```

### 2) Value metadata (`xl/metadata.xml`)

This is where Excel binds a cell’s `vm` index to a rich value index via the richData extension element `<xlrd:rvb>`.

```xml
<metadata
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">

  <!-- The list of metadata “types”. Records refer to this by index (t="…"). -->
  <metadataTypes count="1">
    <metadataType name="XLDAPR" minSupportedVersion="120000"/>
  </metadataTypes>

  <!-- vm="0" points at the first (0th) value-metadata record. -->
  <valueMetadata count="1">
    <bk>
      <!-- Record 0: type t="0" (metadataTypes[0]) -->
      <rc t="0">
        <extLst>
          <ext uri="{BDBB8CDC-FA1E-496E-A857-3C3F30B4D73F}">
            <!-- i="5" => rich value #5 (0-based) in richValue*.xml -->
            <xlrd:rvb i="5"/>
          </ext>
        </extLst>
      </rc>
    </bk>
  </valueMetadata>
</metadata>
```

### 3) Rich values (`xl/richData/richValue*.xml`)

`richValue*.xml` contains an ordered list of `<rv>` entries. The `i` from `<xlrd:rvb i="…"/>` selects an `<rv>` by position.

For images-in-cell, the `<rv>` payload includes (at minimum) some reference to an **image relationship slot index**. The exact field name is Excel-version dependent; the important part is that it’s an **integer slot index** into `richValueRel.xml`.

```xml
<richValues xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <!-- ... rv[0] .. rv[4] ... -->

  <!-- rv[5] (selected by xlrd:rvb i="5") -->
  <rv>
    <!-- Example/synthetic payload structure.
         The key idea: REL_SLOT_INDEX = 0 -->
    <imageRelSlot>0</imageRelSlot>
  </rv>
</richValues>
```

### 4) Relationship slot table (`xl/richData/richValueRel.xml`)

`richValueRel.xml` is an **ordered table** mapping an integer slot index to an `r:id`.

```xml
<richValueRels
  xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

  <!-- Slot 0 (0-based) -->
  <rel r:id="rId1"/>

  <!-- Slot 1 -->
  <rel r:id="rId2"/>
</richValueRels>
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
   - The field names / element structure used to point at the image relationship slot need to be confirmed with real `Place in Cell` files.
2. **Multi-part `richValue*.xml` behavior**
   - When does Excel split into `richValue1.xml`, `richValue2.xml`, etc.?
   - Are indices global across all parts? (Assumed yes.)
3. **Namespace URIs / extension GUIDs**
   - This doc uses the commonly observed `xlrd` prefix and a placeholder `ext/@uri` GUID. Confirm and treat as opaque (preserve even if unknown).
4. **Workbook-level relationships**
   - Confirm how `xl/metadata.xml` and `xl/richData/*` are linked from the workbook part in various Excel versions.

If you add fixtures for this feature, document them under `fixtures/xlsx/**` and update this doc with the observed exact XML.

