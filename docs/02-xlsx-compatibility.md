# XLSX Compatibility Layer

## Overview

Perfect XLSX compatibility is the foundation of user trust. Users must be confident that their complex financial models, scientific calculators, and business-critical workbooks will load, calculate, and save without any loss of fidelity.

## Related docs

- [20-xlsx-rich-data.md](./20-xlsx-rich-data.md) — Excel `richData` / rich values (including “image in cell”)
- [20-images-in-cells.md](./20-images-in-cells.md) — Excel “Images in Cell” (`IMAGE()` / “Place in Cell”) packaging + schema notes
- [20-images-in-cells-richdata.md](./20-images-in-cells-richdata.md) — RichData (`richValue*`) tables used by images-in-cells

---

## XLSX File Format Structure

XLSX is a ZIP archive following Open Packaging Conventions (ECMA-376):

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
│   ├── cellImages.xml           # Excel “image in cell” definitions (in-cell pictures; name/casing varies in the wild)
│   ├── calcChain.xml            # Calculation order hints
│   ├── metadata.xml             # Cell/value metadata (Excel "Rich Data")
│   ├── richData/                # Excel 365+ rich values (data types, in-cell images)
│   │   ├── rdrichvalue.xml
│   │   ├── rdrichvaluestructure.xml
│   │   ├── rdrichvaluetypes.xml
│   │   ├── richValueRel.xml     # Indirection to rich-value relationships (e.g. images)
│   │   └── _rels/
│   │       └── richValueRel.xml.rels  # richValueRel -> xl/media/* (image binaries)
│   ├── theme/
│   │   └── theme1.xml           # Color/font theme
│   ├── worksheets/
│   │   ├── sheet1.xml           # Cell data, formulas
│   │   └── sheet2.xml
│   ├── drawings/
│   │   └── drawing1.xml         # Charts, shapes, images
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
│   ├── queryTables/
│   │   └── queryTable1.xml      # External data queries
│   ├── connections.xml          # External data connections
│   ├── externalLinks/
│   │   └── externalLink1.xml    # Links to other workbooks
│   ├── customXml/               # Power Query definitions (base64)
│   └── vbaProject.bin           # VBA macros (binary)
└── xl/_rels/
    ├── workbook.xml.rels        # Workbook relationships
    ├── cellImages.xml.rels      # Relationships for in-cell images (to xl/media/*; name/casing varies in the wild)
    └── metadata.xml.rels        # Relationships from metadata.xml to xl/richData/*
```

See also:
- [`docs/20-images-in-cells-richdata.md`](./20-images-in-cells-richdata.md) — Excel RichData (`richValue*`) parts used by “Images in Cell” / `IMAGE()`.
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

#### Images in Cells (`IMAGE()` / “Place in Cell”) (Rich Data + `metadata.xml`)

Newer Excel builds can store **images as cell values** (“Place in Cell” pictures, and the `IMAGE()` function) using workbook-level Rich Data parts (`xl/metadata.xml` + `xl/richData/*`) and worksheet cell metadata attributes like `c/@vm`.

This is distinct from legacy “floating” images stored under `xl/drawings/*`.

Further reading:
- [20-images-in-cells.md](./20-images-in-cells.md)
- [20-images-in-cells-richdata.md](./20-images-in-cells-richdata.md)
- [20-xlsx-rich-data.md](./20-xlsx-rich-data.md)

##### Worksheet cell encoding

Key observed behavior for “Place in Cell” images: the worksheet cell is encoded as an **error** (`t="e"`) with cached `#VALUE!`, and the real value is referenced through the `vm` (**value metadata**) attribute.

```xml
<c t="e" vm="N"><v>#VALUE!</v></c>
```

##### Mapping chain (high-level)

`sheetN.xml c@vm` → `xl/metadata.xml <valueMetadata>` → `xl/richData/rdrichvalue.xml` (or `xl/richData/richValue*.xml`) → `xl/richData/richValueRel.xml` → `xl/richData/_rels/richValueRel.xml.rels` → `xl/media/imageN.*`

##### `xl/richData/rdrichvalue.xml` (rich value instances)

`xl/richData/rdrichvalue.xml` is the workbook-level table of rich value *instances*. The `i="…"` from `xlrd:rvb` selects a record from this table.

The exact element vocabulary inside each rich value varies by Excel version and feature, but for in-cell images the rich value ultimately encodes a **relationship slot index** (an integer) that points into `xl/richData/richValueRel.xml`.

Representative (synthetic) shape:

```xml
<rv:richData xmlns:rv="http://schemas.microsoft.com/office/spreadsheetml/2017/06/main">
  <rv:richValues count="1">
    <!-- rich value index 0 -->
    <rv:rv>
      <!-- relationship slot index into richValueRel.xml -->
      <rv:rel>0</rv:rel>
    </rv:rv>
  </rv:richValues>
</rv:richData>
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
      <!-- t = (1-based) index into <metadataTypes>, v = (0-based) index into <futureMetadata> -->
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
```

##### 2) Workbook relationship types (workbook → richData parts)

Excel wires the rich-data parts via OPC relationships. The relationship type URIs we observed for in-cell images are:

- `http://schemas.microsoft.com/office/2022/10/relationships/richValueRel`
- `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue`
- `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure`
- `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes`
- (inside `xl/richData/_rels/richValueRel.xml.rels`) `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`

Representative `xl/_rels/workbook.xml.rels` snippet:

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <!-- ... -->
  <Relationship Id="rIdMeta"
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
                Target="richData/rdrichvaluetypes.xml"/>

  <Relationship Id="rIdRel"
                Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel"
                Target="richData/richValueRel.xml"/>
</Relationships>
```

##### 3) `richValueRel.xml` → `xl/media/*` via `.rels`

`xl/richData/richValueRel.xml` is an **ordered table** that maps a small integer slot index (referenced from rich values) to an `r:id`, which is then resolved via `xl/richData/_rels/richValueRel.xml.rels`.

`xl/richData/richValueRel.xml`:

```xml
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
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

##### 4) `[Content_Types].xml` overrides

Workbooks that include in-cell images typically include overrides like:

```xml
<Override PartName="/xl/metadata.xml"
          ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>

<Override PartName="/xl/richData/rdrichvalue.xml"
          ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>
<Override PartName="/xl/richData/rdrichvaluestructure.xml"
          ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>
<Override PartName="/xl/richData/rdrichvaluetypes.xml"
          ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>
<Override PartName="/xl/richData/richValueRel.xml"
          ContentType="application/vnd.ms-excel.richValueRel+xml"/>
```

##### Note: `xl/cellImages.xml` is optional (and may not appear in “Place in Cell” files)

Some online discussions reference `xl/cellImages.xml` (or lowercase `xl/cellimages.xml`) for in-cell
pictures.

In the “Place in Cell” fixtures we inspected, in-cell images were represented using
`xl/metadata.xml` + `xl/richData/*` + `xl/media/*` and **no `cellImages` part** was present.

However, other producers (and some synthetic fixtures/tests in this repo) do include a `cellImages`
part; for round-trip safety we should treat it as optional and preserve it if present.

### Linked data types / Rich values (Stocks, Geography, etc.)

Modern Excel supports **linked data types** (Stocks, Geography, Organization, Power BI, etc.) where a cell has a normal displayed value *plus* an attached structured payload (used for the “card” UI and for field extraction like `=A1.Price`).

At the worksheet level, this shows up as additional attributes on the `<c>` (cell) element:

- `vm="…"`, the **value metadata index**
- `cm="…"`, the **cell metadata index**

These are integer indices into metadata tables stored at the workbook level (Excel-produced files
commonly use **1-based** indices for `vm`; treat `vm`/`cm` as opaque and preserve them exactly).

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

The structured payloads referenced by this metadata live under `xl/richData/` (commonly
`rdrichvaluetypes.xml` + `rdrichvalue.xml`, plus related supporting tables such as `richValueRel.xml`.
Some Excel builds use the unprefixed naming (`richValue*.xml`) for rich value instances/types/structure.

These pieces are connected via OPC relationships:

- `xl/_rels/workbook.xml.rels` typically has a relationship from the workbook to `xl/metadata.xml`.
- `xl/_rels/metadata.xml.rels` typically has relationships from `xl/metadata.xml` to `xl/richData/*` parts.

Simplified relationship sketch (workbook → metadata uses a stable OOXML relationship type; richData linkage types may vary across Excel builds):

```xml
<!-- xl/_rels/workbook.xml.rels -->
<Relationship Id="rIdMeta"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"
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

- Preserve `vm` / `cm` attributes on `<c>` elements when editing cell values.
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

### In-cell images (cellImages.xml)

Some producer tooling (and possibly some Excel builds) can store “images in cell” (pictures that behave like cell content rather than floating drawing objects) in a dedicated workbook-level OPC part:

- Part: `xl/cellImages.xml` (casing varies; `xl/cellimages.xml` is also seen in the wild)
- Relationships: `xl/_rels/cellImages.xml.rels` (casing varies; `xl/_rels/cellimages.xml.rels` is also seen)

However, in the Excel 365 “Place in Cell” fixtures we inspected, in-cell images were represented via
`xl/metadata.xml` + `xl/richData/*` + `xl/media/*` and no `xl/cellImages.xml` / `xl/cellimages.xml` part
was present. If we encounter a `cellImages` part in the wild, we should preserve it for round-trip safety.

From a **packaging / round-trip** perspective, the important thing is the relationship chain that connects this part to the actual image blobs under `xl/media/*`.

**Schema note:** `xl/cellImages.xml` is a Microsoft extension part; the root namespace / element vocabulary
has been observed to vary across Excel versions (e.g. `…/2019/cellimages`, `…/2022/cellimages`). For
round-trip, treat the **part path** (including its original casing) as authoritative, not the root namespace.

#### How it’s usually connected

1. `xl/workbook.xml` (via `xl/_rels/workbook.xml.rels`) contains a relationship that targets `cellImages.xml` (or `cellimages.xml`):
   - The relationship **Type URI is a Microsoft extension** and has been observed to vary across Excel builds.
   - **Detection strategy**: treat any relationship whose `Target` resolves (case-insensitively) to `xl/cellimages.xml` as authoritative, rather than hardcoding a single `Type` URI.
2. `xl/_rels/cellImages.xml.rels` contains relationships of type `…/relationships/image` pointing at `xl/media/*` files.
   - The relationship `Id` values (e.g. `rId1`) are referenced from within `xl/cellImages.xml` via `r:embed`, so they must be preserved (or updated consistently if rewriting).
   - Targets are typically relative paths like `media/image1.png` (resolving to `/xl/media/image1.png`), but should be preserved as-is.

#### `[Content_Types].xml` requirements

If `xl/cellImages.xml` is present, the package typically includes an override:

- `<Override PartName="/xl/cellImages.xml" ContentType="…"/>` (or the lowercase variant)

Excel uses a **Microsoft-specific** content type string for this part (the exact string may vary between versions).

Observed in this repo (see `crates/formula-xlsx/tests/cell_images.rs` and
`crates/formula-xlsx/tests/cellimages_preservation.rs`; do not hardcode):
- `application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml`
- `application/vnd.ms-excel.cellimages+xml`

**Preservation/detection strategy:**
- Treat any `[Content_Types].xml` `<Override>` whose `PartName` case-insensitively matches `/xl/cellimages.xml` as authoritative.
- Preserve the `ContentType` value byte-for-byte on round-trip; **do not** hardcode a single MIME string in the writer.

#### Relationship type URIs

- `xl/_rels/cellImages.xml.rels` (or `xl/_rels/cellimages.xml.rels`) → `xl/media/*`:
  - **High confidence**: `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"`
- `xl/workbook.xml.rels` → `xl/cellImages.xml`:
  - **Microsoft extension** (variable). Prefer detection by `Target`/part name.

#### Minimal (non-normative) XML snippets

Workbook relationship entry (in `xl/_rels/workbook.xml.rels`):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId42"
                Type="http://schemas.microsoft.com/office/.../relationships/cellimages"
                Target="cellImages.xml"/>
</Relationships>
```

Cellimages-to-media relationship entry (in `xl/_rels/cellImages.xml.rels`):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png"/>
</Relationships>
```

Cellimages part referencing an image by relationship id (in `xl/cellImages.xml`):

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
  the project digest)
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

```typescript
interface XlsxDocument {
  // Fully parsed and modeled
  workbook: Workbook;
  sheets: Sheet[];
  styles: StyleSheet;
  
  // Preserved as raw XML for round-trip
  unknownParts: Map<PartPath, XmlDocument>;
  
  // Preserved byte-for-byte
  binaryParts: Map<PartPath, Uint8Array>;  // vbaProject.bin, etc.
}
```

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
- The diff tool is implemented in Rust: `crates/xlsx-diff`.

Run a diff locally:

```bash
cargo run -p xlsx-diff --bin xlsx_diff -- original.xlsx roundtripped.xlsx
```

Run the fixture harness (used by CI):

```bash
cargo test -p xlsx-diff --test roundtrip_fixtures
```

The harness performs a real load → save using `formula-xlsx::XlsxPackage` (OPC-level
package handling) and then diffs the original vs written output.

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

### Streaming Parsing

For large files, don't load entire XML into memory:

```typescript
// BAD: Load entire file
const xml = await parseXml(await readFile(path));
const cells = xml.querySelectorAll("c");

// GOOD: Stream parsing
const parser = new SaxParser();
parser.on("element:c", (cell) => {
  processCell(cell);
});
await parser.parseStream(fileStream);
```

### Lazy Loading

Don't parse everything upfront:

```typescript
class LazyWorksheet {
  private parsed = false;
  private xmlPath: string;
  private data?: SheetData;
  
  async getData(): Promise<SheetData> {
    if (!this.parsed) {
      this.data = await this.parse();
      this.parsed = true;
    }
    return this.data!;
  }
}
```

### Parallel Processing

Parse independent parts concurrently:

```typescript
const [workbook, styles, sharedStrings] = await Promise.all([
  parseWorkbook(archive),
  parseStyles(archive),
  parseSharedStrings(archive),
]);

// Then parse sheets (which depend on above)
const sheets = await Promise.all(
  workbook.sheets.map(s => parseSheet(archive, s, styles, sharedStrings))
);
```

---

## Future Considerations

1. **Excel for Web compatibility**: Some features differ in web version
2. **Google Sheets export**: Import/export from Google's format
3. **Numbers compatibility**: Apple's format for Mac users
4. **OpenDocument (ODS)**: LibreOffice compatibility
5. **New Excel features**: Monitor Excel updates for new XML schemas
6. **Images in Cell** (Place in Cell / `IMAGE()`): packaging + schema notes in [20-images-in-cells.md](./20-images-in-cells.md)
