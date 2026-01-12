# `image-in-cell.xlsx` (Excel images-in-cells: Place in Cell + `IMAGE()`)

This fixture is an `.xlsx` intended to represent how modern Excel stores **images in cells**
using the Rich Data system (`xl/metadata.xml` + `xl/richData/*`) and the optional dedicated
cell image store part (`xl/cellimages.xml`).

It includes both:

- `Sheet1!A1`: an in-cell image inserted via **Insert → Pictures → Place in Cell**
- `Sheet1!B1`: an `_xlfn.IMAGE(...)` formula cell that also has a `vm="..."` rich-value binding

See also:

- [`docs/20-images-in-cells.md`](../../../docs/20-images-in-cells.md)
- Related real Excel fixture (no `IMAGE()` cell): `fixtures/xlsx/rich-data/images-in-cell.xlsx`
  (notes in `fixtures/xlsx/rich-data/images-in-cell.md`)

## Provenance

`docProps/app.xml` contains:

```xml
<Application>Microsoft Excel</Application>
```

## ZIP part inventory (complete)

Output of:

```bash
unzip -Z1 fixtures/xlsx/images-in-cells/image-in-cell.xlsx | sort
```

```text
[Content_Types].xml
_rels/.rels
docProps/app.xml
docProps/core.xml
xl/_rels/cellimages.xml.rels
xl/_rels/workbook.xml.rels
xl/cellimages.xml
xl/media/image1.png
xl/metadata.xml
xl/richData/_rels/richValueRel.xml.rels
xl/richData/richValue.xml
xl/richData/richValueRel.xml
xl/richData/richValueStructure.xml
xl/richData/richValueTypes.xml
xl/styles.xml
xl/workbook.xml
xl/worksheets/sheet1.xml
```

## Content types (`[Content_Types].xml`)

Key `<Override>` entries in this fixture:

```xml
<Override PartName="/xl/cellimages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>

<Override PartName="/xl/richData/richValue.xml" ContentType="application/vnd.ms-excel.richvalue+xml"/>
<Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
<Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/vnd.ms-excel.richvaluetypes+xml"/>
<Override PartName="/xl/richData/richValueStructure.xml" ContentType="application/vnd.ms-excel.richvaluestructure+xml"/>
```

## Workbook relationships (`xl/_rels/workbook.xml.rels`)

This fixture links both the `metadata.xml` part and the “richValue*” parts directly from the workbook:

```xml
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"
              Target="metadata.xml"/>
<Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/06/relationships/cellImages"
              Target="cellimages.xml"/>

<Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue"
              Target="richData/richValue.xml"/>
<Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel"
              Target="richData/richValueRel.xml"/>
<Relationship Id="rId7" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueTypes"
              Target="richData/richValueTypes.xml"/>
<Relationship Id="rId8" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueStructure"
              Target="richData/richValueStructure.xml"/>
```

## Worksheet cells (`vm="..."` on `<c>`)

From `xl/worksheets/sheet1.xml`:

```xml
<!-- A1: Place-in-Cell image value (rich value via vm) -->
<c r="A1" vm="1"><v>0</v></c>

<!-- B1: IMAGE() formula (rich value via vm) -->
<c r="B1" vm="2">
  <f>_xlfn.IMAGE("https://example.com/image.png")</f>
  <v>0</v>
</c>
```

Notable points:

- Both the “place in cell” value and the `IMAGE()` formula cell use the **numeric cached value** encoding (`<v>0</v>`)
  with `vm="..."` (not `t="e"` / cached `#VALUE!`).
- `vm` values are **1-based** in this file (`vm="1"` for the first record).

## `xl/metadata.xml` (`valueMetadata` → rich value index)

This fixture’s `metadata.xml` includes 2 value-metadata blocks, mapping to 2 rich values:

```xml
<metadataTypes count="2">
  <metadataType name="SOMEOTHERTYPE"/>
  <metadataType name="XLRICHVALUE"/>
</metadataTypes>

<futureMetadata name="XLRICHVALUE" count="2">
  <!-- futureMetadata #0 -->
  <bk>...<xlrd:rvb i="0"/>...</bk>
  <!-- futureMetadata #1 -->
  <bk>...<xlrd:rvb i="1"/>...</bk>
</futureMetadata>

<valueMetadata count="2">
  <!-- vm="1" selects bk #0 -->
  <bk><rc t="2" v="0"/></bk>
  <!-- vm="2" selects bk #1 -->
  <bk><rc t="2" v="1"/></bk>
</valueMetadata>
```

The `xlrd:rvb i="..."` values (`0` and `1`) select records from `xl/richData/richValue.xml`.

## `xl/richData/*` (“richValue*” naming scheme)

### `xl/richData/richValue.xml`

The rich value instance table (two entries in this fixture):

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">0</v>
      <v kind="string">image1</v>
    </rv>
    <rv type="0">
      <v kind="rel">0</v>
      <v kind="string">image1-again</v>
    </rv>
  </values>
</rvData>
```

Both rich values reference relationship slot `0` in `richValueRel.xml`.

### `xl/richData/richValueRel.xml` + `.rels`

Relationship slot indirection table:

```xml
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId1"/>
  </rels>
</rvRel>
```

And `xl/richData/_rels/richValueRel.xml.rels` resolves `rId1` to the media part:

```xml
<Relationship Id="rId1"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="../media/image1.png"/>
```

## `xl/cellimages.xml` (cell image store)

This fixture also includes `xl/cellimages.xml`, which contains a DrawingML `<xdr:pic>` subtree with a standard
`a:blip r:embed="rId1"` reference, resolved via `xl/_rels/cellimages.xml.rels`:

```xml
<Relationship Id="rId1"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/image1.png"/>
```

## Relationship chain summary (high level)

```text
xl/worksheets/sheet1.xml  cell <c r="A1" vm="1"> / <c r="B1" vm="2">
  -> xl/metadata.xml      valueMetadata bk #0/#1 (selected by vm)
  -> xlrd:rvb i="0"/"1"   selects rich value #0/#1 in richValue.xml
  -> richValue.xml        <v kind="rel">0</v> selects rel-slot 0
  -> richValueRel.xml     rel-slot 0 -> r:id="rId1"
  -> richValueRel.xml.rels rId1 -> ../media/image1.png
  -> xl/media/image1.png

(Additionally present)
xl/cellimages.xml -> xl/_rels/cellimages.xml.rels -> xl/media/image1.png
```

