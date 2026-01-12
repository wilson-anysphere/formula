# `image-in-cell-richdata.xlsx` (minimal in-cell image fixture — `richValue.xml` 2017 variant)

This fixture demonstrates an **image stored as a cell value** using Excel’s RichData / RichValue pipeline:

* worksheet cell `c/@vm`
* `xl/metadata.xml` (without `futureMetadata`)
* `xl/richData/richValue.xml` + `xl/richData/richValueRel.xml`
* `xl/richData/_rels/richValueRel.xml.rels` → `xl/media/image*.png`

This file is intentionally minimal (and is tagged in `docProps/app.xml` as `Application=Formula Fixtures`,
not Excel).

See also:

- [`docs/20-images-in-cells.md`](../../../docs/20-images-in-cells.md)
- [`docs/20-images-in-cells-richdata.md`](../../../docs/20-images-in-cells-richdata.md)
- [`docs/20-xlsx-rich-data.md`](../../../docs/20-xlsx-rich-data.md)

## ZIP part inventory (complete)

Output of:

```bash
unzip -l fixtures/xlsx/basic/image-in-cell-richdata.xlsx
```

```text
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

## Worksheet cell (`vm="0"`)

From `xl/worksheets/sheet1.xml`:

```xml
<c r="A1" vm="0"><v>0</v></c>
```

Notes:

* In this fixture, `vm` is **0-based** (`vm="0"` selects the first `<valueMetadata><bk>` record).

## `xl/metadata.xml` (no `futureMetadata`)

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

* There is **no** `<futureMetadata name="XLRICHVALUE">` mapping in this file.
* `rc/@v="0"` appears to directly reference rich value index `0` (0-based) in `xl/richData/richValue.xml`.

## `xl/richData/richValue.xml` (2017 `richdata` namespace)

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv s="0" t="image">
    <v>0</v>
  </rv>
</rvData>
```

Notes:

* Namespace: `http://schemas.microsoft.com/office/spreadsheetml/2017/richdata`
* The `<v>` payload is a **relationship-slot index** (0-based) into `xl/richData/richValueRel.xml`.

## `xl/richData/richValueRel.xml` (2017 `richdata2` namespace)

```xml
<richValueRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>
```

Notes:

* Root local-name: `richValueRel`
* Namespace: `http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2`
* Relationship-slot `0` selects the first `<rel/>`.

## `xl/richData/_rels/richValueRel.xml.rels`

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="../media/image1.png"/>
</Relationships>
```

## Workbook relationships (`xl/_rels/workbook.xml.rels`)

Relevant entries:

```xml
<Relationship Id="rId3"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
              Target="metadata.xml"/>
<Relationship Id="rId4"
              Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue"
              Target="richData/richValue.xml"/>
<Relationship Id="rId5"
              Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel"
              Target="richData/richValueRel.xml"/>
```

## Content types (`[Content_Types].xml`)

This fixture relies on the package default:

```xml
<Default Extension="xml" ContentType="application/xml"/>
```

and includes **no** `<Override>` entries for `xl/metadata.xml` or `xl/richData/*`.

## Relationship chain summary

```text
xl/worksheets/sheet1.xml  cell <c vm="0">
  -> xl/metadata.xml      valueMetadata[0] -> rc/@v = 0
  -> xl/richData/richValue.xml      rich value #0 (t="image") -> <v>0</v> (REL_SLOT)
  -> xl/richData/richValueRel.xml   rel slot #0 -> r:id="rId1"
  -> xl/richData/_rels/richValueRel.xml.rels  rId1 -> ../media/image1.png
  -> xl/media/image1.png
```

