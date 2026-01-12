# `image-in-cell-richdata.xlsx` (minimal in-cell image fixture — `richValue.xml` 2017 variant)

This fixture demonstrates an **image stored as a cell value** using Excel’s RichData / RichValue pipeline:

* worksheet cell `c/@vm`
* `xl/metadata.xml` (with `futureMetadata name="XLRICHVALUE"` + `xlrd:rvb`)
* `xl/richData/richValue.xml` + `xl/richData/richValueRel.xml`
* `xl/richData/_rels/richValueRel.xml.rels` → `xl/media/image*.png`

This file is intentionally minimal (and is tagged in `docProps/app.xml` as `Application=Formula Fixtures`,
not Excel).

See also:

- [`docs/20-images-in-cells.md`](../../../docs/20-images-in-cells.md)
- [`docs/20-images-in-cells-richdata.md`](../../../docs/20-images-in-cells-richdata.md)
- [`docs/20-xlsx-rich-data.md`](../../../docs/20-xlsx-rich-data.md)
- Real Excel fixtures in this repo:
  - `fixtures/xlsx/basic/image-in-cell.xlsx` (real Excel; `rdRichValue*` variant; no `xl/cellimages.xml`)
  - `fixtures/xlsx/rich-data/images-in-cell.xlsx` (real Excel; unprefixed `richValue*` variant; includes `xl/cellimages.xml`)

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

## `xl/metadata.xml` (`futureMetadata name="XLRICHVALUE"` + `xlrd:rvb i="..."`)

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

* `rc/@v="0"` is a 0-based index into `<futureMetadata name="XLRICHVALUE">`.
* `xlrd:rvb/@i` provides the 0-based rich value index into `xl/richData/richValue.xml`.

## `xl/richData/richValue.xml` (2017 `richdata` namespace)

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv s="0" t="image">
    <v kind="rel">0</v>
  </rv>
</rvData>
```

Notes:

* Namespace: `http://schemas.microsoft.com/office/spreadsheetml/2017/richdata`
* The `<v>` payload is a **relationship-slot index** (0-based) into `xl/richData/richValueRel.xml` (and in
  this fixture it is annotated with `kind="rel"`).

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
<Relationship Id="rId99"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"
              Target="metadata.xml"/>
<Relationship Id="rId4"
              Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue"
              Target="richData/richValue.xml"/>
<Relationship Id="rId5"
              Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel"
              Target="richData/richValueRel.xml"/>
```

## Content types (`[Content_Types].xml`)

This fixture includes explicit `<Override>` entries for the metadata and rich-value parts:

```xml
<Override PartName="/xl/metadata.xml"
          ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
<Override PartName="/xl/richData/richValue.xml"
          ContentType="application/vnd.ms-excel.richvalue+xml"/>
<Override PartName="/xl/richData/richValueRel.xml"
          ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
```

## Relationship chain summary

```text
xl/worksheets/sheet1.xml  cell <c vm="0">
  -> xl/metadata.xml      valueMetadata[0] -> rc/@v = 0 -> futureMetadata[0] -> xlrd:rvb/@i = 0
  -> xl/richData/richValue.xml      rich value #0 (t="image") -> <v kind="rel">0</v> (relationship-slot index)
  -> xl/richData/richValueRel.xml   rel slot #0 -> r:id="rId1"
  -> xl/richData/_rels/richValueRel.xml.rels  rId1 -> ../media/image1.png
  -> xl/media/image1.png
```
