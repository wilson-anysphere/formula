# `images-in-cell.xlsx` (Excel “Place in Cell” fixture with `cellimages.xml` + `richValue*`)

This fixture is an `.xlsx` produced by modern Excel that demonstrates **images stored as cell values**
via the Rich Data / Rich Value pipeline (**cell `vm=` metadata** + `xl/metadata.xml` + `xl/richData/*`)
**and** includes the dedicated **cell image store** part `xl/cellimages.xml`.

See also:

- [`docs/20-images-in-cells.md`](../../../docs/20-images-in-cells.md)
- [`docs/20-images-in-cells-richdata.md`](../../../docs/20-images-in-cells-richdata.md)
- [`docs/20-xlsx-rich-data.md`](../../../docs/20-xlsx-rich-data.md)
- Related real Excel fixture (different on-disk shape): `fixtures/xlsx/basic/image-in-cell.xlsx` (uses the
  `rdRichValue*` variant and does not include `xl/cellimages.xml`; see `fixtures/xlsx/basic/image-in-cell.md`).

## Provenance: confirm this file was saved by Excel

`docProps/app.xml` contains:

```xml
<Application>Microsoft Excel</Application>
<AppVersion>16.0300</AppVersion>
```

## ZIP part inventory (complete)

Output of:

```bash
unzip -Z1 fixtures/xlsx/rich-data/images-in-cell.xlsx | sort
```

```text
[Content_Types].xml
_rels/.rels
docProps/app.xml
docProps/core.xml
xl/_rels/cellimages.xml.rels
xl/_rels/metadata.xml.rels
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

## Workbook relationships (`xl/_rels/workbook.xml.rels`)

This file links `metadata.xml` and `cellimages.xml` directly from the workbook:

```xml
<Relationship Id="rId3"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"
              Target="metadata.xml"/>
<Relationship Id="rId4"
              Type="http://schemas.microsoft.com/office/2019/relationships/cellimages"
              Target="cellimages.xml"/>
```

Notably, the `xl/richData/*` parts are related from `xl/_rels/metadata.xml.rels` (below), rather than
directly from `xl/_rels/workbook.xml.rels`.

## `xl/_rels/metadata.xml.rels` (metadata → richData relationships)

```xml
<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2017/relationships/richValue"
              Target="richData/richValue.xml"/>
<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2017/relationships/richValueRel"
              Target="richData/richValueRel.xml"/>
<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/relationships/richValueTypes"
              Target="richData/richValueTypes.xml"/>
<Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/relationships/richValueStructure"
              Target="richData/richValueStructure.xml"/>
```

## Worksheet cell (`vm="1" cm="1"`)

From `xl/worksheets/sheet1.xml`:

```xml
<c r="A1" vm="1" cm="1"><v>0</v></c>
```

Notes:

- This fixture uses a **numeric cached `<v>0</v>`** (no `t="e"` / `#VALUE!`).
- In this file, `vm` and `cm` are **1-based** indices.

## `xl/metadata.xml` (`futureMetadata` + `xlrd:rvb` lookup table)

```xml
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>

  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>

  <valueMetadata count="1">
    <bk><rc t="1" v="0"/></bk>
  </valueMetadata>

  <cellMetadata count="1">
    <bk><rc t="1" v="0"/></bk>
  </cellMetadata>
</metadata>
```

## `xl/richData/*` (“richValue*” naming scheme)

This file uses the unprefixed **`richValue*`** parts (not the `rdRichValue*` naming scheme).

### `xl/richData/richValueTypes.xml` + `xl/richData/richValueStructure.xml`

```xml
<rvTypes xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <types>
    <type id="0" name="com.microsoft.excel.image" structure="s_image"/>
  </types>
</rvTypes>
```

```xml
<rvStruct xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <structures>
    <structure id="s_image">
      <member name="imageRel" kind="rel"/>
    </structure>
  </structures>
</rvStruct>
```

### `xl/richData/richValue.xml` (rich value instance table)

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">0</v>
    </rv>
  </values>
</rvData>
```

### `xl/richData/richValueRel.xml` + `.rels` (relationship-slot indirection)

The rich value payload stores an integer relationship-slot index (`0`) into the ordered list in
`richValueRel.xml`:

```xml
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId1"/>
  </rels>
</rvRel>
```

And `xl/richData/_rels/richValueRel.xml.rels` resolves that `rId` to the actual media part:

```xml
<Relationship Id="rId1"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="../media/image1.png"/>
```

## `xl/cellimages.xml` (cell image store part)

This part stores a DrawingML `<xdr:pic>` subtree referencing the image bytes via `r:embed`:

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

And `xl/_rels/cellimages.xml.rels` resolves that `rId1` to the media part:

```xml
<Relationship Id="rId1"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/image1.png"/>
```

## Relationship chain summary

```text
xl/worksheets/sheet1.xml  cell <c r="A1" vm="1">
  -> xl/metadata.xml      valueMetadata bk #0 (selected by vm="1" in this file) -> rc v="0" -> futureMetadata bk #0 -> xlrd:rvb i="0"
  -> xl/richData/richValue.xml     rich value #0 -> <v kind="rel">0</v> (REL_SLOT)
  -> xl/richData/richValueRel.xml  rel slot #0 -> r:id="rId1"
  -> xl/richData/_rels/richValueRel.xml.rels  rId1 -> ../media/image1.png
  -> xl/media/image1.png

(Additionally present)
xl/cellimages.xml -> xl/_rels/cellimages.xml.rels -> xl/media/image1.png
```
