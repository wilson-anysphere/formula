# `image-in-cell.xlsx` (Excel "image in cell"/rich value fixture)

This fixture is an XLSX produced by modern Excel that demonstrates **images stored as cell values** via the **Rich Data / Rich Value** pipeline (cell `vm=` attribute + `xl/metadata.xml` + `xl/richData/*`), referencing image binaries in `xl/media/*`.

Notably, this fixture **does _not_ contain** an `xl/cellimages.xml` part; instead, `xl/richData/_rels/richValueRel.xml.rels` points directly at `xl/media/image*.png`.

See also:

- [`docs/20-images-in-cells.md`](../../../docs/20-images-in-cells.md) — overall “Images in Cell” packaging + round-trip constraints
- [`docs/20-images-in-cells-richdata.md`](../../../docs/20-images-in-cells-richdata.md) — RichData (`richValue*` / `rdrichvalue*`) tables + index-base notes
- [`docs/xlsx-embedded-images-in-cells.md`](../../../docs/xlsx-embedded-images-in-cells.md) — concrete schema walkthrough (relationships, content types, `_localImage` keys, etc.)
- Related real Excel fixture (different on-disk shape): `fixtures/xlsx/rich-data/images-in-cell.xlsx` (includes
  `xl/cellimages.xml` + unprefixed `richValue*` tables; see `fixtures/xlsx/rich-data/images-in-cell.md`).

## Provenance: confirm this file was saved by Excel

`docProps/app.xml` contains:

```xml
<Application>Microsoft Excel</Application>
<AppVersion>16.0300</AppVersion>
```

## ZIP part inventory (relevant parts only)

Output of:

```bash
unzip -l fixtures/xlsx/basic/image-in-cell.xlsx | rg 'cellimages|richData|metadata|media|rels|\\[Content_Types\\]'
```

```text
4:     1919  1980-01-01 00:00   [Content_Types].xml
5:      588  1980-01-01 00:00   _rels/.rels
7:     1408  1980-01-01 00:00   xl/_rels/workbook.xml.rels
12:      617  1980-01-01 00:00   xl/media/image1.png
13:      808  1980-01-01 00:00   xl/media/image2.png
14:      322  1980-01-01 00:00   xl/worksheets/_rels/sheet1.xml.rels
15:      814  1980-01-01 00:00   xl/metadata.xml
16:      278  1980-01-01 00:00   xl/richData/richValueRel.xml
17:      218  1980-01-01 00:00   xl/richData/rdrichvalue.xml
18:      258  1980-01-01 00:00   xl/richData/rdrichvaluestructure.xml
19:     1187  1980-01-01 00:00   xl/richData/rdRichValueTypes.xml
23:      427  1980-01-01 00:00   xl/richData/_rels/richValueRel.xml.rels
```

## Workbook-level relationships (`xl/_rels/workbook.xml.rels`)

Relevant `<Relationship>` entries (IDs may vary; `Type` + `Target` are the important bits):

```xml
<Relationship Id="rId5"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
              Target="metadata.xml"/>
<Relationship Id="rId6"
              Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel"
              Target="richData/richValueRel.xml"/>
<Relationship Id="rId7"
              Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue"
              Target="richData/rdrichvalue.xml"/>
<Relationship Id="rId8"
              Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure"
              Target="richData/rdrichvaluestructure.xml"/>
<Relationship Id="rId9"
              Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes"
              Target="richData/rdRichValueTypes.xml"/>
```

## Content types (`[Content_Types].xml`)

Relevant `<Override>` entries:

```xml
<Override PartName="/xl/metadata.xml"
          ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
<Override PartName="/xl/richData/richValueRel.xml"
          ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
<Override PartName="/xl/richData/rdrichvalue.xml"
          ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>
<Override PartName="/xl/richData/rdrichvaluestructure.xml"
          ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>
<Override PartName="/xl/richData/rdRichValueTypes.xml"
          ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>
```

## Worksheet cell XML (`vm=` on `<c>`)

From `xl/worksheets/sheet1.xml`:

```xml
<row r="2" spans="1:2" x14ac:dyDescent="0.3">
  <c r="A2"><v>1</v></c>
  <c r="B2" t="e" vm="1"><v>#VALUE!</v></c>
</row>
...
<row r="4" spans="1:2" x14ac:dyDescent="0.3">
  <c r="A4"><v>3</v></c>
  <c r="B4" t="e" vm="2"><v>#VALUE!</v></c>
</row>
```

Cells `B2/B3` share `vm="1"` and `B4` uses `vm="2"`.

## `xl/metadata.xml` (value metadata → rich value bundle index)

From `xl/metadata.xml`:

```xml
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000" .../>
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

## `xl/richData/*` and image relationship

`xl/richData/rdrichvaluestructure.xml` defines a `_localImage` structure with a key named `_rvRel:LocalImageIdentifier`:

```xml
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <s t="_localImage">
    <k n="_rvRel:LocalImageIdentifier" t="i"/>
    <k n="CalcOrigin" t="i"/>
  </s>
</rvStructures>
```

`xl/richData/richValueRel.xml` contains two `<rel r:id="…"/>` entries:

```xml
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
  <rel r:id="rId2"/>
</richValueRels>
```

Those IDs resolve via `xl/richData/_rels/richValueRel.xml.rels`:

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="../media/image2.png"/>
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="../media/image1.png"/>
</Relationships>
```

Important indexing note:

* `xl/richData/richValueRel.xml` is an **ordered** `<rel>` list (relationship-slot table). Rich values index into
  this list by integer slot.
* `xl/richData/_rels/richValueRel.xml.rels` is an **unordered** map from relationship ID (`rId*`) to `Target`.
  Do not assume it has the same ordering as the `<rel>` list.

## Relationship chain (high level)

```text
xl/worksheets/sheet1.xml
  cell <c r="B2" vm="1"> / <c r="B4" vm="2">
    │
    └─(vm index)→ xl/metadata.xml
                  valueMetadata[bk] → futureMetadata(XLRICHVALUE) → xlrd:rvb i="…"
                    │
                    └─(rvb index)→ xl/richData/rdrichvalue.xml  (rvData/rv)
                                   xl/richData/rdrichvaluestructure.xml (structure "_localImage")
                                     │
                                     └─(_rvRel:LocalImageIdentifier)→ xl/richData/richValueRel.xml (<rel r:id="…"/>)
                                                                      xl/richData/_rels/richValueRel.xml.rels
                                                                        │
                                                                        └→ xl/media/image1.png, xl/media/image2.png
```
