# Excel “Place in Cell” embedded images: OOXML schema + mapping

This document records the **OOXML parts and relationship chain** Excel uses for **embedded images in cells** (Excel UI: *Insert → Pictures → Place in Cell*).

The schema described below was confirmed by generating a minimal `.xlsx` using `rust_xlsxwriter` (and inspecting the resulting package). It is recorded here so future engine/model work can round-trip these files without treating them as “mysterious metadata”.

## High-level mapping chain (cell → image bytes)

In this schema, the *cell value itself is an error* (`#VALUE!`). The **image is attached via value-metadata** which points into Excel’s **Rich Data / Rich Value** parts.

```text
xl/worksheets/sheet1.xml    <c t="e" vm="…"><v>#VALUE!</v></c>
          │ vm (value-metadata index)
          ▼
xl/metadata.xml             <valueMetadata> … <rc t="…" v="…"/> … </valueMetadata>
          │ v (rich value index)
          ▼
xl/richData/rdrichvalue.xml <rv s="…"><v>LocalImageIdentifier</v><v>CalcOrigin</v>…</rv>
          │ LocalImageIdentifier (index into richValueRel list)
          ▼
xl/richData/richValueRel.xml   <rel r:id="rId…"/>
          │ relationship id
          ▼
xl/richData/_rels/richValueRel.xml.rels  Target="../media/imageN.png"
          │
          ▼
xl/media/imageN.png         (actual image bytes)
```

## 1) Worksheet cell encoding: `t="e" vm="1"` + `#VALUE!`

The worksheet stores an error value but attaches “rich value” metadata via the `vm` attribute:

```xml
<row r="1" spans="1:1">
  <c r="A1" t="e" vm="1">
    <v>#VALUE!</v>
  </c>
</row>
```

Notes:

* `t="e"` is the standard SpreadsheetML *error* cell type.
* `vm="1"` is the **value-metadata index** that links the cell to `xl/metadata.xml`.
  * In the observed file, `valueMetadata count="1"` but the cell uses `vm="1"`, which strongly suggests `vm` is **1-based** (with `0` meaning “no value metadata”).

## 2) `xl/_rels/workbook.xml.rels`: required relationships + types

The workbook relationships include standard parts (worksheet/styles/theme) plus the “sheetMetadata” and Rich Data parts.

Relevant relationships (IDs will vary, types/targets are the important part):

```xml
<Relationship Id="rId4"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
  Target="metadata.xml"/>

<Relationship Id="rId5"
  Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel"
  Target="richData/richValueRel.xml"/>

<Relationship Id="rId6"
  Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue"
  Target="richData/rdrichvalue.xml"/>

<Relationship Id="rId7"
  Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure"
  Target="richData/rdrichvaluestructure.xml"/>

<Relationship Id="rId8"
  Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes"
  Target="richData/rdRichValueTypes.xml"/>
```

## 3) `xl/metadata.xml`: `metadataTypes` + `futureMetadata name="XLRICHVALUE"` + `valueMetadata`

`xl/metadata.xml` is the “bridge” between worksheet cells (`vm="…"`) and Rich Data rich values (`xl/richData/rdrichvalue.xml`).

Minimal structure from the observed file:

```xml
<metadata
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">

  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000" .../>
  </metadataTypes>

  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}">
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

Interpretation of the important fields (based on observed behavior/structure):

* `metadataTypes/metadataType name="XLRICHVALUE"` declares a metadata “type” used for Rich Values.
* `valueMetadata` is the per-cell table indexed by the cell’s `vm="…"`.
* `<rc t="1" v="0"/>`:
  * `t="1"` points at the `XLRICHVALUE` metadata type (again: appears **1-based**).
  * `v="0"` points at a Rich Value binding that ultimately selects **rich value index 0** in `xl/richData/rdrichvalue.xml`.

## 4) `xl/richData/rdrichvaluestructure.xml`: `_localImage` schema + keys

Rich Values are typed. For “Place in Cell” images, the structure type is `_localImage` and the keys include a relationship lookup key.

```xml
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <s t="_localImage">
    <k n="_rvRel:LocalImageIdentifier" t="i"/>
    <k n="CalcOrigin" t="i"/>
  </s>
</rvStructures>
```

Notes:

* The `<k>` elements define the **ordered key list** for values in `rdrichvalue.xml`.
* `_rvRel:LocalImageIdentifier` is the critical pointer: it indexes into `xl/richData/richValueRel.xml`.
* `CalcOrigin` is an integer flag indicating where the image came from (see next section).

## 5) `xl/richData/rdrichvalue.xml`: value ordering + `CalcOrigin` (5 vs 6)

The rich value records are stored in `rdrichvalue.xml`. The `<v>` elements are **positional**, matching the key order from `rdrichvaluestructure.xml`.

Observed minimal file:

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <rv s="0">
    <v>0</v>
    <v>5</v>
  </rv>
</rvData>
```

Interpretation:

* `<rv s="0">` means “use structure 0”, i.e. the first `<s>` in `rdrichvaluestructure.xml` (here: `_localImage`).
* The two values correspond to:
  1. `_rvRel:LocalImageIdentifier` = `0` → first `<rel>` in `xl/richData/richValueRel.xml`
  2. `CalcOrigin` = `5`

`CalcOrigin` values:

* `5` has been observed for **embedded local images written into the file** (the “Place in Cell” scenario generated by `rust_xlsxwriter`).
* `6` has been observed in other Excel-produced files and appears to correspond to a **calculation-generated** image (for example, an image originating from an Excel function/result rather than an explicitly embedded local image).

The exact enum definition is not documented publicly by Microsoft; treat `CalcOrigin` as an opaque Excel flag but preserve it when round-tripping.

## 6) `xl/richData/richValueRel.xml` + `.rels`: mapping to `xl/media/*`

`richValueRel.xml` provides an ordered list of relationships used by rich values.

```xml
<richValueRels
  xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRels>
```

The relationship ID is resolved via:

`xl/richData/_rels/richValueRel.xml.rels`

```xml
<Relationship Id="rId1"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  Target="../media/image1.png"/>
```

This is how `_rvRel:LocalImageIdentifier = 0` ultimately maps to an actual file in `xl/media/`.

## 7) `[Content_Types].xml` overrides for “Place in Cell” image parts

The following overrides are required for the additional parts involved in this schema (in addition to the normal workbook/worksheet/styles/theme overrides):

```xml
<Override PartName="/xl/metadata.xml"
  ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>

<Override PartName="/xl/richData/rdRichValueTypes.xml"
  ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>

<Override PartName="/xl/richData/rdrichvalue.xml"
  ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>

<Override PartName="/xl/richData/rdrichvaluestructure.xml"
  ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>

<Override PartName="/xl/richData/richValueRel.xml"
  ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
```

## 8) Important: **no `xl/cellimages.xml`** (for Place in Cell)

For the verified “Place in Cell” scenario above, Excel/rust_xlsxwriter **does not** create or reference
`xl/cellImages.xml` (or the lowercase variant `xl/cellimages.xml`).

Instead, it uses:

* `xl/metadata.xml` (sheet metadata)
* `xl/richData/*` (rich value structures/values + relationship indirection)
* `xl/media/*` (actual image bytes)

Open question:

* Excel has multiple image-related features (floating drawings, background images, legacy objects, the `IMAGE()` function, “data types”, etc.). It is still possible that **other** Excel scenarios use `xl/cellimages.xml` or additional parts. If/when we encounter such files in the corpus, we should extend this document with concrete samples.
