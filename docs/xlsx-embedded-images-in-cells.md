# Excel “Place in Cell” embedded images: OOXML schema + mapping

This document records a **concrete OOXML parts + relationship chain** for **embedded images in cells**
(Excel UI: *Insert → Pictures → Place in Cell*).

The schema described below was confirmed by inspecting both:

* the real Excel-generated fixture workbook `fixtures/xlsx/basic/image-in-cell.xlsx` (notes in `fixtures/xlsx/basic/image-in-cell.md`), and
* a minimal `.xlsx` generated using `rust_xlsxwriter` (see `crates/formula-xlsx/tests/embedded_images_place_in_cell_roundtrip.rs`).

It is recorded here so future engine/model work can round-trip these files without treating them as “mysterious metadata”.

See also (broader context + variant coverage):

- [`docs/20-images-in-cells.md`](./20-images-in-cells.md) — overall “Images in Cell” packaging + round-trip constraints
- [`docs/20-images-in-cells-richdata.md`](./20-images-in-cells-richdata.md) — RichData (`richValue*` / `rdrichvalue*`) tables + index-base notes
- [`docs/20-xlsx-rich-data.md`](./20-xlsx-rich-data.md) — shorter overview of the rich-data wiring

Related in-repo references:

* Fixture workbook + notes:
  * `fixtures/xlsx/basic/image-in-cell.xlsx`
  * `fixtures/xlsx/basic/image-in-cell.md`
* Preservation/regression test that generates a “Place in Cell” workbook via `rust_xlsxwriter`:
  * `crates/formula-xlsx/tests/embedded_images_place_in_cell_roundtrip.rs`
* Relevant Formula parsing/extraction helpers:
  * `crates/formula-xlsx/src/rich_data/mod.rs` (`extract_rich_cell_images`)
  * `crates/formula-xlsx/src/rich_data/metadata.rs` (parsing `vm` → rich value index)

## High-level mapping chain (cell → image bytes)

In this schema, the *cell value itself is an error* (`#VALUE!`). The **image is attached via value-metadata** which points into Excel’s **Rich Data / Rich Value** parts.

```text
xl/worksheets/sheet1.xml    <c t="e" vm="…"><v>#VALUE!</v></c>
          │ vm (value-metadata index)
          ▼
xl/metadata.xml             valueMetadata[bk] → <rc t="…" v="…"/>
          │ v (futureMetadata bk index)
          ▼
xl/metadata.xml             futureMetadata(XLRICHVALUE)[bk] → <xlrd:rvb i="…"/>
          │ i (rich value index)
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

### Index/indirection summary

This schema uses multiple indices; the most important ones are:

| Field | Where it lives | Meaning | Index-base / ordering notes |
|------:|----------------|---------|-----------------------------|
| `c/@vm` | `xl/worksheets/sheetN.xml` | Selects a record in `xl/metadata.xml/valueMetadata` | **Ambiguous base** (0- or 1-based). Preserve and resolve via `metadata.xml`. |
| `rc/@t` | `xl/metadata.xml` (`valueMetadata/bk/rc`) | Index into `metadataTypes` (`XLRICHVALUE`) | Appears **1-based** (`t="1"`) in the Excel-generated fixtures in this repo, but other workbooks/tests have been observed to use **0-based** indexing; treat as ambiguous. |
| `rc/@v` | `xl/metadata.xml` (`valueMetadata/bk/rc`) | Index into `futureMetadata name="XLRICHVALUE"` `bk` list | Appears **0-based** in the Excel fixtures in this repo; other indexing schemes (including 1-based) may exist in the wild/tests. |
| `xlrd:rvb/@i` | `xl/metadata.xml` (`futureMetadata/XLRICHVALUE` extension) | Index into `xl/richData/rdrichvalue.xml` `<rv>` list | **0-based** rich value index. |
| `_rvRel:LocalImageIdentifier` | `xl/richData/rdrichvalue.xml` | Relationship-slot index | The relationship slot index is carried by the `<v>` corresponding to the `_rvRel:LocalImageIdentifier` key in `rdrichvaluestructure.xml` (do not assume it is always the first `<v>`; in the observed `_localImage` structure it is first because of the key order). |
| `rel/@r:id` | `xl/richData/richValueRel.xml` | Relationship ID string | Resolve via `.rels` by matching `Id="..."` (not by element order). |

### Practical detection (distinguishing “Place in Cell” images from other `vm` uses)

Not every `vm="…"` cell is necessarily an in-cell image; it just means “this cell has value metadata”.
For the “Place in Cell” local-image shape documented here, a robust detector should:

1. Require `c/@vm` on the cell.
2. Resolve `vm` into `xl/metadata.xml/valueMetadata`.
   * The `vm` base is ambiguous; be prepared to handle both 0-based and 1-based indexing.
3. Confirm the metadata type is `XLRICHVALUE` (`metadataTypes/metadataType name="XLRICHVALUE"`; referenced by `rc/@t`).
4. Resolve the rich value index:
   * `rc/@v` selects the `futureMetadata name="XLRICHVALUE"` `<bk>`.
   * Within that `<bk>`, find `<xlrd:rvb i="…"/>` and use `@i` as the rich value index.
5. Load the rich value from `xl/richData/rdrichvalue.xml` at that index and confirm its structure:
   * `rv/@s` selects a structure in `xl/richData/rdrichvaluestructure.xml`.
   * For images, the structure has `t="_localImage"` and includes keys `_rvRel:LocalImageIdentifier` and `CalcOrigin`.
6. Interpret the rich value payload positionally:
   * The `<v>` corresponding to `_rvRel:LocalImageIdentifier` is the relationship slot index (in the observed fixture it is the first `<v>` because `_rvRel:LocalImageIdentifier` is the first `<k>` in `rdrichvaluestructure.xml`).
   * Second `<v>` = `CalcOrigin` (preserve as an opaque Excel flag).

The cached cell representation in this fixture (`t="e"` + `#VALUE!`) is a strong signal for this
`rdRichValue*` variant, but other real Excel workbooks can use other cached-value encodings (e.g.
placeholder numeric `<v>0</v>`). The structure check above is the more semantically reliable way to
identify local embedded images.

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
  * In this sample, `valueMetadata count="1"` but the cell uses `vm="1"`, which suggests `vm` is **1-based** here.
    Other workbooks can use `vm="0"` for the first record; treat `vm` as ambiguous and resolve best-effort (see [`docs/20-images-in-cells-richdata.md`](./20-images-in-cells-richdata.md)).

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

Minimal structure from the smallest observed file (the `rust_xlsxwriter`-generated sample with a single
image). The real Excel fixture with **multiple images** is shown in [section 9](#9-real-excel-example-multiple-cells--multiple-images-fixture).

```xml
<metadata
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">

  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000"
                  copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1"
                  clearFormats="1" clearComments="1" assign="1" coerce="1"/>
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
  * `t="1"` points at the `XLRICHVALUE` metadata type (in this fixture, `rc/@t` is **1-based**).
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

Observed minimal file (single image). The real Excel fixture with multiple images is shown in
[section 9](#9-real-excel-example-multiple-cells--multiple-images-fixture).

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

* `5` has been observed for **embedded local images written into the file** (the “Place in Cell” scenario), in both:
  * a `rust_xlsxwriter`-generated file, and
  * the Excel-generated fixture `fixtures/xlsx/basic/image-in-cell.xlsx`.
* `6` is not currently observed in the in-repo fixtures; if encountered, treat it as an opaque Excel flag and
  preserve it when round-tripping.

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

Important indexing note:

* `_rvRel:LocalImageIdentifier` is a **0-based index into the `<rel>` list in `richValueRel.xml`**.
* The OPC relationships file (`richValueRel.xml.rels`) is **not** order-sensitive; it is a map from `Id="rId…"` to a `Target=…`.
  * In the Excel-produced fixture `fixtures/xlsx/basic/image-in-cell.xlsx`, the `<rel>` list is
    `rId1`, `rId2`, but the `.rels` file lists `<Relationship Id="rId2" …/>` before `<Relationship Id="rId1" …/>`.
    Consumers must follow the `richValueRel.xml` list order first, then resolve by relationship ID.

The relationship ID is resolved via:

`xl/richData/_rels/richValueRel.xml.rels`

```xml
<Relationship Id="rId1"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  Target="../media/image1.png"/>
```

This is how `_rvRel:LocalImageIdentifier = 0` ultimately maps to an actual file in `xl/media/`.

## 7) `[Content_Types].xml`: defaults + overrides (complete set for the sample)

The rust_xlsxwriter-generated “Place in Cell” workbook used for schema verification contains the
following content types (formatted for readability):

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels"
           ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>

  <Override PartName="/docProps/app.xml"
            ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml"
            ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>

  <Override PartName="/xl/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/workbook.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>

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
</Types>
```

Notes:

* In this repo’s fixtures, explicit overrides are present for these parts; other producers may vary.
  Preserve whatever the source file uses when round-tripping.
* The image bytes are stored under `xl/media/*.png` (or other formats) and typically use the relevant
  `<Default Extension="…">` entry rather than an explicit `<Override>`.

## 8) Cell image store part (`xl/cellimages.xml`) is optional

For the verified “Place in Cell” scenario above (the `rdRichValue*` schema), Excel/rust_xlsxwriter **does not**
create or reference `xl/cellImages.xml` / `xl/cellimages.xml`.

Instead, it uses:

* `xl/metadata.xml` (sheet metadata)
* `xl/richData/*` (rich value structures/values + relationship indirection)
* `xl/media/*` (actual image bytes)

However, other real Excel workbooks **do** include a `cellimages` store part. In this repo:

* `fixtures/xlsx/rich-data/images-in-cell.xlsx` (notes in `fixtures/xlsx/rich-data/images-in-cell.md`) contains:
  - `xl/cellimages.xml` + `xl/_rels/cellimages.xml.rels` + `xl/media/*`, in addition to the RichData tables.

So: do not assume `xl/cellimages.xml` is absent. Treat it as an optional workbook-level image store and preserve it
byte-for-byte when round-tripping.

## 9) Real Excel example: multiple cells + multiple images (fixture)

The fixture `fixtures/xlsx/basic/image-in-cell.xlsx` is a **real Excel-generated** workbook with **two**
images stored as cell values (and the same RichData wiring described above).

### Worksheet cells (`xl/worksheets/sheet1.xml`)

In this file, multiple cells can share the same image binding:

```xml
<c r="B2" t="e" vm="1"><v>#VALUE!</v></c>
<c r="B3" t="e" vm="1"><v>#VALUE!</v></c>
<c r="B4" t="e" vm="2"><v>#VALUE!</v></c>
```

Interpretation:

* `B2` and `B3` both use `vm="1"` → both reference the **same** rich value (and thus the same image).
* `B4` uses `vm="2"` → references a **different** rich value (a different image).

### Value metadata (`xl/metadata.xml`)

The fixture has two rich value bindings:

```xml
<futureMetadata name="XLRICHVALUE" count="2">
  <!-- Note: the `xlrd` prefix is declared on the `metadata` root element. -->
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
```

### Rich values (`xl/richData/rdrichvalue.xml`)

The two rich values reference relationship slots `0` and `1` (and both have `CalcOrigin=5`):

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="2">
  <rv s="0"><v>0</v><v>5</v></rv>
  <rv s="0"><v>1</v><v>5</v></rv>
</rvData>
```

### Relationship slots + media mapping

`richValueRel.xml` defines the slot ordering:

```xml
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
  <rel r:id="rId2"/>
</richValueRels>
```

And the `.rels` maps those `rId`s to actual media parts (note: `.rels` ordering can differ):

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
