# Excel RichData (`richValue*`) parts for Images-in-Cell (`IMAGE()` / “Place in Cell”)

Excel’s “Images in Cell” feature (insert picture → **Place in Cell**, and the `IMAGE()` function) is backed by a **RichData / RichValue** subsystem. Rather than embedding image references directly in worksheet cell XML, Excel stores *typed rich value instances* in workbook-level parts under `xl/richData/`, then attaches cells to those instances via metadata.

This note documents the **expected part set**, the **role of each part**, and the **minimal XML shapes** needed to parse/write Excel-generated files.

For the overall “images in cells” packaging overview (including the optional `xl/cellImages.xml` store part (sometimes `xl/cellimages.xml`), `xl/metadata.xml`,
and current Formula status/tests), see: [20-images-in-cells.md](./20-images-in-cells.md).

For a **concrete, confirmed** “Place in Cell” (embedded local image) package shape (including the exact
`rdrichvalue*` structure keys like `_rvRel:LocalImageIdentifier` and the `CalcOrigin` field), see:

- [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md)

> Status: best-effort reverse engineering. This repo contains **real Excel** “Place in Cell” fixtures for
> both the `rdRichValue*` variant (`fixtures/xlsx/basic/image-in-cell.xlsx`) and the `richValue*` variant
> (`fixtures/xlsx/rich-data/images-in-cell.xlsx`, which also includes `xl/cellimages.xml`), plus
> **synthetic** fixtures used by tests (`fixtures/xlsx/basic/image-in-cell-richdata.xlsx`,
> `fixtures/xlsx/rich-data/richdata-minimal.xlsx`) — see
> [Observed in fixtures](#observed-in-fixtures-in-repo). Exact namespaces / relationship-type URIs may
> still vary by Excel version; preserve unknown attributes and namespaces when round-tripping.

---

## Expected part set (workbook-level)

When a workbook contains at least one RichData value (including images-in-cell), Excel typically adds:

```
xl/
  richData/
    richValue.xml              # or: richValues.xml (naming varies); or: rdrichvalue.xml
    richValueRel.xml
    richValueTypes.xml        # optional (not present in all workbooks); or: rdRichValueTypes.xml
    richValueStructure.xml    # optional (not present in all workbooks); or: rdrichvaluestructure.xml
  richData/_rels/
    richValueRel.xml.rels   # required if richValueRel.xml contains r:id entries
```

Notes:

* The *minimum* observed set for a simple in-cell image can be smaller. For example,
  `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` includes:
  * `xl/richData/richValue.xml`
  * `xl/richData/richValueRel.xml`
  * `xl/richData/_rels/richValueRel.xml.rels`
  and omits `richValueTypes.xml` / `richValueStructure.xml`.
* For linked data types and richer payloads, Excel is expected to add the supporting “types” and
  “structure” tables; treat their presence as feature-dependent.
  * For example, the real Excel fixture `fixtures/xlsx/rich-data/images-in-cell.xlsx` includes
    `richValueTypes.xml` and `richValueStructure.xml`.
* File naming varies across producers (and even across Excel builds):
  * “Excel-like” naming: `richValue.xml`, `richValueTypes.xml`, `richValueStructure.xml`
  * Plural “richValues” naming (observed in tests; not currently observed in the Excel fixtures in this repo):
    * `richValues.xml`, `richValues1.xml`, ...
  * “rdRichValue” naming (observed in the real Excel fixture `fixtures/xlsx/basic/image-in-cell.xlsx`
    and in `rust_xlsxwriter` output in this repo):
    * `rdrichvalue.xml`
    * `rdrichvaluestructure.xml`
    * `rdRichValueTypes.xml` (note casing)
  For robust parsing, prefer relationship discovery + local-name matching rather than hardcoding a single
  filename spelling/casing.

## Observed in fixtures (in-repo)

This repo includes **real Excel fixtures** for multiple "image in cell" / rich-data encodings, plus
**synthetic** fixtures used for regression tests:

* **Real Excel**: `fixtures/xlsx/rich-data/images-in-cell.xlsx` — full `richValue*` part set **plus**
  `xl/cellimages.xml`.
* **Real Excel**: `fixtures/xlsx/basic/image-in-cell.xlsx` — `rdRichValue*` variant that uses a
  **structure table** (`rdrichvaluestructure.xml`) to assign meanings to positional `<v>` fields.
* **Synthetic (Formula fixture)**: `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` — minimal
  `richValue.xml` + `richValueRel.xml` variant (no `richValueTypes.xml` / `richValueStructure.xml`).
* **Synthetic (Formula fixture)**: `fixtures/xlsx/rich-data/richdata-minimal.xlsx` — minimal full
  `richValue*` part set (includes `richValueTypes.xml` / `richValueStructure.xml`) used by tests.

### Fixture: `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` (`richValue.xml` + `richValueRel.xml` 2017 variant)

See also: [`fixtures/xlsx/basic/image-in-cell-richdata.md`](../fixtures/xlsx/basic/image-in-cell-richdata.md) (walkthrough of this fixture).

**Parts present** (complete inventory from `unzip -l`):

```
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

**Sheet cell metadata (`vm`)**

From `xl/worksheets/sheet1.xml`:

```xml
<c r="A1" vm="0"><v>0</v></c>
```

Notes:

* In this fixture, `vm` is **0-based** (`vm="0"` for the first `valueMetadata` record).

**`xl/metadata.xml` (`futureMetadata name="XLRICHVALUE"` + `xlrd:rvb i="..."`)**

Exact shape (note the `xlrd` namespace and the `<futureMetadata name="XLRICHVALUE">` table):

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

* `rc/@t="1"` selects `metadataType name="XLRICHVALUE"`.
* `rc/@v` is a **0-based index into** the `<futureMetadata name="XLRICHVALUE">` `<bk>` list.
* `<xlrd:rvb i="N"/>` provides the **0-based rich value index** (`N`) into `xl/richData/richValue.xml`.

**`xl/richData/richValue.xml` namespace + payload**

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv s="0" t="image">
    <v kind="rel">0</v>
  </rv>
</rvData>
```

Notes:

* Namespace is **`…/2017/richdata`**.
* The payload `<v>` is an integer **relationship-slot index** (0-based) into
  `xl/richData/richValueRel.xml`.
  * In this fixture the `<v>` carries `kind="rel"`.
  * In general the shape is: `<rv t="image"><v kind="rel">REL_SLOT</v></rv>`.
  * In this fixture, `REL_SLOT = 0`.

**`xl/richData/richValueRel.xml` namespace**

```xml
<richValueRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>
```

Notes:

* Root local-name is `richValueRel`.
* Namespace is **`…/2017/richdata2`**.
* Relationship slots are positional: slot `0` selects the first `<rel/>`.

**Workbook relationships**

From `xl/_rels/workbook.xml.rels` (excerpt):

* `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"`
  → `Target="metadata.xml"`
* `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue"`
  → `Target="richData/richValue.xml"`
* `Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel"`
  → `Target="richData/richValueRel.xml"`

`[Content_Types].xml` in this fixture includes explicit overrides for `xl/metadata.xml` and the rich value
parts, e.g.:

* `/xl/metadata.xml` → `application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml`
* `/xl/richData/richValue.xml` → `application/vnd.ms-excel.richvalue+xml`
* `/xl/richData/richValueRel.xml` → `application/vnd.ms-excel.richvaluerel+xml`

### Fixture: `fixtures/xlsx/rich-data/images-in-cell.xlsx` (Excel `richValue*` + `cellimages.xml`)

This is a modern Excel workbook demonstrating an image stored as a cell value using the unprefixed
`richValue*` table naming scheme **and** a `xl/cellimages.xml` store part.

See also: [`fixtures/xlsx/rich-data/images-in-cell.md`](../fixtures/xlsx/rich-data/images-in-cell.md) (walkthrough of this fixture).

**Parts present** (complete inventory from `unzip -l`):

```
[Content_Types].xml
_rels/.rels
docProps/core.xml
docProps/app.xml
xl/workbook.xml
xl/_rels/workbook.xml.rels
xl/worksheets/sheet1.xml
xl/styles.xml
xl/cellimages.xml
xl/_rels/cellimages.xml.rels
xl/metadata.xml
xl/_rels/metadata.xml.rels
xl/richData/richValue.xml
xl/richData/richValueRel.xml
xl/richData/richValueTypes.xml
xl/richData/richValueStructure.xml
xl/richData/_rels/richValueRel.xml.rels
xl/media/image1.png
```

**Sheet cell metadata (`vm`, `cm`)**

From `xl/worksheets/sheet1.xml`:

```xml
<c r="A1" vm="1" cm="1"><v>0</v></c>
```

Notes:

 * In this fixture, `vm` and `cm` are **1-based** (`vm="1"` selects the first `<valueMetadata><bk>`).
* The cached cell value is a placeholder numeric `<v>0</v>` (no `t="e"`/`#VALUE!`); the image binding still
  comes from `vm`/`cm` + `xl/metadata.xml`.

**`xl/metadata.xml` (`futureMetadata name="XLRICHVALUE"` + `xlrd:rvb i="..."`, plus `cellMetadata`)**

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

Notes:

* In this fixture, `rc/@t="1"` is a 1-based index into `<metadataTypes>` (0-based is also observed in the wild/tests).
* `rc/@v="0"` is a 0-based index into the `<futureMetadata name="XLRICHVALUE">` `<bk>` list.
* `xlrd:rvb/@i="0"` is the **0-based rich value index** into `xl/richData/richValue.xml`.

**`xl/richData/richValueTypes.xml` and `xl/richData/richValueStructure.xml`**

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

**`xl/richData/richValue.xml` (image rich value)**

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">0</v>
    </rv>
  </values>
</rvData>
```

Notes:

* `<rv>` is keyed by a numeric `type="0"` which maps into `richValueTypes.xml`.
* `<v kind="rel">0</v>` stores the **0-based relationship slot index** into `xl/richData/richValueRel.xml`.

**`xl/richData/richValueRel.xml` (relationship-slot table; `rvRel` variant)**

```xml
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId1"/>
  </rels>
</rvRel>
```

Notes:

* Root local-name is `rvRel`.
* Namespace is **`…/2017/richdata`** (not `…/2017/richdata2`).
* Relationship slots are positional: slot `0` selects the first `<rel/>` (here: `rId1`).

**Workbook + metadata relationships**

From `xl/_rels/workbook.xml.rels` (excerpt):

* `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"`
  → `Target="metadata.xml"`
* `Type="http://schemas.microsoft.com/office/2019/relationships/cellimages"`
  → `Target="cellimages.xml"`

From `xl/_rels/metadata.xml.rels` (metadata → richData; excerpt):

* `http://schemas.microsoft.com/office/2017/relationships/richValue`
* `http://schemas.microsoft.com/office/2017/relationships/richValueRel`
* `http://schemas.microsoft.com/office/2017/relationships/richValueTypes`
* `http://schemas.microsoft.com/office/2017/relationships/richValueStructure`

`[Content_Types].xml` in this fixture includes explicit overrides for these parts, e.g.:

* `/xl/cellimages.xml` → `application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml`
* `/xl/metadata.xml` → `application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml`
* `/xl/richData/richValueTypes.xml` → `application/vnd.ms-excel.richvaluetypes+xml`
* `/xl/richData/richValueStructure.xml` → `application/vnd.ms-excel.richvaluestructure+xml`

### Fixture: `fixtures/xlsx/basic/image-in-cell.xlsx` (`rdrichvalue.xml` + `richValueRel.xml` 2022 variant)

This is a modern Excel file demonstrating images stored as cell values via the
`metadata.xml` + `rdRichValue*` tables. It contains **two** images.

See also: [`fixtures/xlsx/basic/image-in-cell.md`](../fixtures/xlsx/basic/image-in-cell.md) (walkthrough of this fixture).

**Parts present** (complete inventory from `unzip -l`):

```
[Content_Types].xml
_rels/.rels
docProps/core.xml
docProps/app.xml
xl/workbook.xml
xl/_rels/workbook.xml.rels
xl/theme/theme1.xml
xl/styles.xml
xl/sharedStrings.xml
xl/worksheets/sheet1.xml
xl/worksheets/_rels/sheet1.xml.rels
xl/metadata.xml
xl/richData/richValueRel.xml
xl/richData/rdrichvalue.xml
xl/richData/rdrichvaluestructure.xml
xl/richData/rdRichValueTypes.xml
xl/richData/_rels/richValueRel.xml.rels
xl/media/image1.png
xl/media/image2.png
xl/printerSettings/printerSettings1.bin
```

**Sheet cell metadata (`vm`)**

From `xl/worksheets/sheet1.xml` (excerpt):

```xml
<c r="B2" t="e" vm="1"><v>#VALUE!</v></c>
<c r="B3" t="e" vm="1"><v>#VALUE!</v></c>
<c r="B4" t="e" vm="2"><v>#VALUE!</v></c>
```

Notes:

 * In this fixture, `vm` is **1-based** (`vm="1"` selects the first `<valueMetadata><bk>`).

**`xl/metadata.xml` (`futureMetadata name="XLRICHVALUE"` + `xlrd:rvb i="..."`)**

Exact shape (note the `xlrd` namespace and the `<futureMetadata name="XLRICHVALUE">` table):

```xml
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/>
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

Notes:

* `rc/@t` selects `metadataType name="XLRICHVALUE"`.
* `rc/@v` is a **0-based index into** the `<futureMetadata name="XLRICHVALUE">` `<bk>` list.
* `<xlrd:rvb i="N"/>` provides the **0-based rich value index** (`N`) into `xl/richData/rdrichvalue.xml`.

**`xl/richData/richValueRel.xml` (2022 namespace/root variant)**

In this fixture the relationship-slot table uses a **different root local-name and namespace**:

```xml
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
  <rel r:id="rId2"/>
</richValueRels>
```

Notes:

* Root local-name is `richValueRels` (plural).
* Namespace is **`…/2022/richvaluerel`**.
* Relationship slots are still positional (0-based): slot `0` selects the first `<rel/>`, etc.

**`xl/richData/rdrichvalue.xml` (positional `<v>` fields)**

```xml
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="2">
  <rv s="0"><v>0</v><v>5</v></rv>
  <rv s="0"><v>1</v><v>5</v></rv>
</rvData>
```

Notes:

* Namespace is **`…/2017/richdata`** (same as `richValue.xml` in the other fixture).
* `<rv>` records have **multiple `<v>` fields**; their meaning is **positional** and defined by
  `rdrichvaluestructure.xml`.

**`xl/richData/rdrichvaluestructure.xml` (key ordering defines `<v>` meanings)**

Exact structure definition for `_localImage` (ordering matters):

```xml
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <s t="_localImage">
    <k n="_rvRel:LocalImageIdentifier" t="i"/>
    <k n="CalcOrigin" t="i"/>
  </s>
</rvStructures>
```

Notes:

* The **ordered `<k>` list** defines how to interpret the ordered `<v>` list in `rdrichvalue.xml`.
* The relationship-slot index is stored in the field named
  **`_rvRel:LocalImageIdentifier`** (type `t="i"`), which points to a slot in
  `xl/richData/richValueRel.xml`. Do **not** hardcode "first `<v>`" — use the structure’s key order.
  * In this specific fixture, the `_localImage` key order is:
    1) `_rvRel:LocalImageIdentifier`
    2) `CalcOrigin`
    so `<rv><v>0</v><v>5</v></rv>` means: relationship slot `0`, `CalcOrigin = 5`.
* `xl/richData/_rels/richValueRel.xml.rels` is a **map** from `rId*` to a `Target`; its internal ordering
  is not meaningful. In `fixtures/xlsx/basic/image-in-cell.xlsx` the `.rels` file lists `rId2` before
  `rId1`, while `richValueRel.xml` lists `<rel r:id="rId1"/>` then `<rel r:id="rId2"/>`. Always resolve
  the slot index using the `<rel>` list order, then resolve `r:id` via the `.rels` file.

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

**`xl/richData/rdRichValueTypes.xml` root + namespace**

This fixture also includes `xl/richData/rdRichValueTypes.xml`, which uses the `richdata2` namespace and a
different root local-name:

```xml
<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
             xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
             mc:Ignorable="x"
             xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <global>
    <keyFlags>…</keyFlags>
  </global>
</rvTypesInfo>
```

This part is not required to follow the local-image resolution chain in the observed fixtures (which only
needs `metadata.xml` → `rdrichvalue.xml`/structure → `richValueRel.xml`/`.rels`), but should be preserved
byte-for-byte for round-trip safety.

**Workbook relationships**

From `xl/_rels/workbook.xml.rels` (excerpt):

* `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"`
  → `Target="metadata.xml"`
* `Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue"`
  → `Target="richData/rdrichvalue.xml"`
* `Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure"`
  → `Target="richData/rdrichvaluestructure.xml"`
* `Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes"`
  → `Target="richData/rdRichValueTypes.xml"`
* `Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel"`
  → `Target="richData/richValueRel.xml"`

`[Content_Types].xml` in this fixture includes explicit overrides for these parts, e.g.:

* `/xl/metadata.xml` → `application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml`
* `/xl/richData/richValueRel.xml` → `application/vnd.ms-excel.richvaluerel+xml`
* `/xl/richData/rdrichvalue.xml` → `application/vnd.ms-excel.rdrichvalue+xml`
* `/xl/richData/rdrichvaluestructure.xml` → `application/vnd.ms-excel.rdrichvaluestructure+xml`
* `/xl/richData/rdRichValueTypes.xml` → `application/vnd.ms-excel.rdrichvaluetypes+xml`

### Observed “rdRichValue*” naming (Excel + rust_xlsxwriter)

In addition to the real Excel fixture (`fixtures/xlsx/basic/image-in-cell.xlsx`) documented above, this repo
also contains a test that generates a “Place in Cell” workbook using `rust_xlsxwriter` and asserts the
presence of `rdRichValue*` parts:

* `crates/formula-xlsx/tests/embedded_images_place_in_cell_roundtrip.rs`

The generated workbook uses the same `rdRichValue*` naming convention for the rich value store:

* `xl/richData/rdrichvalue.xml`
* `xl/richData/rdrichvaluestructure.xml`
* `xl/richData/rdRichValueTypes.xml` (note casing)
* `xl/richData/richValueRel.xml` + `xl/richData/_rels/richValueRel.xml.rels`

And the workbook relationships include versioned Microsoft relationship types (also observed in
`fixtures/xlsx/basic/image-in-cell.xlsx`; full relationship/content-type details are documented in
[`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md)):

* `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue` (rdRichValue tables)
* `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure` (rdRichValue tables)
* `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes` (rdRichValue tables)
* `http://schemas.microsoft.com/office/2022/10/relationships/richValueRel` (richValueRel table)

Treat these as equivalent to the `richValue*` tables for the purposes of “images in cell” round-trip.

### Roles (high level)

| Part | Purpose |
|------|---------|
| `xl/richData/richValueTypes.xml` | Defines **type identifiers** (often numeric IDs) and links each type to a **structure ID** (string) that describes its field layout. |
| `xl/richData/richValueStructure.xml` | Defines **structures**: ordered field/member layouts keyed by **string IDs**. |
| `xl/richData/richValue.xml` | Stores the **rich value instances** (“objects”) in a workbook-global table. Each instance references a type (and/or structure) and stores member values. |
| `xl/richData/richValueRel.xml` | Stores a **relationship-ID table** (`r:id` strings) that can be referenced **by index** from rich values, avoiding embedding raw `rId*` strings inside each rich value payload. |
| `xl/richData/_rels/richValueRel.xml.rels` | OPC relationships for the `r:id` entries in `richValueRel.xml` (e.g. to `../media/imageN.png`). |
| `xl/richData/rdrichvalue.xml` | **rdRichValue variant** of the rich value instance table. Instances store positional `<v>` fields and use `rv/@s` to select a structure definition. |
| `xl/richData/rdrichvaluestructure.xml` | **rdRichValue variant** structure table. Defines ordered `<k>` keys; key ordering determines how to interpret positional `<v>` fields. |
| `xl/richData/rdRichValueTypes.xml` | **rdRichValue variant** type/key metadata (e.g. `<keyFlags>`). Not required to resolve local images in the observed fixtures, but should be preserved for round-trip safety. |

---

## How Excel wires cells to rich values (context)

The RichData parts above are workbook-global tables. A worksheet cell does **not** point directly at the
rich value store (`xl/richData/richValue*.xml` or `xl/richData/rdrichvalue.xml`).

Instead, Excel uses **cell metadata** in `xl/metadata.xml` (schema varies across Excel builds):

1. Worksheet cells use `c/@vm` (value-metadata index).
2. `vm` selects a `<valueMetadata><bk>` record in `xl/metadata.xml`.
3. That `<bk>` contains an `<rc t="…" v="…"/>` record.
4. Depending on the `metadata.xml` schema, `rc/@v` can mean different things:
   * In this repo’s images-in-cell fixtures, `rc/@v` is an index into an extension table (a
     `futureMetadata name="XLRICHVALUE"` table containing `xlrd:rvb i="…"` entries, where `rvb/@i` is the rich
     value index).
   * Other schemas may omit `futureMetadata`/`rvb`; in those cases `rc/@v` may need alternate interpretation
     (for example, it may directly be the rich value index). This direct mapping is **not** currently
     observed in the images-in-cell fixtures checked into this repo.

Minimal representative shape for the `futureMetadata`/`rvb` variant (index bases are important; see below):

```xml
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes>
    <!-- `t` in <rc> selects an entry in this list (often 1-based in Excel; 0-based is also observed). -->
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>

  <futureMetadata name="XLRICHVALUE">
    <!-- `v` in <rc> is typically a 0-based index into this bk list -->
    <bk>
      <extLst>
        <ext uri="{...}">
          <!-- `i` is the 0-based index into the rich value store part (often xl/richData/richValue*.xml) -->
          <xlrd:rvb i="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>

  <valueMetadata>
    <!-- vm selects a <bk> record (often 1-based; sometimes 0-based) -->
    <bk><rc t="1" v="0"/></bk>
  </valueMetadata>
</metadata>
```

All rich-value fixtures currently checked into this repo use the `futureMetadata`/`rvb` indirection (e.g.
`fixtures/xlsx/basic/image-in-cell.xlsx`, `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`,
`fixtures/xlsx/rich-data/richdata-minimal.xlsx`, `fixtures/xlsx/metadata/rich-values-vm.xlsx` (synthetic)). We have not
yet checked in a fixture where `rc/@v` directly equals the rich value index without a `futureMetadata`
table.

This indirection is important for engineering because:

* `vm` indexes are **independent** from `richValue.xml` indexes.
* `vm` base is **not consistent** across all observed files; treat `vm` as opaque and resolve best-effort
  (see [Index bases](#index-bases--indirection)).

---

## Index bases & indirection

Excel uses multiple indices; mixing bases is a common source of bugs.

### `vm` (worksheet cell attribute) — **0-based or 1-based (tolerate both)**

In worksheet XML, cells can carry `vm="n"` to attach value metadata:

```xml
<c r="B2" t="str" vm="1">
  <v>…</v>
</c>
```

Some workbooks use `vm="0"` for the first entry (0-based). Example from
`fixtures/xlsx/basic/image-in-cell-richdata.xlsx`:

```xml
<c r="A1" vm="0"><v>0</v></c>
```

Current Formula behavior:

* `vm` is treated as **ambiguous** (0-based or 1-based), because both appear in the wild.
  - Example: `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` uses `vm="0"` for the first record.
  - Example: `fixtures/xlsx/basic/image-in-cell.xlsx` uses `vm="1"` / `vm="2"`.
  Implementations should attempt to resolve both (e.g. try `vm` and `vm-1`), and preserve the original
  values when round-tripping.
* Missing `vm` means “no value metadata”.
* Preserve unusual values like `vm="0"` if encountered (even if they don’t resolve cleanly).

### Indices inside `xl/metadata.xml` used by `XLRICHVALUE`

| Index | Location | Base | Meaning |
|------:|----------|------|---------|
| `t` | `<valueMetadata><bk><rc t="…">` | usually 1-based (0-based also observed) | index into `<metadataTypes>` (selects `metadataType name="XLRICHVALUE"`) |
| `v` | `<valueMetadata><bk><rc v="…">` | usually 0-based | often an index into `<futureMetadata name="XLRICHVALUE"><bk>` (if present); other schemas may use `v` differently (including directly referencing the rich value index). |
| `i` | `<xlrd:rvb i="…"/>` | 0-based | rich value index into the rich value instance table (e.g. `xl/richData/richValue*.xml` or `xl/richData/rdrichvalue.xml`, depending on naming scheme) |

Notes:

* The `metadata.xml` schema varies across Excel builds. In this repo’s images-in-cell fixtures, rich value
  binding uses the `futureMetadata` + `xlrd:rvb` indirection. Other schemas may exist in the wild; if a
  workbook lacks `futureMetadata`/`rvb`, `rc/@v` may need alternate interpretation (e.g. directly as a rich
  value index, or via other extension tables). Preserve unknown metadata and implement mapping best-effort.

### Rich value instance table (`richValue*.xml` / `rdrichvalue.xml`) — **0-based**

Rich values are stored in a list; the rich value index is **0-based** and is referenced from `xl/metadata.xml`
in this repo’s fixtures via `xlrd:rvb/@i` (other mapping schemas may exist in the wild).

### `richValueRel.xml` relationship table — **0-based**

Relationship references used inside rich values are **integers indexing into `richValueRel.xml`**, starting at `0`.

### Why `richValueRel.xml` exists (avoid embedding `rId*`)

OPC relationship IDs (`rId1`, `rId2`, …) are:

* **local to the `.rels` file**
* not semantically meaningful
* often renumbered by writers

Excel avoids storing raw strings like `rId17` inside every rich value instance. Instead:

1. The rich value instance (`richValue.xml` or `rdrichvalue.xml`) stores a **relationship index** (e.g. `rel=0`).
2. That index selects an entry in `richValueRel.xml` (e.g. entry `0` is `r:id="rId5"`).
3. `rId5` is resolved using `xl/richData/_rels/richValueRel.xml.rels` to find the actual `Target` (e.g. `../media/image1.png`).

This design allows relationship IDs to change without rewriting every rich value payload.

---

## End-to-end reference chain (example)

The exact XML vocab inside `richValue.xml` varies across Excel builds, but the *indexing chain* for images-in-cell
is generally:

In this repo’s fixture corpus, mapping `vm`/`metadata.xml` → rich value indices uses the
`futureMetadata`/`xlrd:rvb` indirection described below. Other schemas may exist in the wild; preserve and
round-trip unknown metadata byte-for-byte.

### Observed mapping: `futureMetadata` / `rvb` indirection

1. **Worksheet cell** (`xl/worksheets/sheetN.xml`)
   - Cell has `c/@vm="0"` or `c/@vm="1"` (value metadata index; **0-based or 1-based** in observed files).
2. **Value metadata** (`xl/metadata.xml`)
   - `vm` selects a `<valueMetadata><bk>` record (base varies; preserve and resolve best-effort).
   - That `<bk>` contains `<rc t="…" v="0"/>` where `v` is **0-based** into `futureMetadata name="XLRICHVALUE"`.
3. **Future metadata** (`xl/metadata.xml`)
    - `futureMetadata name="XLRICHVALUE"` `<bk>` #0 contains `<xlrd:rvb i="5"/>`.
    - `i=5` is the **0-based rich value index** into the rich value instance table:
      - `xl/richData/richValue*.xml` (Excel-like naming), or
      - `xl/richData/rdrichvalue.xml` (rdRichValue naming)
4. **Rich value** (rich value instance table)
    - Rich value record #5 is an “image” typed rich value (exact representation varies by Excel build).
   - Its payload contains a **relationship index** (e.g. `relIndex = 0`, **0-based**) into `richValueRel.xml`.
5. **Relationship table** (`xl/richData/richValueRel.xml`)
   - Relationship table entry #0 contains `r:id="rId7"`.
6. **OPC resolution** (`xl/richData/_rels/richValueRel.xml.rels`)
   - Relationship `Id="rId7"` resolves to an OPC `Target` (often a media part like `../media/image1.png`,
     but treat this as opaque; other targets/types may appear depending on Excel build).

So: **cell → vm (0/1-based) → metadata.xml → rvb@i (0-based) → rich value instance table → relIndex (0-based) →
richValueRel.xml → rId → .rels target → image bytes**.

## Minimal XML skeletons (best-effort)

These skeletons aim to show **roots, key child tags, and key attributes** as Excel tends to emit them. Namespaces and some attribute names may differ across builds—treat them as *shape guidance*, not a strict schema.

### 1) `xl/richData/richValueTypes.xml` (optional / feature-dependent)

Defines **type identifiers** and links to a **structure ID** (string).

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypes xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <!-- One entry per type used in this workbook. -->
  <!-- Type identifiers may be numeric IDs or strings depending on Excel build. -->
  <types>
    <type id="0" name="com.microsoft.excel.image" structure="s_image"/>
    <!-- ... -->
  </types>
</rvTypes>
```

Notes:

* `id` is the key: a dense integer domain is typical.
* `structure` is a string key into `richValueStructure.xml`.
* `name` is often present but should be treated as opaque.

### 2) `xl/richData/richValueStructure.xml` (optional / feature-dependent)

Defines member/field layouts keyed by **string IDs**. Structures are typically interpreted as “schemas” for the ordered value list stored in `richValue.xml`.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStruct xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <structures>
    <structure id="s_image">
      <!-- Member ordering matters; richValue.xml payloads are positional. -->
      <member name="imageRel" kind="rel"/>
      <member name="altText"  kind="string"/>
      <!-- ... -->
    </structure>
    <!-- ... -->
  </structures>
</rvStruct>
```

Notes:

 * The `kind="…"` attribute encodes the value representation. In the in-repo images-in-cell fixtures it is
   observed as `kind="rel"` for the image relationship field; other rich value types may use additional
   kinds.
* The **ordering** of `<member>` entries is significant: instances generally encode member values positionally.

### 3) `xl/richData/richValueRel.xml`

Stores a **vector/table** of `r:id` strings. The *index* into this vector is what rich values store.

Three root/namespace variants are observed in-repo:

1) `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` (**synthetic**; `richValueRel` root, `…/2017/richdata2`):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <!-- Table position = relationship index (0-based). -->
  <rel r:id="rId1"/>
  <!-- ... -->
</richValueRel>
```

2) `fixtures/xlsx/rich-data/images-in-cell.xlsx` (**real Excel**; `rvRel` root, `…/2017/richdata`, with a `<rels>` wrapper):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <!-- Table position = relationship index (0-based). -->
  <rels>
    <rel r:id="rId1"/>
  </rels>
</rvRel>
```

3) `fixtures/xlsx/basic/image-in-cell.xlsx` (**real Excel**; `richValueRels` root, `…/2022/richvaluerel`):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <!-- Table position = relationship index (0-based). -->
  <rel r:id="rId1"/>
  <rel r:id="rId2"/>
</richValueRels>
```

In all cases, treat the `<rel>` list as an **ordered table** (0-based indices). Some variants wrap the
entries (e.g. `<rels><rel .../></rels>`); match on element local-names and preserve unknown structure when
round-tripping.

And the corresponding OPC relationships part:

`xl/richData/_rels/richValueRel.xml.rels`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship
    Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image1.png"/>
  <!-- Other relationship types/targets may also occur (unverified); preserve unknown entries. -->
  <!-- ... -->
</Relationships>
```

### 4) `xl/richData/richValue.xml`

Stores the actual rich value instances. Each instance references a type (and/or structure) and encodes member values (often positionally, guided by `richValueStructure.xml`).

Two variants are observed in-repo:

1) `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` (**synthetic**; minimal `t="image"` payload referencing relationship slot `0`):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <!-- Rich value index is typically the 0-based order of <rv> records (unless an explicit id/index is provided). -->
  <rv s="0" t="image">
    <!-- Relationship index (0-based) into richValueRel.xml -->
    <v kind="rel">0</v>
  </rv>
</rvData>
```

2) `fixtures/xlsx/rich-data/images-in-cell.xlsx` (**real Excel**; typed payload referencing relationship slot `0`):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <!-- Rich value index 0 -->
    <rv type="0">
      <!-- Relationship index (0-based) into richValueRel.xml -->
      <v kind="rel">0</v>
    </rv>
  </values>
</rvData>
```

Other builds may:

* split values across `xl/richData/richValue1.xml`, `richValue2.xml`, ... (or the plural `richValues1.xml` / `richValues2.xml` naming variant)
  * Formula treats these as a single logical stream ordered by numeric suffix (see
    `crates/formula-xlsx/tests/rich_value_part_numeric_suffix_order.rs`).
* include an explicit global index attribute on `<rv>` (e.g. `i="…"`, `id="…"`, `idx="…"`)
* include multiple `<v>` members, with types indicated by attributes like `kind="rel"` and/or `t="rel"` / `t="r"` / etc.

Notes:

* The “rich value index” is 0-based. In this repo’s images-in-cell fixtures, the cell metadata reaches it
  indirectly via the `futureMetadata`/`xlrd:rvb` lookup table. Other mapping schemas may exist in the wild;
  treat the metadata tables as opaque and preserve them.
* The “relationship index” stored in the payload is 0-based and indexes into `richValueRel.xml`.

### 4b) `xl/richData/rdrichvalue.xml` + `xl/richData/rdrichvaluestructure.xml` (rdRichValue variant)

Observed in `fixtures/xlsx/basic/image-in-cell.xlsx`:

`rdrichvaluestructure.xml` assigns **names and types** to positional `<v>` fields (ordering matters):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <s t="_localImage">
    <k n="_rvRel:LocalImageIdentifier" t="i"/>
    <k n="CalcOrigin" t="i"/>
  </s>
</rvStructures>
```

`rdrichvalue.xml` then stores **positional values** matching that key order:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="2">
  <rv s="0"><v>0</v><v>5</v></rv>
  <rv s="0"><v>1</v><v>5</v></rv>
</rvData>
```

For images-in-cell in this variant, the relationship-slot index is carried by the key named
`_rvRel:LocalImageIdentifier` (type `i`), which points to an entry in `xl/richData/richValueRel.xml`. Do not
assume it is “the first `<v>`” — it is “the `<v>` corresponding to the `_rvRel:LocalImageIdentifier` key in
the structure’s `<k>` list”.

---

## OPC relationships and `[Content_Types].xml`

### Relationship graph (where Excel puts the links)

Excel uses OPC relationships to connect:

* `xl/workbook.xml` → `xl/metadata.xml` (worksheet cells only carry `vm`/`cm` indices; the actual mapping tables live in metadata).
* `xl/workbook.xml` → `xl/richData/*` (the rich value tables; often directly related from the workbook for in-cell images).
* (Sometimes) `xl/metadata.xml` → `xl/richData/*` via `xl/_rels/metadata.xml.rels`.
* `xl/richData/richValueRel.xml` → external OPC targets (commonly `xl/media/*` images) via `xl/richData/_rels/richValueRel.xml.rels`.

The workbook → metadata relationship uses a standard SpreadsheetML relationship type URI. Two variants are
observed in this repo:

* `http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata`
  * Observed in the synthetic fixture `fixtures/xlsx/metadata/rich-values-vm.xlsx`
  * Observed in the synthetic fixture `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`
  * Observed in `fixtures/xlsx/rich-data/images-in-cell.xlsx`
* `http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata`
  * Observed in `fixtures/xlsx/basic/image-in-cell.xlsx`

Additionally, `xl/workbook.xml` may include a `<metadata r:id="..."/>` element pointing at the relationship
ID for the metadata part (observed in the synthetic fixture `fixtures/xlsx/metadata/rich-values-vm.xlsx`). Some workbooks omit
this element and only include the relationship in `workbook.xml.rels` (observed in
`fixtures/xlsx/basic/image-in-cell-richdata.xlsx`). Preserve whichever representation the source workbook
uses.

The richValue relationships are Microsoft-specific. Observed in this repo:

* `http://schemas.microsoft.com/office/2017/06/relationships/richValue` → `xl/richData/richValue.xml`
* `http://schemas.microsoft.com/office/2017/06/relationships/richValueRel` → `xl/richData/richValueRel.xml`
   * Observed in `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`
* When the richData parts are related from `xl/metadata.xml` (via `xl/_rels/metadata.xml.rels`), Excel uses
  unversioned 2017 relationship Type URIs:
  * `http://schemas.microsoft.com/office/2017/relationships/richValue` → `xl/richData/richValue.xml`
  * `http://schemas.microsoft.com/office/2017/relationships/richValueRel` → `xl/richData/richValueRel.xml`
  * `http://schemas.microsoft.com/office/2017/relationships/richValueTypes` → `xl/richData/richValueTypes.xml`
  * `http://schemas.microsoft.com/office/2017/relationships/richValueStructure` → `xl/richData/richValueStructure.xml`
    * Observed in the real Excel fixture `fixtures/xlsx/rich-data/images-in-cell.xlsx` and the synthetic fixture
      `fixtures/xlsx/rich-data/richdata-minimal.xlsx`
* `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue` → `xl/richData/rdrichvalue.xml`
   * Observed in `fixtures/xlsx/basic/image-in-cell.xlsx`
* `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure` → `xl/richData/rdrichvaluestructure.xml`
   * Observed in `fixtures/xlsx/basic/image-in-cell.xlsx`
* `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes` → `xl/richData/rdRichValueTypes.xml`
  * Observed in `fixtures/xlsx/basic/image-in-cell.xlsx`
* `http://schemas.microsoft.com/office/2022/10/relationships/richValueRel` → `xl/richData/richValueRel.xml`
  * Observed in `fixtures/xlsx/basic/image-in-cell.xlsx`

Some workbooks may instead relate the richData parts from `xl/metadata.xml` via `xl/_rels/metadata.xml.rels`.
For parsing and round-trip safety, treat both relationship layouts as valid.

Implementation guidance:

* When parsing, do not hardcode exact Type URIs; match by resolved `Target` path when necessary and preserve unknown relationship types.
* When writing new files, keep relationship IDs stable and prefer “append-only” updates. Excel may rewrite
  relationship type URIs and renumber `rId*` values.

#### Observed values summary (from in-repo fixtures/tests)

These values are copied from fixtures/tests in this repo. Sources include both real Excel workbooks and
synthetic fixtures (tagged `Application=Formula Fixtures`). Values that are only observed in synthetic
fixtures should not be treated as “confirmed Excel output”; preserve unknown URIs/content types rather than
hardcoding assumptions.

| Kind | Value | Source |
|------|-------|--------|
| Workbook → metadata relationship Type | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata` | `fixtures/xlsx/metadata/rich-values-vm.xlsx`, `fixtures/xlsx/rich-data/images-in-cell.xlsx`, `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` |
| Workbook → metadata relationship Type | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata` | `fixtures/xlsx/basic/image-in-cell.xlsx` |
| Workbook → cellimages relationship Type | `http://schemas.microsoft.com/office/2019/relationships/cellimages` | `fixtures/xlsx/rich-data/images-in-cell.xlsx` |
| Workbook → richValue relationship Type | `http://schemas.microsoft.com/office/2017/06/relationships/richValue` | `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` |
| Workbook → richValueRel relationship Type | `http://schemas.microsoft.com/office/2017/06/relationships/richValueRel` | `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` |
| Metadata → richValue relationship Type (`xl/_rels/metadata.xml.rels`) | `http://schemas.microsoft.com/office/2017/relationships/richValue` | `fixtures/xlsx/rich-data/images-in-cell.xlsx`, `fixtures/xlsx/rich-data/richdata-minimal.xlsx` |
| Metadata → richValueRel relationship Type (`xl/_rels/metadata.xml.rels`) | `http://schemas.microsoft.com/office/2017/relationships/richValueRel` | `fixtures/xlsx/rich-data/images-in-cell.xlsx`, `fixtures/xlsx/rich-data/richdata-minimal.xlsx` |
| Metadata → richValueTypes relationship Type (`xl/_rels/metadata.xml.rels`) | `http://schemas.microsoft.com/office/2017/relationships/richValueTypes` | `fixtures/xlsx/rich-data/images-in-cell.xlsx`, `fixtures/xlsx/rich-data/richdata-minimal.xlsx` |
| Metadata → richValueStructure relationship Type (`xl/_rels/metadata.xml.rels`) | `http://schemas.microsoft.com/office/2017/relationships/richValueStructure` | `fixtures/xlsx/rich-data/images-in-cell.xlsx`, `fixtures/xlsx/rich-data/richdata-minimal.xlsx` |
| Workbook → rdRichValue relationship Type | `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue` | `fixtures/xlsx/basic/image-in-cell.xlsx` |
| Workbook → rdRichValueStructure relationship Type | `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure` | `fixtures/xlsx/basic/image-in-cell.xlsx` |
| Workbook → rdRichValueTypes relationship Type | `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes` | `fixtures/xlsx/basic/image-in-cell.xlsx` |
| Workbook → richValueRel relationship Type | `http://schemas.microsoft.com/office/2022/10/relationships/richValueRel` | `fixtures/xlsx/basic/image-in-cell.xlsx` |
| `richValueRel.xml` root + namespace | `<richValueRel>` / `http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2` | `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` |
| `richValueRel.xml` root + namespace | `<rvRel>` / `http://schemas.microsoft.com/office/spreadsheetml/2017/richdata` | `fixtures/xlsx/rich-data/images-in-cell.xlsx` |
| `richValueRel.xml` root + namespace | `<richValueRels>` / `http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel` | `fixtures/xlsx/basic/image-in-cell.xlsx` |
| `richValue.xml` namespace | `http://schemas.microsoft.com/office/spreadsheetml/2017/richdata` | `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`, `fixtures/xlsx/rich-data/images-in-cell.xlsx` |
| `rdrichvalue.xml` namespace | `http://schemas.microsoft.com/office/spreadsheetml/2017/richdata` | `fixtures/xlsx/basic/image-in-cell.xlsx` |
| `rdrichvaluestructure.xml` root + namespace | `<rvStructures>` / `http://schemas.microsoft.com/office/spreadsheetml/2017/richdata` | `fixtures/xlsx/basic/image-in-cell.xlsx` |
| `rdRichValueTypes.xml` root + namespace | `<rvTypesInfo>` / `http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2` | `fixtures/xlsx/basic/image-in-cell.xlsx` |
| `richValue.xml` content type override | `application/vnd.ms-excel.richvalue+xml` | `fixtures/xlsx/rich-data/images-in-cell.xlsx`, `fixtures/xlsx/rich-data/richdata-minimal.xlsx`, `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` |
| `richValueRel.xml` content type override | `application/vnd.ms-excel.richvaluerel+xml` | `fixtures/xlsx/rich-data/images-in-cell.xlsx`, `fixtures/xlsx/rich-data/richdata-minimal.xlsx`, `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`, `fixtures/xlsx/basic/image-in-cell.xlsx` |
| `richValueTypes.xml` content type override | `application/vnd.ms-excel.richvaluetypes+xml` | `fixtures/xlsx/rich-data/images-in-cell.xlsx`, `fixtures/xlsx/rich-data/richdata-minimal.xlsx` |
| `richValueStructure.xml` content type override | `application/vnd.ms-excel.richvaluestructure+xml` | `fixtures/xlsx/rich-data/images-in-cell.xlsx`, `fixtures/xlsx/rich-data/richdata-minimal.xlsx` |
| `cellimages.xml` content type override | `application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml` | `fixtures/xlsx/rich-data/images-in-cell.xlsx` |
| `metadata.xml` content type override | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml` | `fixtures/xlsx/metadata/rich-values-vm.xlsx`, `fixtures/xlsx/rich-data/images-in-cell.xlsx`, `fixtures/xlsx/rich-data/richdata-minimal.xlsx`, `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`, `fixtures/xlsx/basic/image-in-cell.xlsx` |
| `metadata.xml` content type override | `application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml` | `crates/formula-xlsx/tests/metadata_rich_value_roundtrip.rs` |

#### Minimal `.rels` skeletons (best-effort)

`xl/_rels/workbook.xml.rels` (workbook → metadata and/or richData):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <!-- ...other workbook relationships... -->
  <!-- metadata.xml (Type may be /metadata or /sheetMetadata) -->
  <Relationship Id="rIdMeta"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"
                Target="metadata.xml"/>

  <!-- richData tables (Type URIs are Microsoft-specific; examples observed in fixtures) -->
  <Relationship Id="rIdRichValue"
                Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue"
                Target="richData/richValue.xml"/>
  <Relationship Id="rIdRichValueRel"
                Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel"
                Target="richData/richValueRel.xml"/>
</Relationships>
```

`xl/_rels/metadata.xml.rels` (optional; metadata → richData tables; observed in the real Excel fixture
`fixtures/xlsx/rich-data/images-in-cell.xlsx` and the synthetic regression fixture
`fixtures/xlsx/rich-data/richdata-minimal.xlsx`):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdRichValueTypes"
                Type="http://schemas.microsoft.com/office/2017/relationships/richValueTypes"
                Target="richData/richValueTypes.xml"/>
  <Relationship Id="rIdRichValueStructure"
                Type="http://schemas.microsoft.com/office/2017/relationships/richValueStructure"
                Target="richData/richValueStructure.xml"/>
  <Relationship Id="rIdRichValueRel"
                Type="http://schemas.microsoft.com/office/2017/relationships/richValueRel"
                Target="richData/richValueRel.xml"/>
  <Relationship Id="rIdRichValue"
                Type="http://schemas.microsoft.com/office/2017/relationships/richValue"
                Target="richData/richValue.xml"/>
</Relationships>
```

### `[Content_Types].xml` considerations

In this repo’s fixture corpus, workbooks that include `xl/metadata.xml` and/or `xl/richData/*` also include
explicit `[Content_Types].xml` `<Override>` entries for those parts. Observed patterns in this repo:

* `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` includes explicit overrides for `xl/metadata.xml` and
  the minimal rich value parts (`xl/richData/richValue.xml`, `xl/richData/richValueRel.xml`).
* `fixtures/xlsx/basic/image-in-cell.xlsx` includes explicit `<Override>` entries for `xl/metadata.xml` and
  `xl/richData/*` (including `rdrichvalue.xml`, `rdrichvaluestructure.xml`, `rdRichValueTypes.xml`,
  `richValueRel.xml`).
* `fixtures/xlsx/rich-data/images-in-cell.xlsx` includes explicit `<Override>` entries for:
  * `xl/cellimages.xml`
  * `xl/metadata.xml`
  * `xl/richData/richValue*.xml` (including `richValueTypes.xml` / `richValueStructure.xml`)
* The synthetic fixture `fixtures/xlsx/metadata/rich-values-vm.xlsx` includes an override:
  * `<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>`
* Some tests construct workbooks that use:
  * `application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml` for `/xl/metadata.xml`

For the richData tables themselves, Excel emits Microsoft-specific content types in the real Excel fixture
`fixtures/xlsx/rich-data/images-in-cell.xlsx` and the synthetic regression fixture
`fixtures/xlsx/rich-data/richdata-minimal.xlsx` (other producers may vary; preserve whatever is present):

```xml
<Override PartName="/xl/richData/richValue.xml"          ContentType="application/vnd.ms-excel.richvalue+xml"/>
<Override PartName="/xl/richData/richValueRel.xml"       ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
<Override PartName="/xl/richData/richValueTypes.xml"     ContentType="application/vnd.ms-excel.richvaluetypes+xml"/>
<Override PartName="/xl/richData/richValueStructure.xml" ContentType="application/vnd.ms-excel.richvaluestructure+xml"/>
```

Observed in `fixtures/xlsx/basic/image-in-cell.xlsx` (rdRichValue variant):

```xml
<Override PartName="/xl/richData/richValueRel.xml"         ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
<Override PartName="/xl/richData/rdrichvalue.xml"          ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>
<Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>
<Override PartName="/xl/richData/rdRichValueTypes.xml"     ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>
```

Implementation guidance:

* When round-tripping an existing file: preserve the original overrides verbatim.
* When generating from scratch: emitting overrides for non-standard parts can improve compatibility, but
  preserve and round-trip whatever the source workbook uses.

---

## Practical parsing strategy (recommended)

For images-in-cell, the minimum viable read path usually looks like:

1. Locate `xl/metadata.xml` (typically via `xl/_rels/workbook.xml.rels`, but fall back to part existence).
2. Locate the `xl/richData/*` parts (often directly via `xl/_rels/workbook.xml.rels`; sometimes via
   `xl/_rels/metadata.xml.rels`; also fall back to part existence).
3. If present, parse `richValueTypes.xml` into `type_id -> structure_id`.
4. If present, parse `richValueStructure.xml` into `structure_id -> ordered member schema`.
   * For the `rdRichValue*` naming variant (observed in `fixtures/xlsx/basic/image-in-cell.xlsx`), the
     analogous structure table is `xl/richData/rdrichvaluestructure.xml` (which defines ordered `<k>` keys
     for positional `<v>` fields).
5. Parse `richValueRel.xml` into `rel_index -> rId`.
6. Parse `xl/richData/_rels/richValueRel.xml.rels` into `rId -> target` (image path).
7. Parse the rich value store (`richValue*.xml` and/or `rdrichvalue.xml`) into a table of
   `rich_value_index -> {type_id, payload...}`.
   * For the `rdrichvalue.xml` variant, interpret positional `<v>` fields using
     `rdrichvaluestructure.xml` key order. For local embedded images the relationship-slot index is stored
     in the field named `_rvRel:LocalImageIdentifier` (not necessarily “first `<v>`”).
8. Parse `xl/metadata.xml` to resolve `vm` (cell attribute) → `rich_value_index` (best-effort; schemas vary:
   some use `xlrd:rvb/@i`, others may reference the rich value index directly).
9. Use worksheet cell `vm` values to map cells → `rich_value_index`.

For writing, the safest approach is typically “append-only”:

* Append new relationships to `richValueRel.xml` + `.rels`
* Append new rich values to `richValue.xml`
* Avoid renumbering existing indices unless you fully rebuild all referencing metadata

---

## Where this is implemented in Formula (code pointers)

* Rich-value image extraction helper (follows the chain `vm` → rich value → relationship slot → media):
  * [`crates/formula-xlsx/src/rich_data/mod.rs`](../crates/formula-xlsx/src/rich_data/mod.rs)
    (`extract_rich_cell_images`)
* `vm`/`metadata.xml` parsing used to populate `XlsxDocument::rich_value_index()`:
  * [`crates/formula-xlsx/src/read/mod.rs`](../crates/formula-xlsx/src/read/mod.rs)
    (`MetadataPart`)
