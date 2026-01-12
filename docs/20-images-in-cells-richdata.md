# Excel RichData (`richValue*`) parts for Images-in-Cell (`IMAGE()` / “Place in Cell”)

Excel’s “Images in Cell” feature (insert picture → **Place in Cell**, and the `IMAGE()` function) is backed by a **RichData / RichValue** subsystem. Rather than embedding image references directly in worksheet cell XML, Excel stores *typed rich value instances* in workbook-level parts under `xl/richData/`, then attaches cells to those instances via metadata.

This note documents the **expected part set**, the **role of each part**, and the **minimal XML shapes** needed to parse/write Excel-generated files.

For the overall “images in cells” packaging overview (including the optional `xl/cellImages.xml` store part (a.k.a. `xl/cellimages.xml`), `xl/metadata.xml`,
and current Formula status/tests), see: [20-images-in-cells.md](./20-images-in-cells.md).

For a **concrete, confirmed** “Place in Cell” (embedded local image) package shape generated via `rust_xlsxwriter`
(including the exact `rdrichvalue*` structure keys like `_rvRel:LocalImageIdentifier` and the `CalcOrigin` field), see:

- [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md)

> Status: best-effort reverse engineering. Exact namespaces / relationship-type URIs may vary by Excel version; preserve unknown attributes and namespaces when round-tripping.

---

## Expected part set (workbook-level)

When a workbook contains at least one RichData value (including images-in-cell), Excel typically adds:

```
xl/
  richData/
    richValue.xml
    richValueRel.xml
    richValueTypes.xml        # optional (not present in all workbooks)
    richValueStructure.xml    # optional (not present in all workbooks)
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
* File naming varies across producers:
  * “Excel-like” naming: `richValue.xml`, `richValueTypes.xml`, `richValueStructure.xml`
  * “rdRichValue” naming (observed from `rust_xlsxwriter` output in this repo):
    * `rdrichvalue.xml`
    * `rdrichvaluestructure.xml`
    * `rdRichValueTypes.xml` (note casing)
  For robust parsing, prefer relationship discovery + local-name matching rather than hardcoding a single
  filename spelling/casing.

## Observed “image in cell” fixture (in-repo)

The repository includes `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`, a minimal workbook that contains
an image-in-cell backed by RichData. Key observations (useful for implementers):

* `xl/worksheets/sheet1.xml` contains a cell with `vm="0"`:

  ```xml
  <c r="A1" vm="0"><v>0</v></c>
  ```

* `xl/metadata.xml` contains `<metadataTypes>` and `<valueMetadata>`, but no `futureMetadata` / `rvb`:

  ```xml
  <valueMetadata count="1">
    <bk><rc t="1" v="0"/></bk>
  </valueMetadata>
  ```

* `xl/richData/richValue.xml` stores an image rich value whose payload is a relationship-table index:

  ```xml
  <rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
    <rv s="0" t="image"><v>0</v></rv>
  </rvData>
  ```

* `xl/richData/richValueRel.xml` is a bare `<rel>` list (no `<rels>` wrapper), and uses the `richdata2` namespace:

  ```xml
  <richValueRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <rel r:id="rId1"/>
  </richValueRel>
  ```

* Workbook relationships (`xl/_rels/workbook.xml.rels`) link directly to the rich value parts using Microsoft
  relationship types:
  * `http://schemas.microsoft.com/office/2017/06/relationships/richValue` → `richData/richValue.xml`
  * `http://schemas.microsoft.com/office/2017/06/relationships/richValueRel` → `richData/richValueRel.xml`
* The relationship from the workbook to `xl/metadata.xml` uses `Type="…/sheetMetadata"` (not `…/metadata`) in
  this file.
* `[Content_Types].xml` does **not** include overrides for `xl/metadata.xml` or `xl/richData/*` in this fixture;
  it relies on the default `application/xml`. Preserve whatever the source workbook uses.

### Observed “rdRichValue*” naming (rust_xlsxwriter-generated)

This repo also contains a test that generates a “Place in Cell” workbook using `rust_xlsxwriter` and asserts
the presence of RichData parts:

* `crates/formula-xlsx/tests/embedded_images_place_in_cell_roundtrip.rs`

The generated workbook uses a different naming convention for the rich value store:

* `xl/richData/rdrichvalue.xml`
* `xl/richData/rdrichvaluestructure.xml`
* `xl/richData/rdRichValueTypes.xml` (note casing)
* `xl/richData/richValueRel.xml` + `xl/richData/_rels/richValueRel.xml.rels`

And the workbook relationships include versioned Microsoft relationship types (partial list; some asserted in the test, full set documented in [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md)):

* `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue` (rdRichValue tables)
* `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure` (rdRichValue tables)
* `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes` (rdRichValue tables)
* `http://schemas.microsoft.com/office/2022/10/relationships/richValueRel` (richValueRel table)

Treat these as equivalent to the `richValue*` tables for the purposes of “images in cell” round-trip.

Concrete schema details for the rust_xlsxwriter “Place in Cell” file (including the exact worksheet cell
encoding `t="e"`/`#VALUE!`, the `_localImage` rich value structure keys, `CalcOrigin` ordering/values, and
the exact relationship/content-type URIs) are documented here:

* [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md)

### Roles (high level)

| Part | Purpose |
|------|---------|
| `xl/richData/richValueTypes.xml` | Defines **type identifiers** (often numeric IDs) and links each type to a **structure ID** (string) that describes its field layout. |
| `xl/richData/richValueStructure.xml` | Defines **structures**: ordered field/member layouts keyed by **string IDs**. |
| `xl/richData/richValue.xml` | Stores the **rich value instances** (“objects”) in a workbook-global table. Each instance references a type (and/or structure) and stores member values. |
| `xl/richData/richValueRel.xml` | Stores a **relationship-ID table** (`r:id` strings) that can be referenced **by index** from rich values, avoiding embedding raw `rId*` strings inside each rich value payload. |
| `xl/richData/_rels/richValueRel.xml.rels` | OPC relationships for the `r:id` entries in `richValueRel.xml` (e.g. to `../media/imageN.png`). |

---

## How Excel wires cells to rich values (context)

The RichData parts above are workbook-global tables. A worksheet cell does **not** point directly at `xl/richData/richValue.xml`.

Instead, Excel uses **cell metadata** in `xl/metadata.xml` (schema varies across Excel builds):

1. Worksheet cells use `c/@vm` (value-metadata index).
2. `vm` selects a `<valueMetadata><bk>` record in `xl/metadata.xml`.
3. That `<bk>` contains an `<rc t="…" v="…"/>` record.
4. Depending on the `metadata.xml` shape, `rc/@v` may:
   * directly be the **0-based rich value index** into `xl/richData/richValue.xml`, or
   * be an index into another extension table (commonly a `futureMetadata name="XLRICHVALUE"` table containing
     `xlrd:rvb i="…"` entries, where `rvb/@i` is the rich value index).

Minimal representative shape for the `futureMetadata`/`rvb` variant (index bases are important; see below):

```xml
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes>
    <!-- `t` in <rc> is a 1-based index into this list -->
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>

  <futureMetadata name="XLRICHVALUE">
    <!-- `v` in <rc> is a 0-based index into this bk list -->
    <bk>
      <extLst>
        <ext uri="{...}">
          <!-- `i` is the 0-based index into xl/richData/richValue.xml -->
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

Observed minimal shape without `futureMetadata`/`rvb` (where `rc/@v` appears to reference the rich value
index directly), from `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`:

```xml
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="0" copy="1" pasteAll="1" pasteValues="1"/>
  </metadataTypes>
  <valueMetadata count="1">
    <bk>
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
```

This indirection is important for engineering because:

* `vm` indexes are **independent** from `richValue.xml` indexes.
* `vm` base is **not consistent** across all observed files; treat `vm` as opaque and resolve best-effort
  (see [Index bases](#index-bases--indirection)).

---

## Index bases & indirection

Excel uses multiple indices; mixing bases is a common source of bugs.

### `vm` (worksheet cell attribute) — **0-based vs 1-based**

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

* Excel emits both **0-based** and **1-based** `vm` values in different files/contexts.
  - Example: `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` uses `vm="0"`.
  - Example: `fixtures/xlsx/metadata/rich-values-vm.xlsx` uses `vm="1"`.
* Formula treats `vm` as **ambiguous** (0-based or 1-based) and tries to resolve both bases where possible
  (see `crates/formula-xlsx/src/rich_data/mod.rs`).
* Missing `vm` means “no value metadata”.
* Preserve `vm` exactly on round-trip even if it doesn’t resolve cleanly.

### Indices inside `xl/metadata.xml` used by `XLRICHVALUE`

| Index | Location | Base | Meaning |
|------:|----------|------|---------|
| `t` | `<valueMetadata><bk><rc t="…">` | 1-based | index into `<metadataTypes>` (selects `metadataType name="XLRICHVALUE"`) |
| `v` | `<valueMetadata><bk><rc v="…">` | 0-based | often an index into `<futureMetadata name="XLRICHVALUE"><bk>` (if present); other schemas may use `v` differently (including directly referencing the rich value index). |
| `i` | `<xlrd:rvb i="…"/>` | 0-based | rich value index into `xl/richData/richValue.xml` |

Notes:

* The `metadata.xml` schema varies across Excel builds. Some files do not include `futureMetadata`/`rvb`;
  in those, `rc/@v` may directly refer to the rich value index (or be interpreted via other extension
  tables). Preserve unknown metadata and implement mapping best-effort.

### `richValue.xml` rich value table — **0-based**

Rich values are stored in a list; the rich value index is **0-based** and is referenced from `xl/metadata.xml`
either directly (e.g. `rc/@v = richValueIndex`) or indirectly (e.g. via `xlrd:rvb/@i`).

### `richValueRel.xml` relationship table — **0-based**

Relationship references used inside rich values are **integers indexing into `richValueRel.xml`**, starting at `0`.

### Why `richValueRel.xml` exists (avoid embedding `rId*`)

OPC relationship IDs (`rId1`, `rId2`, …) are:

* **local to the `.rels` file**
* not semantically meaningful
* often renumbered by writers

Excel avoids storing raw strings like `rId17` inside every rich value instance. Instead:

1. `richValue.xml` stores a **relationship index** (e.g. `rel=0`).
2. That index selects an entry in `richValueRel.xml` (e.g. entry `0` is `r:id="rId5"`).
3. `rId5` is resolved using `xl/richData/_rels/richValueRel.xml.rels` to find the actual `Target` (e.g. `../media/image1.png`).

This design allows relationship IDs to change without rewriting every rich value payload.

---

## End-to-end reference chain (example)

The exact XML vocab inside `richValue.xml` varies across Excel builds, but the *indexing chain* for images-in-cell
is generally:

There are (at least) two observed variants for mapping `vm`/`metadata.xml` → rich value indices.

### Variant A: `futureMetadata` / `rvb` indirection

1. **Worksheet cell** (`xl/worksheets/sheetN.xml`)
   - Cell has `c/@vm="0"` or `c/@vm="1"` (value metadata index; **0-based or 1-based** in observed files).
2. **Value metadata** (`xl/metadata.xml`)
   - `vm` selects a `<valueMetadata><bk>` record (base varies; preserve and resolve best-effort).
   - That `<bk>` contains `<rc t="…" v="0"/>` where `v` is **0-based** into `futureMetadata name="XLRICHVALUE"`.
3. **Future metadata** (`xl/metadata.xml`)
   - `futureMetadata name="XLRICHVALUE"` `<bk>` #0 contains `<xlrd:rvb i="5"/>`.
   - `i=5` is the **0-based rich value index** into `xl/richData/richValue.xml`.
4. **Rich value** (`xl/richData/richValue.xml`)
   - Rich value record #5 is an “image” typed rich value.
   - Its payload contains a **relationship index** (e.g. `relIndex = 0`, **0-based**) into `richValueRel.xml`.
5. **Relationship table** (`xl/richData/richValueRel.xml`)
   - Relationship table entry #0 contains `r:id="rId7"`.
6. **OPC resolution** (`xl/richData/_rels/richValueRel.xml.rels`)
   - Relationship `Id="rId7"` resolves to an OPC `Target` (often a media part like `../media/image1.png`,
     but treat this as opaque; other targets/types may appear depending on Excel build).

So: **cell → vm (0/1-based) → metadata.xml → rvb@i (0-based) → richValue.xml → relIndex (0-based) →
richValueRel.xml → rId → .rels target → image bytes**.

### Variant B: `rc/@v` directly references the rich value index

Observed in `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`:

1. **Worksheet cell** has `c/@vm="0"`.
2. In `xl/metadata.xml`, the first `<valueMetadata><bk>` has `<rc t="1" v="0"/>`.
3. `v="0"` is treated as the rich value index (0-based) into `xl/richData/richValue.xml`.
4. The rich value record contains a relationship index into `richValueRel.xml`, which resolves via
   `xl/richData/_rels/richValueRel.xml.rels` to a media part.

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

* Excel appears to support multiple “kinds” (string/number/bool/rel/…).
* The **ordering** of `<member>` entries is significant: instances generally encode member values positionally.

### 3) `xl/richData/richValueRel.xml`

Stores a **vector/table** of `r:id` strings. The *index* into this vector is what rich values store.

Observed in `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <!-- Table position = relationship index (0-based). -->
  <rel r:id="rId1"/>
  <!-- ... -->
</richValueRel>
```

Another observed variant (inspected `rust_xlsxwriter` output; see [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md)) uses a different root name + namespace:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRels>
```

Some variants may wrap the entries (e.g. `<rels><rel .../></rels>`); match on element local-names and
preserve unknown structure when round-tripping.

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

Observed in `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` (image rich value referencing relationship index `0`):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <!-- Rich value index is typically the 0-based order of <rv> records (unless an explicit id/index is provided). -->
  <rv s="0" t="image">
    <!-- Relationship index (0-based) into richValueRel.xml -->
    <v>0</v>
  </rv>
</rvData>
```

Other builds may:

* split values across `xl/richData/richValue1.xml`, `richValue2.xml`, ...
* include an explicit global index attribute on `<rv>` (e.g. `i="…"`, `id="…"`, `idx="…"`)
* include multiple `<v>` members, with types indicated by attributes like `t="rel"` / `t="r"` / etc.

Notes:

* The “rich value index” is 0-based. Depending on the `metadata.xml` schema, the cell metadata may
  reference it either:
  * directly (e.g. `rc/@v = richValueIndex`), or
  * indirectly via a `rvb/@i` lookup table.
* The “relationship index” stored in the payload is 0-based and indexes into `richValueRel.xml`.

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
  * Observed in `fixtures/xlsx/metadata/rich-values-vm.xlsx`
* `http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata`
  * Observed in `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`

Additionally, `xl/workbook.xml` may include a `<metadata r:id="..."/>` element pointing at the relationship
ID for the metadata part (observed in `fixtures/xlsx/metadata/rich-values-vm.xlsx`). Some workbooks omit
this element and only include the relationship in `workbook.xml.rels` (observed in
`fixtures/xlsx/basic/image-in-cell-richdata.xlsx`). Preserve whichever representation the source workbook
uses.

The richValue relationships are Microsoft-specific. Observed in this repo:

* `http://schemas.microsoft.com/office/2017/06/relationships/richValue` → `xl/richData/richValue.xml`
* `http://schemas.microsoft.com/office/2017/06/relationships/richValueRel` → `xl/richData/richValueRel.xml`
  * Observed in `fixtures/xlsx/basic/image-in-cell-richdata.xlsx`
* `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue` → `xl/richData/rdrichvalue.xml` (and related rdRichValue tables)
  * Observed (via assertions) in `crates/formula-xlsx/tests/embedded_images_place_in_cell_roundtrip.rs`
* `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure` → `xl/richData/rdrichvaluestructure.xml`
  * Observed by inspecting the generated “Place in Cell” workbook (see [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md))
* `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes` → `xl/richData/rdRichValueTypes.xml`
  * Observed by inspecting the generated “Place in Cell” workbook (see [`docs/xlsx-embedded-images-in-cells.md`](./xlsx-embedded-images-in-cells.md))
* `http://schemas.microsoft.com/office/2022/10/relationships/richValueRel` → `xl/richData/richValueRel.xml`
  * Observed (via assertions) in `crates/formula-xlsx/tests/embedded_images_place_in_cell_roundtrip.rs`

Likely (not observed in fixtures here, but expected for richer payloads):

* `http://schemas.microsoft.com/office/2017/06/relationships/richValueTypes` → `xl/richData/richValueTypes.xml`
* `http://schemas.microsoft.com/office/2017/06/relationships/richValueStructure` → `xl/richData/richValueStructure.xml`

Some workbooks may instead relate the richData parts from `xl/metadata.xml` via `xl/_rels/metadata.xml.rels`.
For parsing and round-trip safety, treat both relationship layouts as valid.

Implementation guidance:

* When parsing, do not hardcode exact Type URIs; match by resolved `Target` path when necessary and preserve unknown relationship types.
* When writing new files, keep relationship IDs stable and prefer “append-only” updates. Excel may rewrite
  relationship type URIs and renumber `rId*` values.

#### Observed values summary (from in-repo fixtures/tests)

These values are copied from fixtures/tests (and inspected generated workbooks) in this repo and are safe to treat as “known in the wild”:

| Kind | Value | Source |
|------|-------|--------|
| Workbook → metadata relationship Type | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata` | `fixtures/xlsx/metadata/rich-values-vm.xlsx` |
| Workbook → metadata relationship Type | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata` | `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` |
| Workbook → richValue relationship Type | `http://schemas.microsoft.com/office/2017/06/relationships/richValue` | `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` |
| Workbook → richValueRel relationship Type | `http://schemas.microsoft.com/office/2017/06/relationships/richValueRel` | `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` |
| Workbook → rdRichValue relationship Type | `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue` | `crates/formula-xlsx/tests/embedded_images_place_in_cell_roundtrip.rs` (asserted substring) |
| Workbook → rdRichValueStructure relationship Type | `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure` | `docs/xlsx-embedded-images-in-cells.md` (inspected `rust_xlsxwriter` output) |
| Workbook → rdRichValueTypes relationship Type | `http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes` | `docs/xlsx-embedded-images-in-cells.md` (inspected `rust_xlsxwriter` output) |
| Workbook → richValueRel relationship Type | `http://schemas.microsoft.com/office/2022/10/relationships/richValueRel` | `crates/formula-xlsx/tests/embedded_images_place_in_cell_roundtrip.rs` (asserted substring) |
| `richValueRel.xml` namespace | `http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2` | `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` |
| `richValueRel.xml` namespace | `http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel` | `docs/xlsx-embedded-images-in-cells.md` (inspected `rust_xlsxwriter` output) |
| `richValue.xml` namespace | `http://schemas.microsoft.com/office/spreadsheetml/2017/richdata` | `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` |
| `metadata.xml` content type override | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml` | `fixtures/xlsx/metadata/rich-values-vm.xlsx` |
| `metadata.xml` content type override | `application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml` | `crates/formula-xlsx/tests/metadata_rich_value_roundtrip.rs` |
| `richValue.xml` content type override | `application/vnd.ms-excel.richvalue+xml` | `crates/formula-xlsx/tests/rich_data_workbook_structure_edits.rs` (synthetic fixture) |
| `richValueRel.xml` content type override | `application/vnd.ms-excel.richvaluerel+xml` | `crates/formula-xlsx/tests/rich_data_workbook_structure_edits.rs` (synthetic fixture) |
| `rdrichvalue.xml` content type override | `application/vnd.ms-excel.rdrichvalue+xml` | `docs/xlsx-embedded-images-in-cells.md` (inspected `rust_xlsxwriter` output) |
| `rdrichvaluestructure.xml` content type override | `application/vnd.ms-excel.rdrichvaluestructure+xml` | `docs/xlsx-embedded-images-in-cells.md` (inspected `rust_xlsxwriter` output) |
| `rdRichValueTypes.xml` content type override | `application/vnd.ms-excel.rdrichvaluetypes+xml` | `docs/xlsx-embedded-images-in-cells.md` (inspected `rust_xlsxwriter` output) |
| No override for metadata/richData XML parts (default `application/xml`) | (none) | `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` |

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

`xl/_rels/metadata.xml.rels` (optional; metadata → richData tables):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdRichValue"
                Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue"
                Target="richData/richValue.xml"/>
  <Relationship Id="rIdRichValueRel"
                Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel"
                Target="richData/richValueRel.xml"/>
  <!-- richValueTypes/richValueStructure relationships may also appear (unverified) -->
</Relationships>
```

### `[Content_Types].xml` considerations

Excel may or may not add explicit `<Override>` entries for `xl/metadata.xml` and `xl/richData/*`.
Observed patterns in this repo:

* `fixtures/xlsx/basic/image-in-cell-richdata.xlsx` relies on the default:
  * `<Default Extension="xml" ContentType="application/xml"/>`
  and includes no overrides for `metadata.xml` or `xl/richData/*`.
* `fixtures/xlsx/metadata/rich-values-vm.xlsx` includes an override:
  * `<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>`
* Some tests construct workbooks that use:
  * `application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml` for `/xl/metadata.xml`

For the richData tables themselves, producers can emit Microsoft-specific content types.

Observed in this repo:

```xml
<!-- Unprefixed “richValue*” naming (synthetic fixture) -->
<Override PartName="/xl/richData/richValue.xml"    ContentType="application/vnd.ms-excel.richvalue+xml"/>
<Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>

<!-- “rdRichValue*” naming (inspected rust_xlsxwriter output; see docs/xlsx-embedded-images-in-cells.md) -->
<Override PartName="/xl/richData/rdrichvalue.xml"          ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>
<Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>
<Override PartName="/xl/richData/rdRichValueTypes.xml"     ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>
<Override PartName="/xl/richData/richValueRel.xml"         ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
```

Likely patterns for additional tables (not yet verified against a real Excel-generated workbook that emits explicit overrides):

```xml
<Override PartName="/xl/richData/richValueTypes.xml"     ContentType="application/vnd.ms-excel.richvaluetypes+xml"/>
<Override PartName="/xl/richData/richValueStructure.xml" ContentType="application/vnd.ms-excel.richvaluestructure+xml"/>
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
5. Parse `richValueRel.xml` into `rel_index -> rId`.
6. Parse `xl/richData/_rels/richValueRel.xml.rels` into `rId -> target` (image path).
7. Parse `richValue.xml` into a table of `rich_value_index -> {type_id, payload...}`.
8. Parse `xl/metadata.xml` to resolve `vm` (cell attribute) → `rich_value_index` (best-effort; schemas vary:
   some use `xlrd:rvb/@i`, others may reference the rich value index directly).
9. Use worksheet cell `vm` values to map cells → `rich_value_index`.

For writing, the safest approach is typically “append-only”:

* Append new relationships to `richValueRel.xml` + `.rels`
* Append new rich values to `richValue.xml`
* Avoid renumbering existing indices unless you fully rebuild all referencing metadata
