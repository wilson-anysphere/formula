# Excel RichData (`richValue*`) parts for Images-in-Cell (`IMAGE()` / “Place in Cell”)

Excel’s “Images in Cell” feature (insert picture → **Place in Cell**, and the `IMAGE()` function) is backed by a **RichData / RichValue** subsystem. Rather than embedding image references directly in worksheet cell XML, Excel stores *typed rich value instances* in workbook-level parts under `xl/richData/`, then attaches cells to those instances via metadata.

This note documents the **expected part set**, the **role of each part**, and the **minimal XML shapes** needed to parse/write Excel-generated files.

For the overall “images in cells” packaging overview (including `xl/cellimages.xml`, `xl/metadata.xml`,
and current Formula status/tests), see: [20-images-in-cells.md](./20-images-in-cells.md).

> Status: best-effort reverse engineering. Exact namespaces / relationship-type URIs may vary by Excel version; preserve unknown attributes and namespaces when round-tripping.

---

## Expected part set (workbook-level)

When a workbook contains at least one RichData value (including images-in-cell), Excel typically adds:

```
xl/
  richData/
    richValue.xml
    richValueRel.xml
    richValueTypes.xml
    richValueStructure.xml
  richData/_rels/
    richValueRel.xml.rels   # required if richValueRel.xml contains r:id entries
```

### Roles (high level)

| Part | Purpose |
|------|---------|
| `xl/richData/richValueTypes.xml` | Defines **type IDs** (numeric) and links each type to a **structure ID** (string) that describes its field layout. |
| `xl/richData/richValueStructure.xml` | Defines **structures**: ordered field/member layouts keyed by **string IDs**. |
| `xl/richData/richValue.xml` | Stores the **rich value instances** (“objects”) in a workbook-global table. Each instance references a type (and/or structure) and stores member values. |
| `xl/richData/richValueRel.xml` | Stores a **relationship-ID table** (`r:id` strings) that can be referenced **by index** from rich values, avoiding embedding raw `rId*` strings inside each rich value payload. |
| `xl/richData/_rels/richValueRel.xml.rels` | OPC relationships for the `r:id` entries in `richValueRel.xml` (e.g. to `../media/imageN.png`). |

---

## How Excel wires cells to rich values (context)

The RichData parts above are workbook-global tables. A worksheet cell does **not** point directly at `xl/richData/richValue.xml`.

Instead, Excel uses **cell metadata** in `xl/metadata.xml`:

1. Worksheet cells use `c/@vm` (value-metadata index).
2. `vm` selects a `<valueMetadata><bk>` record in `xl/metadata.xml`.
3. That `<bk>` contains an `<rc t="…" v="…"/>` record pointing at `futureMetadata name="XLRICHVALUE"`.
4. The `futureMetadata` record contains an extension element (commonly `xlrd:rvb`) with `i="…"`, which is the
   **0-based rich value index** into `xl/richData/richValue.xml`.

Minimal representative shape (index bases are important; see below):

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
    <!-- vm=1 (1-based) selects the first <bk> -->
    <bk><rc t="1" v="0"/></bk>
  </valueMetadata>
</metadata>
```

This indirection is important for engineering because:

* `vm` indexes are **independent** from `richValue.xml` indexes.
* `vm` is **1-based** in Excel’s `xl/metadata.xml` model (see [Index bases](#index-bases--indirection)).

---

## Index bases & indirection

Excel uses multiple indices; mixing bases is a common source of bugs.

### `vm` (worksheet cell attribute) — **1-based**

In worksheet XML, cells can carry `vm="n"` to attach value metadata:

```xml
<c r="B2" t="str" vm="1">
  <v>…</v>
</c>
```

Current Formula behavior:

* `vm` is treated as **1-based** (i.e. `vm="1"` refers to the *first* `<valueMetadata><bk>` record).
  See `crates/formula-xlsx/src/rich_data/metadata.rs` (and its unit tests).
* Missing `vm` means “no value metadata”.
* Preserve unusual values like `vm="0"` if encountered (even if they don’t resolve cleanly).

### Indices inside `xl/metadata.xml` used by `XLRICHVALUE`

| Index | Location | Base | Meaning |
|------:|----------|------|---------|
| `t` | `<valueMetadata><bk><rc t="…">` | 1-based | index into `<metadataTypes>` (selects `metadataType name="XLRICHVALUE"`) |
| `v` | `<valueMetadata><bk><rc v="…">` | 0-based | index into `<futureMetadata name="XLRICHVALUE"><bk>` |
| `i` | `<xlrd:rvb i="…"/>` | 0-based | rich value index into `xl/richData/richValue.xml` |

### `richValue.xml` rich value table — **0-based**

Rich values are stored in a list; the rich value index is **0-based** and is referenced from `xl/metadata.xml`
via `xlrd:rvb/@i`.

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

1. **Worksheet cell** (`xl/worksheets/sheetN.xml`)
   - Cell has `c/@vm="1"` (value metadata index, **1-based**).
2. **Value metadata** (`xl/metadata.xml`)
   - `vm=1` selects `<valueMetadata><bk>` #1 (**1-based**).
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
   - Relationship `Id="rId7"` resolves to `Target="../media/image1.png"`.

So: **cell → vm (1-based) → metadata.xml → rvb@i (0-based) → richValue.xml → relIndex (0-based) →
richValueRel.xml → rId → .rels target → image bytes**.

## Minimal XML skeletons (best-effort)

These skeletons aim to show **roots, key child tags, and key attributes** as Excel tends to emit them. Namespaces and some attribute names may differ across builds—treat them as *shape guidance*, not a strict schema.

### 1) `xl/richData/richValueTypes.xml`

Defines **type IDs** (numeric) and links to a **structure ID** (string).

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypes xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <!-- One entry per type used in this workbook. -->
  <!-- Type IDs are numeric; richValue.xml instances reference them. -->
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

### 2) `xl/richData/richValueStructure.xml`

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

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <!-- Table position = relationship index (0-based). -->
  <rels>
    <rel r:id="rId1"/>
    <rel r:id="rId2"/>
    <!-- ... -->
  </rels>
</rvRel>
```

And the corresponding OPC relationships part:

`xl/richData/_rels/richValueRel.xml.rels`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship
    Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image1.png"/>
  <!-- ... -->
</Relationships>
```

### 4) `xl/richData/richValue.xml`

Stores the actual rich value instances. Each instance references a type (and/or structure) and encodes member values (often positionally, guided by `richValueStructure.xml`).

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <!-- Table position = rich value index (0-based). -->
  <values>
    <rv type="0">
      <!-- The exact payload encoding varies; this is illustrative. -->
      <!-- Example: relationship index 0 => richValueRel.xml entry 0 => r:id => image target -->
      <v kind="rel">0</v>
      <v kind="string">Alt text</v>
    </rv>
    <!-- ... -->
  </values>
</rvData>
```

Notes:

* The “rich value index” is 0-based and is referenced from `xl/metadata.xml` via `xlrd:rvb/@i`. In practice,
  Excel appears to treat this as the record’s 0-based index within the `richValue.xml` table.
* The “relationship index” stored in the payload is 0-based and indexes into `richValueRel.xml`.

---

## OPC relationships and `[Content_Types].xml`

### Workbook → richData parts

Excel must relate the workbook to the richData parts via `xl/_rels/workbook.xml.rels`. Relationship *targets* are typically:

* `richData/richValue.xml`
* `richData/richValueRel.xml`
* `richData/richValueTypes.xml`
* `richData/richValueStructure.xml`

Relationship **Type URIs** are Microsoft-specific and not yet verified in this repo. Likely patterns include:

* `http://schemas.microsoft.com/office/.../relationships/richValue`
* `http://schemas.microsoft.com/office/.../relationships/richValueRel`
* `http://schemas.microsoft.com/office/.../relationships/richValueTypes`
* `http://schemas.microsoft.com/office/.../relationships/richValueStructure`

Implementation guidance:

* When parsing, do not hardcode exact Type URIs; match by `Target` path when necessary and preserve unknown relationship types.
* When writing new files, choose a single consistent set of Type URIs and keep them stable (but be prepared that Excel may rewrite them).

### `[Content_Types].xml` overrides (likely)

Excel adds explicit `<Override>` entries for each richData part. Exact `ContentType` strings are not yet verified here; likely patterns:

```xml
<Override PartName="/xl/richData/richValue.xml"          ContentType="application/vnd.ms-excel.richvalue+xml"/>
<Override PartName="/xl/richData/richValueRel.xml"       ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
<Override PartName="/xl/richData/richValueTypes.xml"     ContentType="application/vnd.ms-excel.richvaluetypes+xml"/>
<Override PartName="/xl/richData/richValueStructure.xml" ContentType="application/vnd.ms-excel.richvaluestructure+xml"/>
```

Implementation guidance:

* When round-tripping an existing file: preserve the original overrides verbatim.
* When generating from scratch: emit overrides (do not rely on `Default Extension="xml" …`), as Excel tends to do for non-standard parts.

---

## Practical parsing strategy (recommended)

For images-in-cell, the minimum viable read path usually looks like:

1. Parse workbook relationships to locate the four `xl/richData/*` parts (if present).
2. Parse `richValueTypes.xml` into `type_id -> structure_id`.
3. Parse `richValueStructure.xml` into `structure_id -> ordered member schema`.
4. Parse `richValueRel.xml` into `rel_index -> rId`.
5. Parse `xl/richData/_rels/richValueRel.xml.rels` into `rId -> target` (image path).
6. Parse `richValue.xml` into a table of `rich_value_index -> {type_id, payload...}`.
7. Parse `xl/metadata.xml` to resolve `vm` (cell attribute) → `rich_value_index` (`xlrd:rvb/@i`).
8. Use worksheet cell `vm` values to map cells → `rich_value_index`.

For writing, the safest approach is typically “append-only”:

* Append new relationships to `richValueRel.xml` + `.rels`
* Append new rich values to `richValue.xml`
* Avoid renumbering existing indices unless you fully rebuild all referencing metadata
