# `cellimages.xlsx` (synthetic `xl/cellimages.xml` store fixture)

This fixture is a **synthetic** workbook (tagged in `docProps/app.xml` as `Application=Formula Fixtures`).
It exists to exercise Formula’s support for **preserving and extracting images** from the optional
workbook-level `xl/cellimages.xml` part.

See also:

- [`docs/20-images-in-cells.md`](../../../docs/20-images-in-cells.md) (overall images-in-cells packaging + cellImages store notes)
- [`docs/02-xlsx-compatibility.md`](../../../docs/02-xlsx-compatibility.md) (round-trip strategy + URI variability)

## ZIP part inventory (relevant parts)

```text
[Content_Types].xml
xl/cellimages.xml
xl/_rels/cellimages.xml.rels
xl/media/image1.png
```

## Provenance

`docProps/app.xml` contains:

```xml
<Application>Formula Fixtures</Application>
```

## Key relationships

In `xl/_rels/workbook.xml.rels`, the workbook links directly to `xl/cellimages.xml`:

```xml
<Relationship Id="rId3"
              Type="http://schemas.microsoft.com/office/2022/relationships/cellImages"
              Target="cellimages.xml"/>
```

The `cellimages` part links to the image bytes via `xl/_rels/cellimages.xml.rels`:

```xml
<Relationship Id="rId1"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/image1.png"/>
```

## `xl/cellimages.xml` (lightweight `a:blip` shape)

This fixture uses a lightweight schema where the relationship ID is stored directly on an `<a:blip>`:

```xml
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage>
    <a:blip r:embed="rId1"/>
  </cellImage>
</cellImages>
```

Real Excel workbooks can also store `xl/cellimages.xml` but may emit a richer DrawingML subtree (e.g.
`<xdr:pic>`). Treat this fixture as a **minimal** preservation/regression sample, not Excel ground truth.

## Worksheet note

This fixture is focused on the standalone `cellimages` store part; its worksheet cells do **not** include
`vm="…"`/`cm="…"` attributes or rich-data wiring.
