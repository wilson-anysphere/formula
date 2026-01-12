# `cell-images.xlsx` (synthetic `xl/cellImages.xml` casing-variant fixture)

This fixture is a **synthetic** workbook (tagged in `docProps/app.xml` as `Application=Formula Fixtures`).
It exists to exercise:

- the **`xl/cellImages.xml`** (camel-case) part name (OPC part names are case-sensitive), and
- variability in the **workbook → cellImages relationship `Type` URI**.

See also:

- [`docs/20-images-in-cells.md`](../../../docs/20-images-in-cells.md) (cellImages store part overview)
- [`docs/02-xlsx-compatibility.md`](../../../docs/02-xlsx-compatibility.md) (relationship/content-type preservation strategy)

## ZIP part inventory (relevant parts)

```text
[Content_Types].xml
xl/cellImages.xml
xl/_rels/cellImages.xml.rels
xl/media/image1.png
```

## Provenance

`docProps/app.xml` contains:

```xml
<Application>Formula Fixtures</Application>
```

## Key relationships

In `xl/_rels/workbook.xml.rels`, the workbook links to `xl/cellImages.xml`:

```xml
<Relationship Id="rId3"
              Type="http://schemas.microsoft.com/office/2023/02/relationships/cellImage"
              Target="cellImages.xml"/>
```

The corresponding `.rels` part links to the image bytes:

```xml
<Relationship Id="rId1"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/image1.png"/>
```

## `xl/cellImages.xml`

This fixture uses a minimal `<cellImage><a:blip …/></cellImage>` shape:

```xml
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2023/02/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <cellImage>
    <a:blip r:embed="rId1"/>
  </cellImage>
</cellImages>
```

Real Excel workbooks can emit different namespaces and/or embed a full DrawingML `<xdr:pic>` subtree.
Treat this file as a **minimal** casing/URI-variability fixture, not Excel ground truth.

