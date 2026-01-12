# `linked-data-types-excel.xlsx` (Excel Linked Data Types fixture: Stocks + Geography)

This fixture is an `.xlsx` intended to represent an **Excel 365** workbook containing **Linked Data Types**
(a.k.a. rich data types) stored via the Rich Data / Rich Value pipeline:

- worksheet cells with `vm="…"` / `cm="…"` metadata indices
- `xl/metadata.xml` + `xl/_rels/metadata.xml.rels`
- `xl/richData/*` parts (`richValueTypes`, `richValueStructure`, `richValue`)

## Key cells

- `Sheet1!A1`: Stocks linked data type (`MSFT`)
- `Sheet1!A2`: Geography linked data type (`Seattle`)

## Provenance: confirm this file was saved by Excel

`docProps/app.xml` contains:

```xml
<Application>Microsoft Excel</Application>
<AppVersion>16.0300</AppVersion>
```

## ZIP part inventory (complete)

Output of:

```bash
unzip -Z1 fixtures/xlsx/rich-data/linked-data-types-excel.xlsx | sort
```

```text
[Content_Types].xml
_rels/.rels
docProps/app.xml
docProps/core.xml
xl/_rels/metadata.xml.rels
xl/_rels/workbook.xml.rels
xl/metadata.xml
xl/richData/richValue.xml
xl/richData/richValueRel.xml
xl/richData/richValueStructure.xml
xl/richData/richValueTypes.xml
xl/styles.xml
xl/workbook.xml
xl/worksheets/sheet1.xml
```

## Worksheet cells (`vm`/`cm`)

From `xl/worksheets/sheet1.xml`:

```xml
<c r="A1" t="inlineStr" vm="1" cm="1"><is><t>MSFT</t></is></c>
<c r="A2" t="inlineStr" vm="2" cm="2"><is><t>Seattle</t></is></c>
```

## Rich data parts

- `xl/metadata.xml` wires `vm` indices to rich value indices (`xlrd:rvb i="…"`).
- `xl/_rels/metadata.xml.rels` relates `xl/metadata.xml` → `xl/richData/*`.
- `xl/richData/richValueTypes.xml` includes type names:
  - `com.microsoft.excel.stocks`
  - `com.microsoft.excel.geography`
