# RichData / Linked Data Types fixtures

This directory contains **small** `.xlsx` fixtures intended to exercise Excel’s
modern **Linked Data Types** / **RichData** packaging:

- `xl/metadata.xml`
- `xl/_rels/metadata.xml.rels`
- `xl/richData/*`
- `vm="…"` / `cm="…"` attributes on worksheet `<c>` (cell) elements

## `linked-data-types.xlsx`

`linked-data-types.xlsx` is meant to be a minimal workbook containing:

- `Sheet1!A1`: a **Stocks** linked data type (e.g. `MSFT`)
- `Sheet1!A2`: a **Geography** linked data type (e.g. `Seattle`)

### Regenerating (Excel 365)

1. Create a new blank workbook.
2. Enter `MSFT` in `A1`, then **Data** → **Data Types** → **Stocks**.
3. Enter `Seattle` in `A2`, then **Data** → **Data Types** → **Geography**.
4. Save as `linked-data-types.xlsx`.
5. Verify the rich-data parts exist:

```bash
unzip -Z1 fixtures/xlsx/richdata/linked-data-types.xlsx | rg '^xl/(metadata\\.xml|_rels/metadata\\.xml\\.rels|richData/)'
unzip -p fixtures/xlsx/richdata/linked-data-types.xlsx xl/worksheets/sheet1.xml | rg ' vm=| cm='
```

