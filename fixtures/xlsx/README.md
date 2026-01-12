# XLSX Fixture Corpus

These files are used by the round-trip validation harness (`crates/xlsx-diff`).

Goals:
- Keep fixtures **small** and **in-repo** so CI is fast and deterministic.
- Cover representative XLSX features incrementally (basic cells, formulas, styles, etc.).

## Layout

```
fixtures/xlsx/
  basic/
  metadata/
  rich-data/
  richdata/
  formulas/
  styles/
  conditional-formatting/
  hyperlinks/
  charts/
  pivots/
  macros/
```

`charts/` and `pivots/` will expand over time as we add more complex corpora.
`macros/` contains small `.xlsm` fixtures used to validate VBA project preservation.
`rich-data/` is reserved for modern Excel 365 “rich value” features (linked data
types and images-in-cells) which require ground-truth workbooks saved by Excel.

## Notable fixtures

- `basic/`:
  - `basic.xlsx` - minimal numeric + inline string.
  - `bool-error.xlsx` - boolean + error cell types.
  - `shared-strings.xlsx` - uses `xl/sharedStrings.xml`.
  - `multi-sheet.xlsx` - 2 worksheets.
  - `comments.xlsx` - legacy comments parts.
  - `grouped-rows.xlsx` - outline/grouped rows metadata.
  - `image.xlsx` - embedded image (`xl/media/image1.png`) + drawing relationship.
  - `image-in-cell.xlsx` - **real Excel** “Place in Cell” image values using `xl/metadata.xml` + `xl/richData/*` (no `xl/cellimages.xml`; see `basic/image-in-cell.md`).
  - `cellimages.xlsx` - includes an Excel-style in-cell image store part (`xl/cellimages.xml` + `.rels`) referencing `xl/media/image1.png`.
  - `activex-control.xlsx` - minimal ActiveX/form control parts (`xl/ctrlProps/*` + `xl/activeX/*`) with a worksheet `<controls>` fragment.
  - `print-settings.xlsx` - page setup + print titles/areas.
- `formulas/`:
  - `formulas.xlsx` - simple formula + cached result.
  - `formulas-stale-cache.xlsx` - formula with an intentionally stale cached result (used to ensure imports do not auto-recalc).
  - `shared-formula.xlsx` - shared formula range (`t="shared"`) with textless followers.
- `metadata/`:
  - `row-col-properties.xlsx` - custom row height + hidden row, custom column width + hidden column.
  - `data-validation-list.xlsx` - simple list data validation (`<dataValidations>`).
  - `rich-values-vm.xlsx` - workbook-level `xl/metadata.xml` part + workbook relationship + worksheet cell `vm="..."` rich-value binding (`futureMetadata` / `rvb`).
  - `defined-names.xlsx` - workbook named ranges (`<definedNames>` in `xl/workbook.xml`).
  - `external-link.xlsx` - minimal external link parts (`xl/externalLinks/externalLink1.xml` + rels).
- `styles/`:
  - `styles.xlsx` - simple bold cell style.
  - `rich-text-shared-strings.xlsx` - shared strings with rich-text runs.
- `rich-data/`:
  - See `rich-data/README.md` for how to generate fixtures that include
    `xl/metadata.xml`, `xl/richData/*`, and in-sheet `vm`/`cm` cell attributes
    (linked data types like Stocks/Geography, and images placed in cells).
 - `richdata/`:
   - `linked-data-types.xlsx` - minimal workbook exercising `xl/metadata.xml` + `xl/richData/*`
     parts and `vm`/`cm` cell attributes (used by
     `crates/formula-xlsx/tests/linked_data_types_fixture.rs`).
- `conditional-formatting/`:
  - `conditional-formatting.xlsx` - simple `cfRule` example.
  - `conditional-formatting-2007.xlsx` - Excel 2007-style conditional formatting.
  - `conditional-formatting-x14.xlsx` - Excel 2010+ (`x14`) conditional formatting.
- `hyperlinks/`:
  - `hyperlinks.xlsx` - external + internal hyperlinks (tests `hyperlinks` + rels).
- `pivots/`:
  - `pivot-fixture.xlsx` - minimal pivot table parts (cache definition/records + pivotTable).
- `charts/`:
  - `basic-chart.xlsx` - minimal chart parts (drawing + chart referencing sheet data).
  - `bar.xlsx`, `line.xlsx`, `pie.xlsx`, `scatter.xlsx` - small fixtures for common chart types.
- `macros/`:
  - `basic.xlsm` - minimal VBA project preservation fixture.
  - `signed-basic.xlsm` - `basic.xlsm` with a `\x05DigitalSignature` stream containing an
    Office DigSig wrapper (length-prefixed DigSigInfoSerialized-like header) whose PKCS#7/CMS
    `SignedData` payload embeds an Authenticode `SpcIndirectDataContent` digest (self-signed fixture
    cert; no private key).

## Regenerating the initial fixtures

Some `.xlsx` files (the small synthetic ones) are generated without external dependencies:

```bash
python3 fixtures/xlsx/generate_fixtures.py
```

The generator uses deterministic timestamps so diffs are stable.
