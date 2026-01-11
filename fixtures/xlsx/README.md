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

## Notable fixtures

- `basic/`:
  - `basic.xlsx` - minimal numeric + inline string.
  - `shared-strings.xlsx` - uses `xl/sharedStrings.xml`.
  - `multi-sheet.xlsx` - 2 worksheets.
  - `comments.xlsx` - legacy comments parts.
  - `grouped-rows.xlsx` - outline/grouped rows metadata.
  - `image.xlsx` - embedded image (`xl/media/image1.png`) + drawing relationship.
  - `print-settings.xlsx` - page setup + print titles/areas.
- `formulas/`:
  - `formulas.xlsx` - simple formula + cached result.
- `metadata/`:
  - `row-col-properties.xlsx` - custom row height + hidden row, custom column width + hidden column.
  - `data-validation-list.xlsx` - simple list data validation (`<dataValidations>`).
  - `defined-names.xlsx` - workbook named ranges (`<definedNames>` in `xl/workbook.xml`).
  - `external-link.xlsx` - minimal external link parts (`xl/externalLinks/externalLink1.xml` + rels).
- `styles/`:
  - `styles.xlsx` - simple bold cell style.
  - `rich-text-shared-strings.xlsx` - shared strings with rich-text runs.
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

## Regenerating the initial fixtures

Some `.xlsx` files (the small synthetic ones) are generated without external dependencies:

```bash
python3 fixtures/xlsx/generate_fixtures.py
```

The generator uses deterministic timestamps so diffs are stable.
