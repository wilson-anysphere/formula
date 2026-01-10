# XLSX Fixture Corpus

These files are used by the round-trip validation harness (`crates/xlsx-diff`).

Goals:
- Keep fixtures **small** and **in-repo** so CI is fast and deterministic.
- Cover representative XLSX features incrementally (basic cells, formulas, styles, etc.).

## Layout

```
fixtures/xlsx/
  basic/
  formulas/
  styles/
  conditional-formatting/
  charts/
  pivots/
  macros/
```

`pivots/` is currently a scaffold for future expansion. `macros/` contains small
`.xlsm` fixtures used to validate VBA project preservation.

## Notable fixtures

- `basic/`:
  - `basic.xlsx` - minimal numeric + inline string.
  - `shared-strings.xlsx` - uses `xl/sharedStrings.xml`.
  - `multi-sheet.xlsx` - 2 worksheets.
- `styles/`:
  - `styles.xlsx` - simple bold cell style.
  - `rich-text-shared-strings.xlsx` - shared strings with rich-text runs.
- `pivots/`:
  - `pivot-fixture.xlsx` - minimal pivot table parts (cache definition/records + pivotTable).
- `charts/`:
  - `basic-chart.xlsx` - minimal chart parts (drawing + chart referencing sheet data).
  - `bar.xlsx`, `line.xlsx`, `pie.xlsx`, `scatter.xlsx` - small fixtures for common chart types.
- `macros/`:
  - `basic.xlsm` - minimal VBA project preservation fixture.

## Regenerating the initial fixtures

The initial `.xlsx` files are generated without external dependencies:

```bash
python3 fixtures/xlsx/generate_fixtures.py
```

The generator uses deterministic timestamps so diffs are stable.
