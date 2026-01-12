# Locale function/error translation data (`*.tsv`)

The formula engine persists and evaluates formulas in **canonical Excel (en-US) form**:

- English function names (e.g. `SUM`, `CUBEVALUE`)
- `,` argument separators
- `.` decimal separator
- canonical error literals (e.g. `#VALUE!`, `#GETTING_DATA`)

For UI workflows, we support translating formulas to/from locale-specific display forms via:

- `locale::canonicalize_formula*` (localized -> canonical)
- `locale::localize_formula*` (canonical -> localized)
- `Engine::set_cell_formula_localized*`

## TSV format

Each `*.tsv` file is a simple tab-separated mapping:

```
Canonical<TAB>Localized
```

- Lines starting with `#` and empty lines are ignored.
- Function names should be provided in the **same spelling Excel displays for that locale**.
- Entries are treated as case-insensitive in parsing (function names are uppercased during
  translation), so the convention is to store them as uppercase.

## External-data functions / errors

These locales include explicit coverage for external-data worksheet functions:

- `RTD`
- `CUBEVALUE`
- `CUBEMEMBER`
- `CUBEMEMBERPROPERTY`
- `CUBERANKEDMEMBER`
- `CUBESET`
- `CUBESETCOUNT`
- `CUBEKPIMEMBER`

And for the external-data loading error literal:

- `#GETTING_DATA`

The expected spellings are encoded in:

- `de-DE.tsv`
- `fr-FR.tsv`
- `es-ES.tsv`

and exercised by `crates/formula-engine/tests/locale_parsing.rs` so any regression (e.g. locale
input turning into `#NAME?`) is caught.

### Newer external-data errors

The newer external-data errors (`#CONNECT!`, `#FIELD!`, `#BLOCKED!`, `#UNKNOWN!`) are currently
treated as canonical (English) for all supported locales, until we have a verified Excel
localization list for them. Tests assert they round-trip unchanged.

