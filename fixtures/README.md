# Fixtures

This directory contains small, in-repo workbooks used by tests and documentation.

## Layout

```
fixtures/
  xlsx/         # ZIP/OPC-based XLSX/XLSM fixtures used by the `xlsx-diff` round-trip harness
  encrypted/    # Password-to-open/encrypted workbooks (OOXML-in-OLE + legacy `.xls`; excluded from the ZIP/OPC round-trip corpus)
  charts/       # Chart-specific fixtures + generated models + Excel golden PNGs
```

### `fixtures/xlsx/`

The main XLSX fixture corpus used by round-trip validation (load → save → diff). See
`fixtures/xlsx/README.md`.

### `fixtures/encrypted/`

Encrypted/password-protected Excel workbooks. Encrypted OOXML `.xlsx`/`.xlsm`/`.xlsb` files are
**OLE/CFB containers**, not ZIP archives, so they are excluded from the ZIP/OPC round-trip corpus
under `fixtures/xlsx/` (which is enumerated by `xlsx-diff::collect_fixture_paths`). See
`fixtures/encrypted/README.md`.

### `fixtures/charts/`

Chart regression fixtures, model dumps, and Excel-rendered golden images. See
`fixtures/charts/README.md`.
