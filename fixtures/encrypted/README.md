# Encrypted workbook fixtures

This directory is the canonical location for **password-to-open / encrypted** Excel workbook
fixtures used by tests that validate **format detection** and **error handling**.

This is **file encryption** (“Encrypt with Password”), not workbook/worksheet protection (“password
to edit”).

## Why this is separate from `fixtures/xlsx/`

Excel “password to open” OOXML spreadsheets (e.g. `.xlsx`, `.xlsm`, `.xlsb`) are **not ZIP
archives**. They are OLE/CFB (Compound File Binary) containers with `EncryptionInfo` and
`EncryptedPackage` streams (MS-OFFCRYPTO).

The ZIP/OPC round-trip harness (`crates/xlsx-diff`) enumerates its corpus via
`xlsx-diff::collect_fixture_paths(fixtures/xlsx/...)` and then opens each file as a ZIP archive.
Putting encrypted OOXML fixtures under `fixtures/xlsx/` would cause those round-trip tests to fail
during ZIP parsing.

## Layout

```
fixtures/encrypted/
  ooxml/      # Encrypted OOXML spreadsheets (e.g. `.xlsx`, `.xlsm`, `.xlsb` OLE/CFB containers)
```

Tests that need encrypted fixtures should reference these paths **explicitly** (they are not part
of the round-trip corpus).

Note: Some encryption tests build minimal encrypted containers programmatically (see
`crates/formula-io/tests/encrypted_xls.rs` and the synthetic container in
`crates/formula-io/tests/encrypted_ooxml.rs`). For end-to-end “password required” regression tests,
we also keep small in-repo encrypted OOXML fixtures under `fixtures/encrypted/ooxml/` (including
empty-password + Unicode-password samples, macro-enabled `.xlsm` fixtures, and multi-segment
`*-large.xlsx` variants; see `fixtures/encrypted/ooxml/README.md` for passwords and provenance).

For background on Excel encryption formats and terminology, see `docs/21-encrypted-workbooks.md`.

## Regenerating fixtures

Encrypted OOXML fixtures under `fixtures/encrypted/ooxml/` can be regenerated without Excel:

- Preferred/documented workflow (matches committed fixtures): see `fixtures/encrypted/ooxml/README.md`
- Alternative generator (Apache POI): `tools/encrypted-ooxml-fixtures/generate.sh`
