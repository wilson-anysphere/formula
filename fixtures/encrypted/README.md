# Encrypted workbook fixtures

This directory is the canonical location for **password-to-open / encrypted** Excel workbook
fixtures used by tests that validate **format detection** and **error handling**.

This is **file encryption** (“Encrypt with Password”), not workbook/worksheet protection (“password
to edit”).

## Why this is separate from `fixtures/xlsx/`

Excel “encrypted” `.xlsx` files are **not ZIP archives**. They are OLE/CFB (Compound File Binary)
containers with `EncryptionInfo` and `EncryptedPackage` streams (MS-OFFCRYPTO).

The XLSX round-trip harness (`crates/xlsx-diff`) enumerates its corpus via
`xlsx-diff::collect_fixture_paths(fixtures/xlsx/...)` and then opens each `.xlsx` as a ZIP/OPC
package. Putting encrypted `.xlsx` files under `fixtures/xlsx/` would cause those round-trip tests
to fail during ZIP parsing.

## Layout

```
fixtures/encrypted/
  ooxml/      # Encrypted OOXML workbooks (password-protected `.xlsx` OLE/CFB containers)
```

Tests that need encrypted fixtures should reference these paths **explicitly** (they are not part
of the round-trip corpus).

Note: Some encryption tests build minimal encrypted containers programmatically instead of checking
in binary files (see `crates/formula-io/tests/encrypted_ooxml.rs` and
`crates/formula-io/tests/encrypted_xls.rs`). If/when we add real Excel-saved encrypted workbooks to
the repo, they should live under `fixtures/encrypted/`.

For background on Excel encryption formats and terminology, see `docs/21-encrypted-workbooks.md`.
