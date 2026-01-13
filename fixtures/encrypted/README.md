# Encrypted workbook fixtures

This directory holds password-protected / encrypted Excel workbooks used by tests that validate
**format detection** and **error handling**.

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
