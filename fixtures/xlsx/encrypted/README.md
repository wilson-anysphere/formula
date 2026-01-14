# Encrypted OOXML fixtures (legacy)

This directory contains small **password-to-open** OOXML fixtures stored as **OLE/CFB** (Compound File
Binary) containers with `EncryptionInfo` + `EncryptedPackage` streams.

Why this is under `fixtures/xlsx/`:

- The ZIP/OPC round-trip harness (`crates/xlsx-diff`) enumerates `fixtures/xlsx/**` and then opens each
  file as a ZIP archive.
- Encrypted OOXML workbooks are **not ZIP files on disk**, so `xlsx-diff::collect_fixture_paths`
  intentionally skips the `fixtures/xlsx/encrypted/` subtree.

Note: these particular fixtures are **not** referenced by the main encryption test suite today. The
canonical encrypted fixture inventory lives under:

- `fixtures/encrypted/` (real-world encrypted `.xlsx`/`.xlsb`/`.xls`)
- `fixtures/encrypted/ooxml/` (vendored encrypted `.xlsx`/`.xlsm` corpus + passwords)

## Fixtures

- `agile-password.xlsx` — ECMA-376 **Agile** encryption
- `standard-password.xlsx` — ECMA-376 **Standard** encryption

Password for both files: `Password1234_`

These fixtures are self-generated from `fixtures/xlsx/basic/basic.xlsx` and are kept for
convenience/manual inspection.
