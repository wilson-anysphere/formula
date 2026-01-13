# Encrypted OOXML spreadsheet fixtures

This directory is the home for Excel OOXML spreadsheets that require a password to open (for
example `.xlsx`, `.xlsm`, `.xlsb`).

Even though they use spreadsheet extensions like `.xlsx` / `.xlsm` / `.xlsb`, encrypted OOXML
workbooks are **not ZIP/OPC packages**. Excel stores them as OLE/CFB (Compound File Binary)
containers with `EncryptionInfo` + `EncryptedPackage` streams (see MS-OFFCRYPTO).

These fixtures live outside `fixtures/xlsx/` so they are not picked up by
`xlsx-diff::collect_fixture_paths` (which drives the ZIP-based round-trip test corpus).

## Fixtures

- `agile.xlsx` – encrypted OOXML (Agile encryption; MS-OFFCRYPTO `EncryptionInfo` version 4.4)
- `standard.xlsx` – encrypted OOXML (Standard encryption; MS-OFFCRYPTO `EncryptionInfo` version 3.2)

These are used by:

- `crates/formula-io/tests/encrypted_ooxml.rs` (ensures `open_workbook` / `open_workbook_model`
  surface a password/encryption error when no password is supplied)
- `crates/formula-io/tests/encrypted_ooxml_fixtures.rs` (format sniffing/detection; optional)

These fixtures are currently only used to exercise the “password required” error path, so the
actual passwords are not needed by tests (Formula does not decrypt encrypted workbooks yet).

## Regenerating fixtures (without Excel)

Use the Apache POI-based generator:

- Source: `tools/encrypted-ooxml-fixtures/GenerateEncryptedXlsx.java`
- Wrapper script (downloads pinned jars + verifies SHA256): `tools/encrypted-ooxml-fixtures/generate.sh`

Apache POI version: **5.2.5** (see `tools/encrypted-ooxml-fixtures/generate.sh` for the full pinned jar set).

Example (from repo root):

```bash
# Pick any plaintext `.xlsx` as the payload.
PLAINTEXT=fixtures/xlsx/basic/basic.xlsx

tools/encrypted-ooxml-fixtures/generate.sh agile password "$PLAINTEXT" fixtures/encrypted/ooxml/agile.xlsx
tools/encrypted-ooxml-fixtures/generate.sh standard password "$PLAINTEXT" fixtures/encrypted/ooxml/standard.xlsx
```

Note: the generated encrypted files are not expected to be byte-for-byte stable across runs
(encryption uses random salts/IVs). They should still be structurally valid encrypted OOXML containers.

## Passwords

If you add additional encrypted workbook fixtures intended for decryption tests, document the expected
password in the adjacent README (or encode it in the filename) so tests can open them deterministically.

See `docs/21-encrypted-workbooks.md` for details on OOXML encryption containers (`EncryptionInfo` /
`EncryptedPackage`).
