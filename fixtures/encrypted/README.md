# Encrypted workbook fixtures

This directory is the canonical location for **password-to-open / encrypted** Excel workbook
fixtures used by tests that validate **format detection**, **error handling**, and (when enabled)
**decryption** compatibility.

This is **file encryption** (“Encrypt with Password”), not workbook/worksheet protection (“password
to edit”).

For the small, regeneratable encrypted workbook fixtures used by end-to-end decryption integration
tests (Agile/Standard OOXML + legacy `.xls` `FILEPASS`), see [`fixtures/encryption/`](../encryption/).

## Why this is separate from `fixtures/xlsx/`

Excel “password to open” OOXML spreadsheets (e.g. `.xlsx`, `.xlsm`, `.xlsb`) are **not ZIP
archives**. They are OLE/CFB (Compound File Binary) containers with `EncryptionInfo` and
`EncryptedPackage` streams (MS-OFFCRYPTO).

The ZIP/OPC round-trip harness (`crates/xlsx-diff`) enumerates its corpus via
`xlsx-diff::collect_fixture_paths(fixtures/xlsx/...)` and then opens each file as a ZIP archive.
Putting encrypted OOXML fixtures under `fixtures/xlsx/` would cause those round-trip tests to fail
during ZIP parsing. If you *do* need encrypted fixtures under `fixtures/xlsx/` (for example to pair
an encrypted file with its decrypted plaintext for end-to-end tests), they must live under
`fixtures/xlsx/encrypted/` which is explicitly skipped by `xlsx-diff::collect_fixture_paths`.

## Layout

```
fixtures/encrypted/
  ooxml/                # Encrypted OOXML spreadsheets used by targeted decryption tests
  encrypted_agile.xlsx  # Real-world encrypted XLSX (Agile)
  encrypted_standard.xlsx  # Real-world encrypted XLSX (Standard/CryptoAPI)
  encrypted.xlsb        # Real-world encrypted XLSB (OOXML-in-OLE)
  encrypted.xls         # Real-world encrypted legacy XLS (BIFF8 FILEPASS)
```

Tests that need encrypted fixtures should reference these paths **explicitly** (they are not part
of the round-trip corpus).

Note: Some encryption tests build minimal encrypted containers programmatically (see
`crates/formula-io/tests/encrypted_xls.rs` and the synthetic container in
`crates/formula-io/tests/encrypted_ooxml.rs`). For end-to-end “password required” regression tests,
we also keep small in-repo encrypted OOXML fixtures under `fixtures/encrypted/ooxml/` (including
empty-password + Unicode-password samples, macro-enabled `.xlsm` fixtures, and multi-segment
`*-large.xlsx` variants; see `fixtures/encrypted/ooxml/README.md` for passwords and provenance).
That directory includes `basic-password.xlsm`, an encrypted macro-enabled workbook used to validate
that `xl/vbaProject.bin` survives decryption.

For background on Excel encryption formats and terminology, see `docs/21-encrypted-workbooks.md`.

## Regenerating fixtures

Encrypted OOXML fixtures under `fixtures/encrypted/ooxml/` can be regenerated without Excel:

- Preferred/documented workflow (matches committed fixtures): see `fixtures/encrypted/ooxml/README.md`
- Alternative generator (Apache POI): `tools/encrypted-ooxml-fixtures/generate.sh`

## Real-world fixtures (used by `open_encrypted_fixtures.rs`)

These binary fixtures are used by `crates/formula-io/tests/open_encrypted_fixtures.rs` to ensure
that `formula-io` can decrypt and open **real-world** password-protected Excel files.

| File | Format | Encryption | Notes | Expected contents |
|---|---|---|---|---|
| `encrypted_agile.xlsx` | XLSX | ECMA-376 Agile (OOXML-in-OLE) | Sourced from `msoffcrypto-tool` test corpus | Password: `Password1234_` • `Sheet1!A1="lorem"`, `Sheet1!B1="ipsum"` |
| `encrypted_standard.xlsx` | XLSX | ECMA-376 Standard/CryptoAPI (OOXML-in-OLE) | Apache POI output (copied from `fixtures/encrypted/ooxml/standard-4.2.xlsx`) | Password: `password` • `Sheet1!A1=1`, `Sheet1!B1="Hello"` |
| `encrypted.xlsb` | XLSB | ECMA-376 Agile (OOXML-in-OLE) | Sourced from Apache POI test corpus (`protected_passtika.xlsb`) | Password: `tika` • `Sheet1!A1="You can't see me"` |
| `encrypted.xls` | XLS | BIFF8 `FILEPASS` RC4 CryptoAPI | Microsoft Excel-generated file; crosses the 1024-byte RC4 re-key boundary | Password: `password` • `Sheet1!A400=399`, `Sheet1!B400="RC4_BOUNDARY_OK"` |

### Provenance

`encrypted_agile.xlsx` is downloaded from the upstream `msoffcrypto-tool` repository:

- https://github.com/nolze/msoffcrypto-tool

`encrypted_standard.xlsx` is copied from the Apache POI-generated Standard fixture in this repo:

- `fixtures/encrypted/ooxml/standard-4.2.xlsx`

`encrypted.xlsb` is downloaded from the Apache POI repository:

- https://github.com/apache/poi

`encrypted.xls` is copied from the Excel-generated RC4 CryptoAPI boundary fixture in this repo:

- `fixtures/encryption/biff8_rc4_cryptoapi_boundary_pw_open.xls`

### Regenerating `encrypted.xlsb`

From the repo root:

```bash
curl -L -o fixtures/encrypted/encrypted.xlsb \
  https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/protected_passtika.xlsb
```
