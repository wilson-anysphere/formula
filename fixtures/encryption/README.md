# Encrypted workbook fixtures

These fixtures are used to validate password-protected workbook import support.

## `biff8_rc4_cryptoapi_pw_open.xls`

- Format: legacy `.xls` (BIFF8)
- Encryption: `FILEPASS` RC4 CryptoAPI (Excel 97-2003 era)
- Password: `correct horse battery staple`

This file is also used by the `formula-xls` integration tests under
`crates/formula-xls/tests/fixtures/encrypted/`.
