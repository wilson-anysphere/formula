# Encrypted XLSX fixtures

These files are **password-protected** `.xlsx` workbooks (1 sheet with `A1=1`, `B1="Hello"`), used
by `crates/formula-io` integration tests to validate OOXML decryption support.

Fixtures:

- `agile-password.xlsx` — ECMA-376 **Agile** encryption
- `standard-password.xlsx` — ECMA-376 **Standard** encryption

Password for both files: `Password1234_`

These fixtures are self-generated from `fixtures/xlsx/basic/basic.xlsx`.
