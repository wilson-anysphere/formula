# Encrypted `.xls` fixtures

This directory holds **password-protected / encrypted** legacy Excel `.xls` samples used by
`formula-xls` tests.

Why a dedicated directory?

- `.xls` supports multiple legacy encryption schemes. Keeping these files together (and documented)
  avoids confusion when new schemes are added.
- Passwords are **intentionally non-secret test values**. Do **not** reuse any real passwords when
  creating fixtures.

## Location & naming convention

- **Location:** `crates/formula-xls/tests/fixtures/encrypted/`
- **Naming:** `biff<version>_<scheme>_pw_open.xls`
  - `biff8_xor_pw_open.xls`
  - `biff8_rc4_standard_pw_open.xls`
  - `biff8_rc4_cryptoapi_pw_open.xls`

## Fixture inventory

All fixtures are intentionally tiny.

- The XOR + “RC4 standard” fixtures are **detection fixtures** (minimal CFB + `Workbook` stream
  containing just enough BIFF to surface a `FILEPASS` record).
- The RC4 CryptoAPI fixture is a **real encrypted workbook** (still small) used to validate the
  decryption path.

`formula-xls` treats `FILEPASS` as a signal that the workbook is encrypted/password-protected.

- `import_xls_path` does **not** support encrypted workbooks and returns `ImportError::EncryptedWorkbook`.
- `import_xls_path_with_password` supports a subset of BIFF8 encryption: **RC4 CryptoAPI**
  (`wEncryptionType=0x0001`, `wEncryptionSubType=0x0002`).

Note: In BIFF8, both RC4 variants use `wEncryptionType=0x0001`; the `subType` field distinguishes
“RC4 standard” (`subType=0x0001`) from “RC4 CryptoAPI” (`subType=0x0002`).

| File | Encryption scheme | BIFF version | Created with | Test password |
|---|---|---:|---|---|
| `biff8_xor_pw_open.xls` | XOR (legacy obfuscation) | BIFF8 | `cargo test -p formula-xls --test regenerate_encrypted_xls_fixtures -- --ignored` (this repo; writes a tiny CFB/BIFF stream via `cfb` `0.10`) | `password` |
| `biff8_rc4_standard_pw_open.xls` | RC4 “standard” | BIFF8 | same as above | `password` |
| `biff8_rc4_cryptoapi_pw_open.xls` | RC4 (CryptoAPI) | BIFF8 | same as above; additionally used by `tests/import_encrypted_rc4_cryptoapi.rs` to validate `import_xls_path_with_password` | `correct horse battery staple` |

## Regenerating fixtures

### Preferred: repository generator (deterministic)

Run the ignored generator test:

```bash
cargo test -p formula-xls --test regenerate_encrypted_xls_fixtures -- --ignored
```

This overwrites the files in this directory.

### Manual: Excel / LibreOffice (UI steps)

Manual regeneration is acceptable if you **verify the encryption scheme** afterwards (different
apps/versions can choose different `.xls` encryption schemes).

**Excel (Windows)**

1. Create a new workbook with minimal contents (e.g. set `A1` to `1`).
2. `File` → `Save As…`
3. Choose file type: **Excel 97-2003 Workbook (`*.xls`)**
4. In the Save As dialog: `Tools` → `General Options…`
5. Set **Password to open** to the test password listed above (e.g. `password`).
6. Save.

**LibreOffice**

1. Create a new workbook with minimal contents.
2. `File` → `Save As…`
3. Choose file type: **Excel 97-2003 (`*.xls`)**
4. Check **Save with password**
5. Click `Save`, then enter the test password.

## Size expectations

- Keep each fixture **< 100KB** (preferably a few KB).
- Rationale:
  - Faster `cargo test` runs and less I/O.
  - Prevents committing large, hard-to-audit binary blobs.
  - Minimizes the risk of accidentally embedding real customer data.

## Security note

These passwords are **not secrets** and exist solely so tests can reference a known value.
Never commit real passwords or sensitive workbooks as fixtures.
