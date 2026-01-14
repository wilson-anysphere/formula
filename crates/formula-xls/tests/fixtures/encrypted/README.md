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

All fixtures **in this directory** are intentionally tiny **real encrypted workbooks** generated
deterministically by `tests/regenerate_encrypted_xls_fixtures.rs`.

Some additional encrypted `.xls` edge-case fixtures live alongside this directory (for cases that
the deterministic generator does not currently encode). Those fixtures are documented below.

Each fixture contains:

- workbook globals with at least one `FONT` + `XF` record **after** `FILEPASS` (so those payload bytes
  are encrypted)
- a single worksheet `Sheet1` with `A1=42` and a non-default cell style (vertical alignment = Top)

The RC4 Standard fixture additionally ensures the encrypted record-data stream after `FILEPASS`
crosses the 1024-byte boundary (to exercise RC4 per-block rekeying).

`formula-xls` treats `FILEPASS` as a signal that the workbook is encrypted/password-protected.

- `import_xls_path` / `import_xls_bytes` (no password) do **not** support encrypted workbooks and
  return `ImportError::EncryptedWorkbook`.
- `import_xls_path_with_password` / `import_xls_bytes_with_password` support these legacy `.xls`
  encryption schemes:
  - XOR obfuscation (`wEncryptionType=0x0000`)
  - RC4 “standard” (`wEncryptionType=0x0001`, `wEncryptionSubType=0x0001`)
  - RC4 CryptoAPI (`wEncryptionType=0x0001`, `wEncryptionSubType=0x0002`)

Note: In BIFF8, both RC4 variants use `wEncryptionType=0x0001`; the `subType` field distinguishes
“RC4 standard” (`subType=0x0001`) from “RC4 CryptoAPI” (`subType=0x0002`).

### Password semantics (Excel legacy)

- **XOR obfuscation:** legacy `.xls` passwords are effectively limited to **15 characters**; extra
  characters are ignored.
- **RC4 “standard” truncation:** only the first **15 UTF-16 code units** of the password are
  significant; extra characters are ignored (so a 16-character password and its first 15 characters
  are treated as equivalent).
- **RC4 CryptoAPI:** uses the full password string (no 15-character truncation).
- **Empty passwords:** third-party writers can emit a `FILEPASS` workbook with an empty password;
  this is supported by the underlying key derivation. Excel UI flows may refuse to create such a
  file, but we keep a fixture to ensure we handle it correctly.

| File | Encryption scheme | BIFF version | Created with | Test password |
|---|---|---:|---|---|
| `biff8_xor_pw_open.xls` | XOR (legacy obfuscation) | BIFF8 | `cargo test -p formula-xls --test regenerate_encrypted_xls_fixtures -- --ignored` (this repo; writes a tiny CFB/BIFF stream via `cfb` `0.10`) | `password` |
| `biff8_xor_unicode_pw_open.xls` | XOR (legacy obfuscation) | BIFF8 | same as above | `pässwörd` |
| `biff8_xor_pw_open_long_password.xls` | XOR (legacy obfuscation) | BIFF8 | same as above | `0123456789abcdef` (effective: `0123456789abcde`) |
| `biff8_xor_pw_open_unicode_method2.xls` | XOR (legacy obfuscation) | BIFF8 | same as above | `Ā` |
| `biff8_xor_pw_open_empty_password.xls` | XOR (legacy obfuscation) | BIFF8 | same as above | `""` |
| `biff8_rc4_standard_pw_open.xls` | RC4 “standard” | BIFF8 | same as above | `password` |
| `biff8_rc4_standard_unicode_pw_open.xls` | RC4 “standard” | BIFF8 | same as above | `pässwörd` |
| `biff8_rc4_standard_unicode_emoji_pw_open.xls` | RC4 “standard” | BIFF8 | same as above | `pässwörd🔒` |
| `biff8_rc4_standard_pw_open_long_password.xls` | RC4 “standard” | BIFF8 | generated from `basic.xls` (this repo) | `0123456789abcdef` (effective: `0123456789abcde`) |
| `biff8_rc4_standard_pw_open_empty_password.xls` | RC4 “standard” | BIFF8 | generated from `basic.xls` (this repo) | `""` |
| `biff8_rc4_cryptoapi_pw_open.xls` | RC4 (CryptoAPI) | BIFF8 | same as above; additionally used by `tests/import_encrypted_rc4_cryptoapi.rs` to validate `import_xls_path_with_password` | `correct horse battery staple` |
| `biff8_rc4_cryptoapi_md5_pw_open.xls` | RC4 (CryptoAPI, MD5 `AlgIDHash`) | BIFF8 | same as above; used by `tests/import_encrypted_rc4_cryptoapi_md5.rs` to validate MD5-based CryptoAPI RC4 | `password` |
| `biff8_rc4_cryptoapi_unicode_pw_open.xls` | RC4 (CryptoAPI) | BIFF8 | same as above | `pässwörd` |
| `biff8_rc4_cryptoapi_unicode_emoji_pw_open.xls` | RC4 (CryptoAPI) | BIFF8 | same as above | `pässwörd🔒` |
| `../encrypted_rc4_cryptoapi_boundary.xls` | RC4 (CryptoAPI, legacy FILEPASS layout) | BIFF8 | Microsoft Excel (real file; used by `tests/import_encrypted_rc4_boundary.rs` to exercise legacy CryptoAPI + 1024-byte rekey boundary behavior) | `password` |
| `../encrypted_xor_biff5.xls` | XOR (legacy obfuscation) | BIFF5 | LibreOffice (real file; used by `tests/import_encrypted_xor_biff5.rs`) | `xorpass` |

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
