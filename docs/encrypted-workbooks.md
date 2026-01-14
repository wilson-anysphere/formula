# Encrypted / Password‑Protected Excel Workbooks

This page describes Formula’s **file-level workbook encryption** (“Password to open”) support:

- what “encrypted workbook” means in this codebase
- what encryption schemes are supported
- how to open encrypted files (Rust libraries + desktop UX flow)
- error behavior + security notes

For full implementation details and deeper spec/debugging notes, see:

- [`docs/21-encrypted-workbooks.md`](./21-encrypted-workbooks.md) — overview + support matrix + API notes + error semantics
- [`docs/21-offcrypto.md`](./21-offcrypto.md) — MS‑OFFCRYPTO primer + scheme detection + `formula-io` password APIs
- [`docs/22-ooxml-encryption.md`](./22-ooxml-encryption.md) — Agile (4.4) OOXML decryption details (HMAC target bytes, IV/salt gotchas)
- [`docs/office-encryption.md`](./office-encryption.md) — maintainer-level reference

---

## What “encrypted workbook” means (in this project)

Excel has multiple “password” features. This page is about **file encryption** (a password is
required to *open/read* the workbook), not worksheet/workbook *protection* (which restricts editing
but does not encrypt the file).

Formula recognizes two common encrypted workbook container shapes:

1. **OOXML encrypted package** (`.xlsx` / `.xlsm` / `.xlsb` “Encrypt with Password”):
   - The file on disk is an **OLE/CFB compound file** with streams:
     - `EncryptionInfo`
     - `EncryptedPackage` (encrypted bytes of the real ZIP/OPC package)
2. **Legacy `.xls` encryption** (BIFF):
   - The workbook stream contains a BIFF `FILEPASS` record (`0x002F`), indicating the BIFF payload
     must be decrypted before parsing.

---

## Supported encryption schemes

### OOXML (`EncryptionInfo` + `EncryptedPackage`)

Supported:

- **Agile** (`EncryptionInfo` version `4.4`)
- **Standard/CryptoAPI** (`EncryptionInfo` with `versionMinor == 2`, commonly `3.2` / `4.2`)

Unsupported encrypted OOXML containers surface as “unsupported encryption” errors (see below).

### Legacy `.xls` (BIFF `FILEPASS`)

Supported (common real-world variants):

- **XOR obfuscation**
- **RC4 “Standard”**
- **RC4 CryptoAPI**

Other/unknown legacy schemes may be detected but rejected as unsupported.

---

## Opening encrypted workbooks

### Rust: `formula-io` (recommended)

Use `formula-io` for format detection + password handling. Prefer the options-based API:

```rust
use formula_io::{open_workbook_with_options, Error, OpenOptions};

let path = "encrypted.xlsx";

match open_workbook_with_options(
    path,
    OpenOptions {
        password: Some("password".to_string()),
        ..Default::default()
    },
) {
    Ok(wb) => {
        let _ = wb;
    }
    Err(Error::PasswordRequired { .. }) => {
        // Prompt the user and retry with a password.
    }
    Err(Error::InvalidPassword { .. }) => {
        // Wrong password (or integrity mismatch for some OOXML schemes).
    }
    Err(err) => return Err(err),
}
```

Notes:

- Encrypted OOXML decryption is behind the `formula-io` cargo feature
  **`encrypted-workbooks`** (enabled by default). If you depend on `formula-io` with
  `default-features = false`, you must enable it explicitly or encrypted OOXML files will surface
  as `UnsupportedEncryption`.
- For “open into model” use cases, prefer `open_workbook_model_with_password(..)` for encrypted
  OOXML containers; see `docs/21-offcrypto.md` for the exact API matrix.
- If you want to prompt before attempting a full open, use
  `formula_io::detect_workbook_encryption(..)`.

Example `Cargo.toml` (when using `default-features = false`):

```toml
# Enable encrypted workbook support (OOXML + legacy `.xls`) explicitly.
formula-io = { version = "*", default-features = false, features = ["encrypted-workbooks"] }
```

### Rust: `formula-xlsx` convenience decryption helpers (advanced)

If you already have encrypted OOXML bytes in memory and want to manually decrypt to plaintext ZIP
bytes, `formula-xlsx` exposes MS‑OFFCRYPTO helpers under `formula_xlsx::offcrypto`:

```rust
let encrypted = std::fs::read("encrypted.xlsx")?;
let decrypted_zip = formula_xlsx::offcrypto::decrypt_ooxml_from_ole_bytes(&encrypted, "password")?;
assert!(decrypted_zip.starts_with(b"PK"));
```

Once decrypted, you can open the ZIP bytes using normal XLSX readers. For
streaming/round-trip preservation, prefer:

- `formula_xlsx::XlsxLazyPackage::from_vec(decrypted_zip)`

If you need full in-memory access to all parts, you can instead materialize the
package via `formula_xlsx::XlsxPackage::from_bytes(&decrypted_zip)`.

Note: encrypted `.xlsb` files also decrypt to a ZIP/OPC package, but the payload contains
`xl/workbook.bin` instead of `xl/workbook.xml`.

- `formula-io`’s password-aware open APIs will decrypt and return `Workbook::Xlsb` for these files.
- If you are working directly with bytes, you can also open them via
  `formula_io::xlsb::XlsbWorkbook::open_from_bytes_with_password(&bytes, password, formula_io::xlsb::OpenOptions::default())`.

### Desktop app UX flow

The expected UX is:

1. Attempt open without a password.
2. If the backend returns **PasswordRequired**, prompt the user for a password (secure input).
3. Retry open with the password.
4. If **InvalidPassword**, show an error and allow retry/cancel.
5. If **UnsupportedEncryption**, explain that the encryption scheme is not supported (suggest
   re-saving without encryption in Excel).

Passwords should not be persisted or logged by default.

Implementation note (desktop IPC):

- The Tauri backend command `open_workbook` accepts an optional `password` parameter.
- For encrypted workbooks (OOXML `EncryptedPackage` and legacy `.xls` `FILEPASS`), the backend surfaces sentinel error strings prefixed with:
  - `PASSWORD_REQUIRED:` (no password provided)
  - `INVALID_PASSWORD:` (wrong password)
  so the frontend can prompt/retry without doing string matching on the rest of the message.

Current desktop limitations:

- Password-aware open is supported for:
  - **Encrypted OOXML workbooks** (`EncryptionInfo` + `EncryptedPackage`, decrypting to XLSX/XLSM/XLSB ZIP packages)
  - **Legacy encrypted `.xls` workbooks** (BIFF `FILEPASS`)
  via the `open_workbook` command’s `password` parameter. The desktop frontend prompts for a
  password and retries open when it receives `PASSWORD_REQUIRED:` / `INVALID_PASSWORD:` errors.
- Encrypted `.xlsb` opens are supported, but because the input on disk is an encrypted OLE/CFB
  container (not a plaintext `.xlsb` ZIP package), the desktop app cannot round-trip `.xlsb` yet and
  forces “Save As”. Save as `.xlsx` instead.
- Saving **encrypted** `.xlsx`/`.xlsm` is supported when the save command is invoked with a
  `password` (the output is wrapped back into an OLE/CFB `EncryptedPackage` container).
- Saving encrypted `.xlsb` is not supported yet (save as `.xlsx` instead).

---

## Error behavior (what callers should handle)

At the `formula-io` layer, callers should branch on the error variant (not string matching):

- `PasswordRequired` — encrypted workbook detected, but no password provided
- `InvalidPassword` — password provided, but decryption failed
- `UnsupportedOoxmlEncryption` / `UnsupportedEncryption` — encrypted workbook detected, but the
  scheme or build configuration doesn’t support decrypting it
- `EncryptedWorkbook` — legacy `.xls` encryption detected (BIFF `FILEPASS`); retry via the password
  APIs (`open_workbook_with_password` / `open_workbook_model_with_password`)

See [`docs/21-offcrypto.md`](./21-offcrypto.md#error-mapping-debugging--user-facing-messaging) for a
full mapping table and remediation suggestions.

---

## Security notes

- Decryption should be **in-memory**; do not write plaintext workbook bytes to disk.
- Passwords are secrets:
  - do not log them
  - do not include them in telemetry/crash reports
- Large encrypted OOXML workbooks decrypt to a full ZIP/OPC package; this can have significant peak
  memory usage. Prefer model/streaming APIs when possible.

## Limitations

- **Not all encryption schemes are supported.** Files using unknown `EncryptionInfo` versions or
  unsupported legacy `.xls` `FILEPASS` variants will return an “unsupported encryption” error.
- **Encryption is not automatically preserved on save.** In general, Formula operates on plaintext
  workbook bytes once opened; callers must opt into re-encryption for the output file.
  - To write a **new encrypted output** workbook, use `save_workbook_with_options(..)` with
    `SaveOptions { password: Some(..), encryption_scheme: SaveEncryptionScheme::Agile | ::Standard }`
    (requires `formula-io/encrypted-workbooks`).
    - Note: encrypted `.xlsb` output currently requires saving a `Workbook::Xlsb` to an `.xlsb` path
      (Formula does not export arbitrary workbooks *to* `.xlsb`; save as `.xlsx` instead).
  - For Office-encrypted OOXML inputs, callers that want to preserve encryption should use
    `open_workbook_with_password_and_preserved_ole(..)` and then
    `OpenedWorkbookWithPreservedOle::save_preserving_encryption(..)` (requires
    `formula-io/encrypted-workbooks`).
    - This path also preserves any additional non-encryption OLE streams/storages (metadata, etc).
    - This re-encrypts using `crates/formula-office-crypto` (currently with Agile defaults). See
    [`docs/21-encrypted-workbooks.md#saving--round-trip-limitations`](./21-encrypted-workbooks.md#saving--round-trip-limitations).
