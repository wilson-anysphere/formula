# Password-protected OOXML workbooks (MS-OFFCRYPTO)

Excel “Encrypt with Password” for `.xlsx` / `.xlsm` / `.xlsb` (and OOXML templates/add-ins like
`.xltx` / `.xltm` / `.xlam`) is **not** ZIP-level encryption. Instead, Excel wraps the real OOXML ZIP
package inside an **OLE Compound File Binary Format (CFB)** container and stores the encryption
metadata + ciphertext in two well-known streams.

This document explains what those files look like on disk and how to debug user reports about
“password-protected workbooks”.

Note: the more complete, up-to-date overview lives in
[`docs/21-encrypted-workbooks.md`](./21-encrypted-workbooks.md) (covers both OOXML + legacy `.xls`).

Important: do not confuse **file encryption** (“Password to open” / “Encrypt with Password”) with
**workbook/worksheet protection** (“Password to edit”). Protection settings live inside the normal
OOXML package (e.g. `xl/workbook.xml` `<workbookProtection>` and `xl/worksheets/sheetN.xml`
`<sheetProtection>`) and the file still starts with `PK`.

For maintainer-level implementation notes (supported parameter subsets, KDF nuances, writer
defaults), see [`docs/office-encryption.md`](./office-encryption.md).

For a deep dive on **Agile (4.4)** OOXML password decryption (including the *exact* `dataIntegrity`
HMAC target bytes and common IV/salt mixups), see
[`docs/22-ooxml-encryption.md`](./22-ooxml-encryption.md).

For a deep dive on **Standard (CryptoAPI)** (`EncryptionInfo` `versionMinor == 2`; commonly `3.2`)
password decryption, see:

- Key derivation + verifier validation: [`docs/offcrypto-standard-cryptoapi.md`](./offcrypto-standard-cryptoapi.md)
- RC4 variant specifics (0x200 block size, vectors): [`docs/offcrypto-standard-cryptoapi-rc4.md`](./offcrypto-standard-cryptoapi-rc4.md)
- `EncryptedPackage` framing + truncation: [`docs/offcrypto-standard-encryptedpackage.md`](./offcrypto-standard-encryptedpackage.md)

Relevant specs:

- **MS-OFFCRYPTO** (Office document encryption): https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
- **MS-CFB** (Compound File Binary File Format / OLE): https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/

## What encrypted OOXML files look like

An encrypted `.xlsx` / `.xlsm` / `.xlsb` is an **OLE/CFB** file (magic bytes
`D0 CF 11 E0 A1 B1 1A E1`), not a ZIP file (magic bytes `PK`).

Inside the CFB container, Excel stores:

- `EncryptionInfo` — an MS-OFFCRYPTO header describing the encryption scheme + KDF parameters.
- `EncryptedPackage` — the encrypted bytes of the original OOXML ZIP package.

Notes:

- `EncryptedPackage` begins with an **8-byte little-endian length prefix** (the plaintext package
  size), followed by ciphertext (block-aligned for AES-based schemes).
- After decrypting `EncryptedPackage`, the plaintext is a normal OOXML ZIP/OPC package:
  - `.xlsx` / `.xlsm` → ZIP containing `xl/workbook.xml`
  - `.xlsb` → ZIP containing `xl/workbook.bin`
- The file extension is still usually `.xlsx`, so callers that assume “xlsx == zip” will surface
  confusing errors like “invalid zip header”.

### Quick on-disk diagnosis (when a user says “xlsx won’t open”)

1. Check the first bytes:
   - `PK` → normal OOXML ZIP.
   - `D0 CF 11 E0 A1 B1 1A E1` → OLE container (either legacy `.xls` **or** encrypted OOXML).

   Example (print the first 8 bytes as hex):

   ```bash
   xxd -l 8 -p workbook.xlsx
   ```

   - `504b0304...` ⇒ ZIP (`PK..`)
   - `d0cf11e0a1b11ae1` ⇒ OLE/CFB
2. If it’s OLE, list streams and look for `EncryptionInfo` + `EncryptedPackage`.

Minimal Rust snippet for a bug report:

```rust
use std::io::Read;

let file = std::fs::File::open("workbook.xlsx")?;
let mut ole = cfb::CompoundFile::open(file)?;

for entry in ole.walk() {
    println!("{}", entry.path().display());
}

let mut info = ole.open_stream("EncryptionInfo")?;
let mut ver = [0u8; 4];
info.read_exact(&mut ver)?;
let major = u16::from_le_bytes([ver[0], ver[1]]);
let minor = u16::from_le_bytes([ver[2], ver[3]]);
println!("EncryptionInfo version: {major}.{minor}");
```

If you just need a lightweight encryption classification (without parsing OLE streams yourself),
`formula-io` exposes `detect_workbook_encryption(path)` which returns `WorkbookEncryption` (e.g.
`OoxmlEncryptedPackage` vs legacy `.xls` `FILEPASS`).

## Supported encryption schemes

MS-OFFCRYPTO defines multiple encryption “containers” for OOXML packages. For Excel-encrypted
spreadsheets you will most commonly see:

- **Standard** encryption (binary headers; CryptoAPI) — `EncryptionInfo` version **3.2** (or sometimes
  `2.2`/`4.2` in the wild; i.e. `versionMinor == 2` with `versionMajor ∈ {2,3,4}`)
- **Agile** encryption (XML descriptor) — `EncryptionInfo` version **4.4**

Current state in this repo (important nuance):

  - Decryption primitives exist in multiple crates:
  - Higher-level decrypt helpers (OLE wrapper → decrypted ZIP bytes) and an OOXML encryption writer
    (Agile by default; Standard/CryptoAPI AES is also supported — Standard RC4 writing is not):
    `crates/formula-office-crypto`
    - Note: `formula-office-crypto`'s Agile (4.4) decrypt path validates `<dataIntegrity>` when
      present (mismatch ⇒ `OfficeCryptoError::IntegrityCheckFailed`). Some real-world producers omit
      `<dataIntegrity>`; in that case decryption can still succeed, but no integrity verification is
      performed (decrypted bytes are unauthenticated).
      - When `<dataIntegrity>` *is* present, `formula-office-crypto` is permissive about which bytes
        are authenticated by the HMAC for compatibility (e.g. ciphertext-only, plaintext-only, or
        header + plaintext); see
        [`docs/22-ooxml-encryption.md`](./22-ooxml-encryption.md).
  - MS-OFFCRYPTO parsing + decrypt helpers (Standard + Agile):
    `crates/formula-offcrypto`
    - Note: Agile `dataIntegrity` verification is optional there (`DecryptOptions.verify_integrity`).
  - Agile (4.4) parsing + decryption + `dataIntegrity` HMAC verification:
    `crates/formula-xlsx::offcrypto`
- The high-level `formula-io` open path always provides **detection + error classification**
  (`PasswordRequired` / `InvalidPassword` / `UnsupportedOoxmlEncryption`) so callers can prompt for a
  password and route errors correctly.
  - Without the `formula-io` cargo feature **`encrypted-workbooks`**, encrypted OOXML containers
    surface `Error::UnsupportedEncryption` (and `Error::UnsupportedOoxmlEncryption` for unknown/invalid
    `EncryptionInfo` versions).
    - Note: `formula-io` enables `encrypted-workbooks` by default. This “without” case only applies if
      you build with `default-features = false` (or otherwise disable the feature) to exclude
      password-based decryption (and its crypto dependencies).
  - With **`encrypted-workbooks`** enabled:
    - `open_workbook(..)` / `open_workbook_model(..)` surface `Error::PasswordRequired` when no
      password is provided.
    - The password-aware `open_workbook_with_password` / `open_workbook_model_with_password` can
      decrypt and open **Agile (4.4)** and **Standard/CryptoAPI** (minor=2; major ∈ {2,3,4}) encrypted
      `.xlsx`/`.xlsm`/`.xlsb` workbooks in memory.
      - For Agile, `dataIntegrity` (HMAC) is validated when present; some real-world producers omit
        it.
    - `open_workbook_with_options` can also decrypt and open encrypted OOXML wrappers when a password
      is provided (returns `Workbook::Xlsx` / `Workbook::Xlsb` depending on the decrypted payload).
  - `open_workbook_model_with_options` can also decrypt encrypted OOXML wrappers when
    `formula-io/encrypted-workbooks` is enabled (and surfaces `PasswordRequired` when
    `OpenOptions.password` is `None`). Without that feature, encrypted OOXML containers surface
    `UnsupportedEncryption`. `open_workbook_model_with_password` is a convenience wrapper around it.
  - A streaming decrypt reader exists in `crates/formula-io/src/encrypted_ooxml.rs` +
    `crates/formula-io/src/encrypted_package_reader.rs`.
    - This is used for some compatibility fallbacks (for example Agile files that omit
      `<dataIntegrity>`) and to open Standard/CryptoAPI AES-encrypted `.xlsx`/`.xlsm` into a model
      without materializing the decrypted ZIP bytes (`open_workbook_model_with_password` /
      `open_workbook_model_with_options`).
    - The workbook-returning open APIs (`open_workbook_with_options` / `open_workbook_with_password`)
      still decrypt `EncryptedPackage` into an in-memory buffer first.

When triaging user reports, the most important thing is to capture the `EncryptionInfo` version
because it determines which scheme you’re dealing with:

- `4.4` ⇒ Agile
- `minor == 2` with `major ∈ {2,3,4}` ⇒ Standard/CryptoAPI (often seen as `3.2` or `4.2`)

## Public API usage

Use the password-aware helpers (in `crates/formula-io`) when handling encrypted OOXML files.

Preferred API: `open_workbook_with_options` + `OpenOptions.password`:

```rust
use formula_io::{open_workbook_with_options, OpenOptions};

let workbook = open_workbook_with_options(
    "book.xlsx", // may be encrypted
    OpenOptions {
        password: Some("correct horse battery staple".to_string()),
        ..Default::default()
    },
)?;
```

Interactive flow (prompt when needed):

```rust
use formula_io::{open_workbook, open_workbook_with_options, Error, OpenOptions};

let path = "book.xlsx"; // may be encrypted

match open_workbook(path) {
    Ok(workbook) => {
        // Unencrypted workbook opened normally.
        let _ = workbook;
    }
    Err(Error::PasswordRequired { .. }) => {
        // Encrypted OOXML container detected (requires `formula-io/encrypted-workbooks`).
        // Prompt user for password and retry:
        let password = "correct horse battery staple";
        let _ = open_workbook_with_options(
            path,
            OpenOptions {
                password: Some(password.to_string()),
                ..Default::default()
            },
        )?;
    }
    Err(Error::UnsupportedEncryption { .. }) => {
        // Encrypted workbook detected, but this build doesn't support decrypting it (for example
        // `formula-io/encrypted-workbooks` is not enabled).
    }
    Err(other) => return Err(other),
}
```

If you want a `formula_model::Workbook` directly:

```rust
use formula_io::{open_workbook_model_with_password, Error};

match open_workbook_model_with_password("book.xlsx", Some("...")) {
    Ok(model) => {
        let _ = model;
    }
    Err(Error::PasswordRequired { .. }) => {
        // Prompt user for password and retry.
    }
    Err(other) => return Err(other),
}
```

Convenience helpers:

- `open_workbook_with_password(path, Some(password))` is a thin wrapper around `OpenOptions.password`.
- `open_workbook_model_with_password(..)` opens directly into a `formula_model::Workbook`.

Behavior notes:

- Calling `open_workbook(..)` / `open_workbook_model(..)` on an encrypted OOXML file will return:
  - `PasswordRequired` when `formula-io/encrypted-workbooks` is enabled (for supported `EncryptionInfo`
    versions),
  - `UnsupportedEncryption` when it is not, and
  - `UnsupportedOoxmlEncryption` when the `EncryptionInfo` version is unknown/unsupported.
- The `_with_password` variants are intended to work for both encrypted and unencrypted inputs; for
  unencrypted workbooks they behave like the non-password variants.
  - With `formula-io/encrypted-workbooks` enabled:
  - The legacy `_with_password` variants can open Agile (4.4) and Standard/CryptoAPI (minor=2)
    encrypted `.xlsx`/`.xlsm`/`.xlsb`.
  - `open_workbook_with_options` can also decrypt and open encrypted OOXML wrappers when a password
    is provided.
  - `open_workbook_model_with_options` can also decrypt encrypted OOXML wrappers when `password` is
    provided (or surface `PasswordRequired` when missing). `open_workbook_model_with_password` is the
    convenience wrapper.
  - Decryption happens in memory; do not write decrypted bytes to disk. See
    [`docs/21-encrypted-workbooks.md#security-notes-handling-decrypted-bytes-safely`](./21-encrypted-workbooks.md#security-notes-handling-decrypted-bytes-safely).
- Without `formula-io/encrypted-workbooks`, encrypted OOXML containers surface
  `Error::UnsupportedEncryption` (the password-aware entrypoints do not decrypt them end-to-end).

## Saving encrypted workbooks (optional)

With the `formula-io/encrypted-workbooks` feature enabled, Formula can also **write**
password-protected OOXML workbooks (an OLE/CFB wrapper containing `EncryptionInfo` +
`EncryptedPackage`).

- To write a **new encrypted output** workbook, use `save_workbook_with_options(..)` with
  `SaveOptions { password: Some(..), encryption_scheme: SaveEncryptionScheme::Agile | ::Standard }`.
- To preserve **extra non-encryption OLE streams/storages** (metadata, etc) from an *encrypted input*
  workbook, use `open_workbook_with_password_and_preserved_ole(..)` and then
  `OpenedWorkbookWithPreservedOle::save_preserving_encryption(..)`.

See [`docs/21-encrypted-workbooks.md#saving--round-trip-limitations`](./21-encrypted-workbooks.md#saving--round-trip-limitations)
for details and limitations.

## Error mapping (debugging + user-facing messaging)

When handling user reports, these error variants map cleanly to “what happened”:

| Error | Meaning | Typical remediation |
|------|---------|---------------------|
| `PasswordRequired` | The file is encrypted/password-protected, but no password was provided. (For encrypted OOXML containers, this is only returned when `formula-io/encrypted-workbooks` is enabled; otherwise they surface as `UnsupportedEncryption`.) | Retry with `open_workbook_with_options(.., OpenOptions { password: Some(..) })` (or `open_workbook_with_password(.., Some(password))`), or ask the user to remove encryption in Excel. |
| `InvalidPassword` | Password was provided but decryption/verification failed (wrong password, Agile `dataIntegrity` mismatch, unsupported encryption parameters, or a corrupt container). | Ask the user to re-enter the password and confirm it opens in Excel. If you control the build, ensure `formula-io/encrypted-workbooks` is enabled. If Formula still cannot open it, capture `EncryptionInfo` version + the exact error, then ask the user to remove encryption and re-save. |
| `UnsupportedOoxmlEncryption` | We identified an encrypted OOXML container, but the `EncryptionInfo` version is not recognized/implemented (i.e. not Standard/CryptoAPI `minor == 2` or Agile `4.4`). | Ask the user to remove encryption (open in Excel → remove password → re-save), or provide an unencrypted copy. |
| `EncryptedWorkbook` | Legacy `.xls` BIFF encryption was detected (BIFF `FILEPASS`). | Retry with `open_workbook_with_password(.., Some(password))` / `open_workbook_model_with_password(.., Some(password))` (routes to the `.xls` importer for legacy `.xls`). If that still fails, the encryption scheme may be unsupported; ask the user to remove encryption in Excel or provide an unencrypted copy. |
| `UnsupportedEncryption` | The workbook is encrypted, but this build/decryptor doesn’t support decrypting it (most commonly: `formula-io/encrypted-workbooks` is disabled for encrypted OOXML, or an unsupported legacy `.xls` encryption scheme). | Capture the `EncryptionInfo` version and error string; consider enabling `formula-io/encrypted-workbooks` (OOXML Agile 4.4 + Standard/CryptoAPI minor=2) or ask the user to remove encryption in Excel and re-save. |

If you need to distinguish **Agile vs Standard** for triage, include the `EncryptionInfo`
`major.minor` pair in the bug report (or at least whether `minor == 2` vs `4.4`).

### Lower-level error mapping (`formula-office-crypto`)

If you are debugging decryption directly (outside the `formula-io` open path), the helper crate
`crates/formula-office-crypto` exposes:

- `formula_office_crypto::decrypt_encrypted_package_ole(bytes, password)` → decrypted ZIP bytes
- `OfficeCryptoError::{InvalidPassword, UnsupportedEncryption, InvalidFormat, Io, ...}`

This is where you will most commonly see the **`UnsupportedEncryption`** error: it indicates we
recognized the encrypted OOXML wrapper, but the specific cipher/hash/KDF settings are not implemented
by the decryptor.

### Manual decryption (debug-only)

If you want to debug a user report *outside* the `formula-io` open path, you can decrypt the OOXML
package bytes directly and then inspect the decrypted ZIP contents:

```rust
use std::io::Cursor;

let encrypted = std::fs::read("encrypted.xlsx")?;
let decrypted_zip = formula_office_crypto::decrypt_encrypted_package_ole(&encrypted, "password")?;
assert!(decrypted_zip.starts_with(b"PK"));

let zip = zip::ZipArchive::new(Cursor::new(&decrypted_zip))?;
for name in zip.file_names() {
    println!("{name}");
}
```

If the decrypted bytes do **not** start with `PK`, treat that as either:

- wrong password, or
- unsupported/corrupt encryption wrapper (the decryptor should ideally surface `InvalidPassword` /
  `UnsupportedEncryption` / `InvalidFormat` in these cases).

Note: `formula-office-crypto` validates Agile `<dataIntegrity>` when it is present, but does not
require it. Some real-world producers omit `<dataIntegrity>`; in that case, decryption can still
succeed, but integrity cannot be verified (the decrypted bytes are unauthenticated).
When `<dataIntegrity>` is present, `formula-office-crypto` is permissive about HMAC target bytes for
compatibility (e.g. ciphertext-only, plaintext-only, or header + plaintext); see
[`docs/22-ooxml-encryption.md`](./22-ooxml-encryption.md).

### Inspecting Agile `EncryptionInfo` XML (debug-only)

For **Agile** encryption (`EncryptionInfo` version `4.4`), the remainder of the `EncryptionInfo`
stream is typically an XML descriptor. If you need to capture the exact cipher/KDF parameters for a
bug report, `formula-io` has a helper that extracts and validates that XML while handling common
real-world encodings (UTF-8 vs UTF-16LE, optional length prefixes, padding):

```rust
use std::io::Read;

let mut ole = cfb::CompoundFile::open(std::fs::File::open("encrypted.xlsx")?)?;
let mut encryption_info = Vec::new();
ole.open_stream("EncryptionInfo")?
    .read_to_end(&mut encryption_info)?;

let xml = formula_io::extract_agile_encryption_info_xml(&encryption_info)?;
println!("{xml}");
```

This is safe to log from a password perspective (it does **not** include the password), but it does
include salts/IV material and other encryption metadata, so treat it as sensitive file content in
privacy-sensitive environments.

## Test fixtures and attribution

Encrypted workbook fixtures are stored under:

- `fixtures/encrypted/` (see `fixtures/encrypted/README.md`; includes real-world encrypted `.xlsx`/`.xlsb`/`.xls` used by end-to-end tests like `crates/formula-io/tests/open_encrypted_fixtures.rs`)
- `fixtures/encrypted/ooxml/` (see `fixtures/encrypted/ooxml/README.md` for the full list; includes
  Agile + Standard encrypted fixtures, empty-password and Unicode-password samples, macro-enabled
  `.xlsm` samples, and `*-large.xlsx` multi-segment variants)

### Fixture passwords (quick reference)

`fixtures/encrypted/ooxml/README.md` is the **canonical** source of truth. At time of writing:

- `agile.xlsx`, `standard.xlsx`, `standard-4.2.xlsx`, `standard-rc4.xlsx`, `agile-large.xlsx`,
  `standard-large.xlsx`, `agile-basic.xlsm`, `standard-basic.xlsm`, `basic-password.xlsm`: `password`
- `agile-empty-password.xlsx`: empty string (`""`) (distinct from a *missing* password)
- `agile-unicode.xlsx`: `pässwörd` (Unicode, NFC form)
- `agile-unicode-excel.xlsx`: `pässwörd🔒` (Unicode, NFC form, includes non-BMP emoji)
- `standard-unicode.xlsx`: `pässwörd🔒` (Unicode, NFC form, includes non-BMP emoji)

These are **vendored test fixtures** used by multiple encryption-focused tests (for example under
`crates/formula-io/tests/*encrypted_ooxml*` and `crates/formula-xlsx/tests/encrypted_ooxml_*.rs`).
See `fixtures/encrypted/ooxml/README.md` for the canonical list, passwords, and where each fixture
is referenced.

Attribution / provenance:

- Most committed fixture bytes were generated using Python +
  [`msoffcrypto-tool`](https://github.com/nolze/msoffcrypto-tool), but some fixtures are generated
  via Apache POI (for example `standard-4.2.xlsx`, `standard-unicode.xlsx`). See
  [`fixtures/encrypted/ooxml/README.md`](../fixtures/encrypted/ooxml/README.md) for the canonical
  per-fixture provenance + the exact tool versions/passwords.
- Alternative regeneration tooling also exists under `tools/encrypted-ooxml-fixtures/` (Apache POI)
  for creating new encrypted OOXML wrappers without Excel.

### Useful debugging tool

For quick triage, `formula-io` includes a small helper binary that prints a one-line summary of an
encrypted OOXML file’s `EncryptionInfo` header:

```bash
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- path/to/encrypted.xlsx
```

This prints a string like `Agile (4.4) flags=0x...` or `Standard (3.2) flags=0x...` (but the Standard
major version can vary: `2.2`/`3.2`/`4.2`), which is often enough to route a bug report to the right
implementation path.

When collecting a bug report, include:

- output of `ooxml-encryption-info` (scheme + version)
- file extension (`.xlsx`/`.xlsm`/`.xlsb`) and whether it opens in Excel
- the exact `formula-io` error variant + message (`PasswordRequired` vs `InvalidPassword` vs
  `UnsupportedOoxmlEncryption`)
