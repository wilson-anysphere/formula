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
  size), followed by block-aligned ciphertext.
- After decrypting `EncryptedPackage`, the plaintext is a normal OOXML ZIP/OPC package:
  - `.xlsx` / `.xlsm` → ZIP containing `xl/workbook.xml`
  - `.xlsb` → ZIP containing `xl/workbook.bin`
- The file extension is still usually `.xlsx`, so callers that assume “xlsx == zip” will surface
  confusing errors like “invalid zip header”.

### Quick on-disk diagnosis (when a user says “xlsx won’t open”)

1. Check the first bytes:
   - `PK` → normal OOXML ZIP.
   - `D0 CF 11 E0 A1 B1 1A E1` → OLE container (either legacy `.xls` **or** encrypted OOXML).
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

If you just need a yes/no classification (without parsing OLE streams yourself), `formula-io` also
exposes `detect_workbook_encryption(path)` which returns `WorkbookEncryptionKind` (e.g.
`OoxmlOleEncryptedPackage` vs legacy `.xls` `FILEPASS`).

## Supported encryption schemes

MS-OFFCRYPTO defines multiple encryption “containers” for OOXML packages. For Excel-encrypted
spreadsheets you will most commonly see:

- **Standard** encryption (binary headers; CryptoAPI) — `EncryptionInfo` version **3.2** (or sometimes
  `2.2`/`4.2` in the wild; i.e. `versionMinor == 2` with `versionMajor ∈ {2,3,4}`)
- **Agile** encryption (XML descriptor) — `EncryptionInfo` version **4.4**

Current state in this repo (important nuance):

- Decryption primitives exist in multiple crates:
  - Higher-level decrypt helpers (OLE wrapper → decrypted ZIP bytes) and an Agile encryption writer:
    `crates/formula-office-crypto`
  - Standard/CryptoAPI parsing + password key derivation + `EncryptedPackage` helpers:
    `crates/formula-offcrypto`
  - Agile (4.4) parsing + decryption + `dataIntegrity` HMAC verification:
    `crates/formula-xlsx::offcrypto`
- The high-level `formula-io` open path does **not** yet decrypt OOXML workbooks end-to-end; it
  primarily provides **detection + error classification** so callers can prompt for a password and
  route errors correctly.

When triaging user reports, the most important thing is to capture the `EncryptionInfo` version
because it determines which scheme you’re dealing with:

- `4.4` ⇒ Agile
- `minor == 2` with `major ∈ {2,3,4}` ⇒ Standard/CryptoAPI (often seen as `3.2` or `4.2`)

## Public API usage

Use the password-aware helpers (in `crates/formula-io`) when handling encrypted OOXML files:

```rust
use formula_io::{open_workbook_with_password, Error};

let path = "book.xlsx"; // may be encrypted

match open_workbook_with_password(path, None) {
    Ok(workbook) => {
        // Unencrypted workbook opened normally.
        let _ = workbook;
    }
    Err(Error::PasswordRequired { .. }) => {
        // Encrypted OOXML container detected; prompt user for password and retry:
        let password = "correct horse battery staple";
        match open_workbook_with_password(path, Some(password)) {
            Ok(workbook) => {
                // Once OOXML decryption is wired end-to-end in `formula-io`, this will return `Ok(...)`.
                let _ = workbook;
            }
            Err(Error::InvalidPassword { .. }) => {
                // Wrong password (or decryption not implemented yet).
            }
            Err(other) => return Err(other),
        }
    }
    Err(other) => return Err(other),
}
```

If you want a `formula_model::Workbook` directly (streaming, lower-memory):

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

Behavior notes:

- Calling `open_workbook(..)` / `open_workbook_model(..)` on an encrypted OOXML file will return an
  `PasswordRequired` error (because those APIs do not prompt for passwords).
- The `_with_password` variants are intended to work for both encrypted and unencrypted inputs; for
  unencrypted workbooks they behave like the non-password variants.
- Until OOXML decryption is wired end-to-end in `formula-io`, the `_with_password` variants will not
  successfully open encrypted OOXML workbooks (they currently surface `InvalidPassword` when a
  password is provided).

## Error mapping (debugging + user-facing messaging)

When handling user reports, these error variants map cleanly to “what happened”:

| Error | Meaning | Typical remediation |
|------|---------|---------------------|
| `PasswordRequired` | The file is encrypted/password-protected, but no password was provided. | Retry with `open_workbook_with_password(.., Some(password))` / `open_workbook_model_with_password(.., Some(password))`, or ask the user to remove encryption in Excel. |
| `InvalidPassword` | Password was provided but the workbook could not be opened/decrypted. (This can mean “wrong password”; until decryption is wired end-to-end, it can also mean “not implemented yet”.) | Ask the user to re-enter the password; confirm it opens in Excel with the same password. If Formula still cannot open it, capture `EncryptionInfo` version + the exact error, then ask the user to remove encryption and re-save. |
| `UnsupportedOoxmlEncryption` | We identified an encrypted OOXML container, but the `EncryptionInfo` version is not recognized/implemented (i.e. not Standard/CryptoAPI `minor == 2` or Agile `4.4`). | Ask the user to remove encryption (open in Excel → remove password → re-save), or provide an unencrypted copy. |
| `EncryptedWorkbook` | Legacy `.xls` BIFF encryption was detected (BIFF `FILEPASS`). | Ask the user to remove encryption in Excel, or convert the workbook to an unencrypted format. |
| `UnsupportedEncryption` | (Lower-level decryptors) The encrypted OOXML wrapper was recognized, but the cipher/KDF parameters are unsupported by the decryptor. | Capture the `EncryptionInfo` version and decryptor error string; ask the user to remove encryption in Excel and re-save as a fallback. |

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

- `fixtures/encrypted/ooxml/` (see `fixtures/encrypted/ooxml/README.md` for the full list; includes
  Agile + Standard encrypted fixtures, an empty-password sample, and `*-large.xlsx` multi-segment variants)

These are **vendored test fixtures** used by `crates/formula-io/tests/encrypted_ooxml.rs`.

Attribution / provenance:

- The committed fixture bytes were generated using Python +
  [`msoffcrypto-tool`](https://github.com/nolze/msoffcrypto-tool) (see
  [`fixtures/encrypted/ooxml/README.md`](../fixtures/encrypted/ooxml/README.md) for the exact tool
  versions + passwords).
- Alternative regeneration tooling also exists under `tools/encrypted-ooxml-fixtures/` (Apache POI),
  but it is not used for the committed fixture bytes.

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
