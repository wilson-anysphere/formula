# Password-protected OOXML workbooks (MS-OFFCRYPTO)

Excel “Encrypt with Password” for `.xlsx` / `.xlsm` files is **not** ZIP-level encryption. Instead,
Excel wraps the real OOXML ZIP package inside an **OLE Compound File Binary Format (CFB)** container
and stores the encryption metadata + ciphertext in two well-known streams.

This document explains what those files look like on disk, what we currently support, and how to
debug user reports about “password-protected workbooks”.

Relevant specs:

- **MS-OFFCRYPTO** (Office document encryption): https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
- **MS-CFB** (Compound File Binary File Format / OLE): https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/

## What encrypted OOXML files look like

An encrypted `.xlsx` / `.xlsm` is an **OLE/CFB** file (magic bytes `D0 CF 11 E0 A1 B1 1A E1`), not a
ZIP file (magic bytes `PK`).

Inside the CFB container, Excel stores:

- `EncryptionInfo` — an MS-OFFCRYPTO header describing the encryption scheme + KDF parameters.
- `EncryptedPackage` — the encrypted bytes of the original OOXML ZIP package.

Notes:

- `EncryptedPackage` begins with an **8-byte little-endian length prefix** (the plaintext package
  size), followed by block-aligned ciphertext.
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

## Supported encryption schemes

MS-OFFCRYPTO defines multiple encryption “containers” for OOXML packages:

- **Standard** encryption (binary headers; CryptoAPI) — `EncryptionInfo` version **3.2**
- **Agile** encryption (XML descriptor) — `EncryptionInfo` version **4.4**

Current support in this repo:

- ✅ **Standard (CryptoAPI / AES)**: supported
- ❌ **Agile**: detected and currently **unsupported**

If you encounter a password-protected workbook from modern Excel, it is very commonly **Agile**
encryption. In that case we will return an “unsupported encryption” error (see below).

## Public API usage

Use the password-aware helpers (in `crates/formula-io`) when opening encrypted OOXML files:

```rust
use formula_io::open_workbook_with_password;

let workbook = open_workbook_with_password(
    "encrypted.xlsx",
    Some("correct horse battery staple"),
)?;
```

If you want a `formula_model::Workbook` directly (streaming, lower-memory):

```rust
use formula_io::open_workbook_model_with_password;

let model = open_workbook_model_with_password(
    "encrypted.xlsx",
    Some("correct horse battery staple"),
)?;
```

Behavior notes:

- Calling `open_workbook(..)` / `open_workbook_model(..)` on an encrypted OOXML file will return an
  `PasswordRequired` error (because those APIs do not prompt for passwords).
- The `_with_password` variants are intended to work for both encrypted and unencrypted inputs; for
  unencrypted workbooks they behave like the non-password variants.

## Error mapping (debugging + user-facing messaging)

When handling user reports, these error variants map cleanly to “what happened”:

| Error | Meaning | Typical remediation |
|------|---------|---------------------|
| `PasswordRequired` | The file is encrypted/password-protected, but no password was provided. | Retry with `open_workbook_with_password(.., Some(password))` / `open_workbook_model_with_password(.., Some(password))`, or ask the user to remove encryption in Excel. |
| `InvalidPassword` | The workbook uses a supported encryption scheme, but the password does not decrypt the package. | Ask the user to re-enter the password; confirm it opens in Excel with the same password. |
| `UnsupportedOoxmlEncryption` | We identified the OOXML encryption *container*, but it is not implemented (most commonly **Agile 4.4**). | Ask the user to remove encryption (open in Excel → remove password → re-save), or provide an unencrypted copy. |
| `EncryptedWorkbook` | Legacy `.xls` BIFF encryption was detected (BIFF `FILEPASS`). | Ask the user to remove encryption in Excel, or convert the workbook to an unencrypted format. |

If you need to distinguish **Agile vs Standard** for triage, include the `EncryptionInfo` version
from the snippet above in the bug report:

- `3.2` → Standard (supported)
- `4.4` → Agile (unsupported)

## Test fixtures and attribution

Encrypted workbook fixtures are stored under:

- `fixtures/encrypted/ooxml/` (`agile.xlsx`, `standard.xlsx`)

These are **vendored test fixtures** used by `crates/formula-io/tests/encrypted_ooxml.rs`.

Attribution / provenance:

- The fixtures are generated via an Apache POI-based generator (pinned jars + SHA256 verification)
  under `tools/encrypted-ooxml-fixtures/`.
- See `fixtures/encrypted/ooxml/README.md` for the regeneration recipe and tool versions.
