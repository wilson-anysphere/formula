# Encrypted / Password-Protected Excel Workbooks

This document covers **file encryption** (a password is required to *open* the file) and how it
differs from Excel’s **workbook/worksheet protection** features (a password is required to *edit*
certain parts of the workbook, but the file contents are not encrypted).

If you just need a short overview and entrypoints, see:

- [`docs/encrypted-workbooks.md`](./encrypted-workbooks.md)

Formula’s goal is to open encrypted spreadsheets when possible, surface **actionable** errors when
not, and avoid security pitfalls (like accidentally persisting decrypted bytes to disk).

## Related docs

- [`docs/office-encryption.md`](./office-encryption.md) — maintainer-level reference (supported
  parameter subsets, KDF nuances, writer defaults).
- [`docs/21-offcrypto.md`](./21-offcrypto.md) — short MS-OFFCRYPTO overview focused on “what the file
  looks like”, scheme detection, and `formula-io` password APIs.
- [`docs/22-ooxml-encryption.md`](./22-ooxml-encryption.md) — Agile (4.4) OOXML decryption details
  (HMAC target bytes, IV/salt gotchas).
- [`docs/offcrypto-standard-encryptedpackage.md`](./offcrypto-standard-encryptedpackage.md) —
  Standard/CryptoAPI AES `EncryptedPackage` decryption notes (AES-ECB framing + truncation).

## Status (current behavior vs intended behavior)

**Current behavior (in this repo today):**

- Encrypted workbooks are **detected** and surfaced with dedicated errors so the caller/UI can do the
  right thing (prompt for password vs show “unsupported encryption”).
- Encrypted **OOXML** (`EncryptionInfo` + `EncryptedPackage`) yields:
  - `formula_io::Error::UnsupportedOoxmlEncryption` for unknown/unimplemented `EncryptionInfo`
    versions.
  - Without the `formula-io` cargo feature **`encrypted-workbooks`**, Office-encrypted OOXML
    workbooks are treated as unsupported and surface `formula_io::Error::UnsupportedEncryption`
    (Formula does not attempt to decrypt them).
    - Note: `formula-io` enables `encrypted-workbooks` by default. This “without” case only applies if
      you build with `default-features = false` (or otherwise disable the feature) to exclude
      password-based decryption (and its crypto dependencies).
  - With **`encrypted-workbooks`** enabled:
    - `open_workbook(..)` / `open_workbook_model(..)` surface `formula_io::Error::PasswordRequired`
      when no password is provided.
    - Use `open_workbook_with_options(OpenOptions { password: ... })` (or the `_with_password`
      helpers) to decrypt and open supported encrypted workbooks in memory:
      - **Agile (4.4)** encrypted `.xlsx`/`.xlsm`/`.xlsb` workbooks (validates Agile `dataIntegrity`
        (HMAC) when present; some producers omit it).
        - Wrong password *or* integrity mismatch surfaces as `formula_io::Error::InvalidPassword`.
      - **Standard/CryptoAPI** (`versionMinor == 2` with `versionMajor ∈ {2,3,4}`; commonly `3.2`/`4.2`)
        encrypted `.xlsx`/`.xlsm`/`.xlsb` workbooks (wrong password surfaces as
        `formula_io::Error::InvalidPassword`).
      - Note: encrypted `.xlsb` workbooks decrypt to an OOXML ZIP containing `xl/workbook.bin` and
        are opened as `Workbook::Xlsb` (or converted to a model workbook via the password APIs).
        - If you need to open an encrypted `.xlsb` from bytes directly, use
          `formula_io::xlsb::XlsbWorkbook::open_from_bytes_with_password(..)`.
- Legacy **`.xls`** with BIFF `FILEPASS` yields:
  - `formula_io::Error::EncryptedWorkbook` via `open_workbook(..)` / `open_workbook_model(..)` (prompt
    callers to retry via the password-capable APIs).
  - `formula_io::Error::PasswordRequired` / `formula_io::Error::InvalidPassword` via
    `open_workbook_with_password(..)` / `open_workbook_model_with_password(..)` when the password is
    missing/incorrect.
  - When a password is provided, Formula will attempt to decrypt common BIFF `FILEPASS` schemes via
    the `.xls` importer (XOR, RC4 “standard”, and RC4 CryptoAPI; see below).
- The desktop app (with `formula-io/encrypted-workbooks` enabled) surfaces a **password required**
  style error (so the UI can prompt for a password).
- The web/WASM engine can decrypt and open Office-encrypted OOXML `.xlsx`/`.xlsb` bytes in-memory:
  - WASM entrypoint: `crates/formula-wasm::WasmWorkbook::fromEncryptedXlsxBytes(bytes, password)`
  - Worker API: `packages/engine` exposes `EngineClient.loadWorkbookFromEncryptedXlsxBytes(bytes, password)`
  - Decrypted packages are routed based on the inner workbook part:
    - `xl/workbook.xml` → `.xlsx` / `.xlsm`
    - `xl/workbook.bin` → `.xlsb`
  - Implementation note: this path uses `crates/formula-office-crypto` for decryption. For Agile
    encryption, `formula-office-crypto` validates the `<dataIntegrity>` HMAC when present; when
    `<dataIntegrity>` is missing, decryption still works but the integrity check is skipped.

**Intended behavior / remaining work:**

- Continue to support opening Excel-encrypted workbooks without writing decrypted bytes to disk.
- Distinguish “password required” vs “invalid password” vs “unsupported encryption scheme”
  (see [Error semantics](#error-semantics)).

### Support matrix (current vs planned)

| File type | Encryption marker | Schemes (common) | Current behavior | Planned/target behavior |
|---|---|---|---|---|
| `.xlsx` / `.xlsm` / `.xlsb` (OOXML) | OLE/CFB streams `EncryptionInfo` + `EncryptedPackage` | Agile (4.4), Standard (minor=2; e.g. 3.2/4.2) | With `formula-io/encrypted-workbooks`: decrypt + open in memory via `open_workbook_with_options` / `_with_password` for `.xlsx`/`.xlsm`/`.xlsb`, surfacing `PasswordRequired` / `InvalidPassword` / `UnsupportedOoxmlEncryption`. Without that feature: `UnsupportedEncryption` for Office-encrypted OOXML. | Expand scheme/parameter coverage; consider streaming decrypt + tighter resource limits; optionally tighten integrity verification defaults. |
| `.xls` (BIFF) | BIFF `FILEPASS` record in workbook stream | XOR, RC4, CryptoAPI | `formula-io`: `EncryptedWorkbook` when no password is provided. `open_workbook_with_options` / `_with_password` route to `formula-xls`’s decrypting importer (supports XOR, RC4 “standard”, and RC4 CryptoAPI). | Expand scheme coverage as needed (see [Legacy `.xls` encryption](#legacy-xls-encryption-biff-filepass)) |

---

## Terminology: protection vs encryption (do not confuse these)

Excel has multiple “password” features that behave very differently:

### Workbook / worksheet protection (hash-based, not encryption)

These features gate *editing UI actions*, not *confidentiality*:

- **Worksheet protection** (OOXML: `xl/worksheets/sheetN.xml` → `<sheetProtection …/>`)
  - Prevents editing cells, formatting, inserting rows, etc.
- **Workbook protection** (OOXML: `xl/workbook.xml` → `<workbookProtection …/>`)
  - Prevents structural changes like adding/removing sheets or resizing windows.

Key properties:

- The workbook can be opened and read without a password.
- Password material is stored as a **hash/obfuscation** (not strong crypto).
- These mechanisms should be treated as **compatibility/UI controls**, not security boundaries.

Formula generally treats protection parts/records as “normal workbook content”:
we parse and preserve them for round-trip, but we do not treat them as encryption.

### File encryption (password required to open)

This is the Excel feature often labeled **“Encrypt with Password”** / **“Password to open”**:

- The entire workbook payload is encrypted.
- Opening the file requires decrypting bytes *before* any workbook parsing can occur.

This document focuses on **file encryption**.

---

## OOXML encryption (`.xlsx` / `.xlsm` / `.xlsb` / templates / add-ins)

### Container shape: an OLE/CFB wrapper around the encrypted package

Modern password-encrypted OOXML workbooks are **not ZIP files on disk**, even if the extension is
`.xlsx`/`.xlsm`/`.xlsb`.

Instead, Excel writes an **OLE Compound File Binary** (CFB, “Structured Storage”) container whose
top-level streams include:

- `EncryptionInfo` — parameters describing the encryption scheme and key-derivation details
- `EncryptedPackage` — the encrypted bytes of the *real* workbook package
  - For `.xlsx`/`.xlsm`: the encrypted payload decrypts to a normal OPC ZIP.
  - For `.xlsb`: the encrypted payload decrypts to a normal OPC ZIP containing `xl/workbook.bin`.

Detection heuristics:

- File begins with the OLE magic header `D0 CF 11 E0 A1 B1 1A E1`.
- The compound file contains both `EncryptionInfo` and `EncryptedPackage` streams.

### `EncryptionInfo` versions (Standard vs Agile)

The `EncryptionInfo` stream begins with a small fixed header:

- `majorVersion: u16` (little-endian)
- `minorVersion: u16` (little-endian)
- `flags: u32` (little-endian)

The `(majorVersion, minorVersion)` pair determines how the remainder of the stream is interpreted.
In practice:

| Scheme | `major.minor` | Notes |
|--------|---------------|------|
| **Standard** | `*.2` (`versionMinor == 2`, `versionMajor ∈ {2,3,4}` in the wild; commonly `3.2`/`4.2`) | CryptoAPI-style header/verifier structures (binary). |
| **Agile** | `4.4` | XML-based encryption descriptor. |

Implementation notes:

- `crates/formula-io/src/bin/ooxml-encryption-info.rs` prints a one-line scheme/version summary
  based on the `(majorVersion, minorVersion)` header (e.g. `2.2`/`3.2`/`4.2` → Standard, `4.4` → Agile).
  Run it with:

  ```bash
  bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- path/to/encrypted.xlsx
  ```
- `crates/formula-office-crypto` is the primary end-to-end MS-OFFCRYPTO implementation in this repo:
  - parses `EncryptionInfo` for Agile (4.4) and Standard/CryptoAPI (`versionMinor == 2`), and
  - decrypts `EncryptedPackage` to the plaintext OOXML ZIP bytes (and also includes an OOXML
    encryption writer: Agile by default; Standard/CryptoAPI AES is also supported).
- `crates/formula-offcrypto` provides MS-OFFCRYPTO parsing helpers and standalone decrypt primitives
  (useful for inspection tooling, and for some Standard/CryptoAPI helper APIs).
- Standard/CryptoAPI `EncryptedPackage` decryption differs by cipher:
  - **AES** (ECMA-376 baseline) uses **AES-ECB** (no IV), with plaintext truncated to the 8-byte
    plaintext size prefix.
    - See `crates/formula-offcrypto` (`decrypt_standard_ooxml_from_bytes` /
      `decrypt_encrypted_package_ecb`) and `docs/offcrypto-standard-encryptedpackage.md`.
  - **RC4** (`CALG_RC4`) uses 0x200-byte blocks with per-block keys; see
    `docs/offcrypto-standard-cryptoapi-rc4.md`.
- For Agile (4.4) decryption details (HMAC target bytes + IV/salt usage gotchas), see
  [`docs/22-ooxml-encryption.md`](./22-ooxml-encryption.md).

### Supported OOXML encryption schemes

The `EncryptionInfo` stream encodes one of the schemes defined in **[MS-OFFCRYPTO]**. In practice,
Excel-produced encrypted workbooks primarily use:

- **Agile Encryption** (modern; Office 2010+)
- **Standard Encryption** (older; Office 2007 era)

Formula’s encrypted-workbook support targets these two schemes:

- **Agile** (`EncryptionInfo` version 4.4; XML-based descriptor inside `EncryptionInfo`)
- **Standard** (`EncryptionInfo` `versionMinor == 2`; CryptoAPI-style header/verifier)

Everything else should fail with a specific “unsupported encryption scheme” error (see
[Error semantics](#error-semantics)).

### `EncryptedPackage` layout (OOXML)

For both Standard and Agile encryption, the OLE stream `EncryptedPackage` begins with:

- **8 bytes**: `original_package_size` (8-byte plaintext size prefix; see note below)
- remaining bytes: encrypted OPC package data (ciphertext + padding for block ciphers)

Compatibility note: while the spec describes this prefix as a `u64le`, some producers/libraries treat
it as `u32 totalSize` + `u32 reserved` (often 0). For compatibility, parse it as two little-endian
DWORDs and recombine: `size = lo as u64 | ((hi as u64) << 32)`.

After decrypting the ciphertext, the plaintext bytes should be truncated to `original_package_size`
to recover the real workbook package (a normal ZIP/OPC archive for `.xlsx`/`.xlsm`, or a ZIP/OPC
archive containing `xl/workbook.bin` for `.xlsb`).

The encryption *mode* differs by scheme:

- **Standard (CryptoAPI; `versionMinor == 2`):** the cipher is specified by `EncryptionHeader.algId`:
  - **AES** (`CALG_AES_128`/`CALG_AES_192`/`CALG_AES_256`): decrypt the ciphertext bytes (after the
    8-byte size prefix) with **AES-ECB** (no IV). The ciphertext is block-aligned (`len % 16 == 0`);
    after decrypting, **truncate the plaintext to `original_package_size`** (see
    `docs/offcrypto-standard-encryptedpackage.md`). Excel-default Standard AES has **no per-segment
    IV** (ECB). Some third-party producers use non-standard **AES-CBC** variants; our decryptors may
    attempt those as a fallback (see `docs/offcrypto-standard-encryptedpackage.md`).
  - **RC4** (`CALG_RC4`): decrypt the ciphertext as an RC4 stream in **0x200-byte blocks** with
    per-block keys derived from the password hash (no padding/block alignment). See
    `docs/offcrypto-standard-cryptoapi-rc4.md`.
- **Agile (4.4):** encrypted in **4096-byte plaintext segments** with a per-segment IV derived from
  `keyData/@saltValue` and the segment index, and cipher/chaining parameters specified by the XML
  descriptor.

### High-level decrypt/open algorithm (OOXML)

At a high level, opening a password-encrypted OOXML workbook is:

1. **Detect the OLE wrapper**
   - Open the file as a CFB container.
   - Confirm `EncryptionInfo` + `EncryptedPackage` streams exist.
2. **Read and classify `EncryptionInfo`**
   - Parse the `(major, minor)` version header.
   - **Standard (minor=2; major ∈ {2,3,4}):** parse the binary CryptoAPI header + verifier structures.
   - **Agile (4.4):** parse the XML `<encryption>` descriptor (key derivation params, ciphers,
     integrity metadata, …).
3. **Derive keys from the password**
   - Convert the password to UTF-16LE as required by the spec (no BOM/terminator).
   - Apply the KDF described by the encryption scheme (salt + spin/iteration count + hash).
4. **Verify the password and decrypt**
   - **Standard:** use the verifier fields to validate the derived key and decrypt
     `EncryptedPackage`.
   - **Agile:** decrypt the package key via the password key-encryptor; then decrypt
     `EncryptedPackage`.
5. **(Recommended) Verify integrity**
   - Agile encryption can include a package-level HMAC (`dataIntegrity` /
      `encryptedHmacKey`+`encryptedHmacValue`). Verifying this detects tampering and wrong passwords
      earlier than “does the ZIP parse”.
   - For the exact HMAC target bytes and IV derivation rules (common source of bugs), see
     [`docs/22-ooxml-encryption.md`](./22-ooxml-encryption.md).
6. **Hand off to the normal workbook readers**
   - Once decrypted, `EncryptedPackage` yields the plaintext OPC ZIP. Route that ZIP through the
     existing `.xlsx`/`.xlsm`/`.xlsb` readers as if it were an unencrypted file.
     - Note: decrypted `.xlsb` packages contain `xl/workbook.bin`; route them through the `.xlsb`
       reader (`formula-xlsb` / `formula-io`’s `Workbook::Xlsb` open path).

Security requirements for this flow:

- **No plaintext to disk:** decrypt in memory and pass bytes/readers down the stack.
- **Bound resource use:** the decrypted package size is encoded in `EncryptedPackage`; enforce a
  maximum size and handle corrupt/truncated streams defensively.

### Implementation pointers (Formula code)

Useful entrypoints when working on encrypted workbook support:

- **Encryption detection / classification (OLE vs ZIP vs BIFF):**
  - `crates/formula-io/src/lib.rs`:
    - `detect_workbook_encryption`
    - `WorkbookEncryption`
    - error surfacing:
      - OOXML encrypted wrapper:
        - with `formula-io/encrypted-workbooks`: `Error::PasswordRequired` / `Error::InvalidPassword` /
          `Error::UnsupportedOoxmlEncryption`
        - without it: `Error::UnsupportedEncryption` (and `Error::UnsupportedOoxmlEncryption` for
          unknown `EncryptionInfo` versions)
      - legacy `.xls` `FILEPASS`:
        - without a password: `Error::EncryptedWorkbook`
        - with a password (via `open_workbook_with_password` / `open_workbook_model_with_password`):
          routed to `formula-xls` and surfaced as:
          - `Error::InvalidPassword` when the password is incorrect
          - `Error::UnsupportedEncryption` for unsupported/invalid `FILEPASS` encryption metadata
- **OOXML decrypt helpers (Agile + Standard/CryptoAPI):**
  - End-to-end decrypt (OLE wrapper → decrypted ZIP bytes):
    - `crates/formula-office-crypto` (end-to-end decrypt; supports Agile + Standard/CryptoAPI, plus an
      OOXML encryption writer: Agile by default; Standard/CryptoAPI AES is also supported)
  - MS-OFFCRYPTO parsing helpers + standalone decrypt primitives:
    - `crates/formula-offcrypto` (also used by `formula-xlsx`’s Standard path)
  - Standard/CryptoAPI specifics:
    - Parse Standard `EncryptionInfo`, derive/verify password keys: `crates/formula-offcrypto`
    - Decrypt Standard AES-ECB `EncryptedPackage` (baseline): `crates/formula-offcrypto/src/lib.rs`
      (`decrypt_encrypted_package_ecb`)
    - Decrypt Standard RC4 `EncryptedPackage`: see `docs/offcrypto-standard-cryptoapi-rc4.md`
    - Segment framing helper (size prefix + 0x1000 segment boundaries):
      `crates/formula-offcrypto/src/encrypted_package.rs`
- **Agile (4.4) OOXML decryption details (HMAC target bytes + IV usage gotchas):**
  - `docs/22-ooxml-encryption.md`
- **Standard (CryptoAPI) developer notes:**
  - Key derivation + verifier validation: `docs/offcrypto-standard-cryptoapi.md`
  - `EncryptedPackage` stream framing + truncation: `docs/offcrypto-standard-encryptedpackage.md`
- **Agile encryption primitives (password hash / key+IV derivation):**
  - `crates/formula-xlsx/src/offcrypto/crypto.rs`
- **Standard (CryptoAPI) RC4 algorithm writeup (KDF, 0x200 block size, test vectors):**
  - `docs/offcrypto-standard-cryptoapi-rc4.md`
- **Standard (CryptoAPI) AES/RC4 key derivation + verifier validation (implementation guide):**
  - `docs/offcrypto-standard-cryptoapi.md`

---

## Legacy `.xls` encryption (BIFF `FILEPASS`)

Legacy `.xls` files are also OLE/CFB containers, but the encrypted content is indicated inside the
BIFF workbook stream rather than via `EncryptedPackage`.

### Detection

In BIFF, encryption is signaled by the **`FILEPASS` record** in the workbook globals substream:

- Record id: `0x002F` (`FILEPASS`)
- Spec: **[MS-XLS]** `FILEPASS`

If `FILEPASS` is present, the workbook stream (and possibly other streams) must be decrypted before
BIFF parsing can proceed.

### Common `.xls` encryption schemes

`FILEPASS` identifies the encryption method; common variants in the wild include:

- **XOR obfuscation** (“XOR” / “XOR password”): very weak, legacy compatibility feature
- **RC4** (BIFF8 “RC4 encryption”)
- **CryptoAPI** (RC4/AES with CryptoAPI header structures; sometimes called “strong encryption”)

Formula’s decryption support is intentionally scoped:

- We aim to support the widely encountered BIFF8 variants needed for “real-world compatibility”.
- Rare/obsolete variants should produce a clear “unsupported encryption scheme” error.

Current behavior:

- `formula_xls::import_xls_path(...)` and `formula_xls::import_xls_bytes(...)` (no password) return
  `ImportError::EncryptedWorkbook` when `FILEPASS` is present.
- `formula_xls::import_xls_path_with_password(...)` and `formula_xls::import_xls_bytes_with_password(...)`
  support common BIFF `FILEPASS` encryption schemes, including:
  - XOR obfuscation (`wEncryptionType=0x0000`)
  - RC4 “standard” (`wEncryptionType=0x0001`, `wEncryptionSubType=0x0001`)
  - RC4 CryptoAPI (`wEncryptionType=0x0001`, `wEncryptionSubType=0x0002`)
    - Some Excel-produced workbooks use an older FILEPASS layout where the second field is
      `wEncryptionInfo=0x0004`; this is also supported.
  - (best-effort) BIFF5-era XOR obfuscation (Excel 5/95)
  Unsupported/unknown schemes surface as `ImportError::UnsupportedEncryption(..)` (for example,
  CryptoAPI AES algorithms are detected but not currently supported).

---

## Public API: supplying passwords

### `formula-io` API

Today, the `formula-io` crate:

- **Detects** encrypted workbooks via `detect_workbook_encryption(...)`.
- For encrypted OOXML containers (`EncryptionInfo` + `EncryptedPackage`):
  - With **`formula-io/encrypted-workbooks`** enabled:
    - `open_workbook(..)` / `open_workbook_model(..)` return `Error::PasswordRequired` when no
      password is provided (or `Error::UnsupportedOoxmlEncryption` when the `EncryptionInfo` version
      is unknown).
    - To supply a password:
      - Use `open_workbook_with_options(path, OpenOptions { password: Some(...) })` (preferred) or the
        convenience wrapper `open_workbook_with_password(path, Some(password))` to decrypt and open
        encrypted `.xlsx`/`.xlsm`/`.xlsb` workbooks into a `Workbook` (encrypted `.xlsb` opens as
        `Workbook::Xlsb`).
      - To open directly into a `formula_model::Workbook`, use:
        - `open_workbook_model_with_options(path, OpenOptions { password: Some(...) })`, or
        - the convenience wrapper `open_workbook_model_with_password(path, Some(password))`.
    - Supported schemes:
      - **Agile (4.4)** and **Standard/CryptoAPI** (`versionMinor == 2`; commonly `3.2`/`4.2`)
      - For Agile, `dataIntegrity` (HMAC) is validated when present; some real-world producers omit it.
  - Without that feature, encrypted OOXML containers surface `Error::UnsupportedEncryption` (or
    `Error::UnsupportedOoxmlEncryption` for unknown `EncryptionInfo` versions).
- For legacy `.xls` with BIFF `FILEPASS`:
  - without a password, `open_workbook(...)` / `open_workbook_model(...)` return
    `Error::EncryptedWorkbook`
  - password-aware APIs (`open_workbook_with_options` / `_with_password`) return:
    - `Error::PasswordRequired` when `password` is `None`
    - `Error::InvalidPassword` when a password is provided but decryption fails
    - and otherwise route to `formula_xls::import_xls_path_with_password(...)` (supports XOR, RC4
      “standard”, and RC4 CryptoAPI)

If you specifically need to open an **encrypted legacy `.xls`** workbook *today*, you can either:

- call `open_workbook_with_password` / `open_workbook_model_with_password` (which routes to the `.xls`
  importer), or
- bypass `formula-io` and use the `.xls` importer directly:

```rust
use formula_io::xls::import_xls_path_with_password;

let imported = import_xls_path_with_password("book.xls", Some("password"))?;
let workbook_model = imported.workbook;
```

Example (detection + UX routing):

```rust
use formula_io::{
    detect_workbook_encryption,
    open_workbook,
    open_workbook_with_options,
    Error,
    OpenOptions,
    WorkbookEncryption,
};

let path = "book.xlsx";
match open_workbook(path) {
    Ok(workbook) => {
        // Opened normally.
        let _ = workbook;
    }
    Err(Error::PasswordRequired { .. }) => {
        // Encrypted OOXML container (`EncryptionInfo` + `EncryptedPackage`).
        //
        // Note: this branch only applies when `formula-io/encrypted-workbooks` is enabled. Without
        // that feature, encrypted OOXML containers surface as `UnsupportedEncryption` instead.
        //
        // Prompt the user for a password, then retry with it:
        let password = "user-input-password";
        match open_workbook_with_options(
            path,
            OpenOptions {
                password: Some(password.to_string()),
                ..Default::default()
            },
        ) {
            Ok(workbook) => {
                // With the `formula-io/encrypted-workbooks` feature enabled, this succeeds for
                // Agile (4.4) and Standard/CryptoAPI (minor=2) encrypted `.xlsx`/`.xlsm`/`.xlsb`
                // when the password is correct.
                // Decrypted `.xlsb` payloads are returned as `Workbook::Xlsb`.
                let _ = workbook;
            }
            Err(Error::InvalidPassword { .. }) => {
                // Wrong password (or integrity mismatch, or unsupported encrypted format in this
                // layer).
            }
            Err(other) => return Err(other),
        }
    }
    Err(Error::UnsupportedEncryption { .. }) => {
        // Encrypted workbook detected, but encryption support isn't enabled/implemented in this
        // build. Suggest re-saving without encryption or enabling `formula-io/encrypted-workbooks`.
    }
    Err(Error::UnsupportedOoxmlEncryption { .. }) => {
        // Encrypted OOXML, but the EncryptionInfo version/scheme isn't supported.
    }
    Err(Error::EncryptedWorkbook { .. }) => {
        // Legacy `.xls` encryption (BIFF `FILEPASS`) or other encrypted container.
        // Use `detect_workbook_encryption` to classify further if needed:
        let _encryption = detect_workbook_encryption(path)?;
    }
    Err(other) => return Err(other),
}
```

With the `formula-io` cargo feature **`encrypted-workbooks`** enabled:

- Password-aware entrypoints (`open_workbook_with_options` / `_with_password`) will **decrypt and open**
  Agile (4.4) and Standard/CryptoAPI (minor=2; commonly `3.2`/`4.2`) encrypted `.xlsx`/`.xlsm`/`.xlsb`
  workbooks in memory. For Agile, `dataIntegrity` (HMAC) is validated when present; some real-world
  producers omit it.
  - Encrypted `.xlsb` workbooks are supported in native `formula-io` (decrypted payload contains
    `xl/workbook.bin`) and open as `Workbook::Xlsb`. The WASM loader also supports decrypted `.xlsb`
    payloads via `WasmWorkbook::fromEncryptedXlsxBytes`.

Without that feature, encrypted OOXML containers surface as `UnsupportedEncryption` (the password-aware
entrypoints do not decrypt them end-to-end).

#### API notes

- Passwords are treated as **UTF-8 strings at the API boundary** and encoded internally according
  to the relevant spec requirements (typically UTF-16LE for key derivation).
- Callers should avoid logging passwords or embedding them in error messages.
- Decrypted bytes are sensitive: do not write them to disk. See
  [Security notes](#security-notes-handling-decrypted-bytes-safely).
- `detect_workbook_encryption` can be used to decide whether to prompt for a password before
  attempting a full open (see below).

#### Preflight detection (optional)

Callers that want to decide whether to prompt for a password *before* attempting a full open can
use `detect_workbook_encryption`:

```rust
use formula_io::{detect_workbook_encryption, WorkbookEncryption};

match detect_workbook_encryption("book.xlsx")? {
    WorkbookEncryption::None => {}
    WorkbookEncryption::OoxmlEncryptedPackage { .. } => {
        // Encrypted OOXML wrapper (`EncryptionInfo` + `EncryptedPackage`).
    }
    WorkbookEncryption::LegacyXlsFilePass { .. } => {
        // Legacy `.xls` workbook stream contains BIFF `FILEPASS`.
    }
}
```

### Desktop app flow (IPC + password prompt)

In the desktop app, the file-open path is interactive. The intended flow is:

1. Frontend requests open: `openWorkbook({ path })`
2. Backend attempts to open without a password.
3. If the backend returns **PasswordRequired**, the frontend shows a password prompt.
   - For encrypted OOXML containers, this requires `formula-io/encrypted-workbooks` to be enabled.
   - Otherwise (no decryption support), the backend may return `UnsupportedEncryption` instead.
4. Frontend retries open with a password: `openWorkbook({ path, password })`
5. If the password is wrong, show an “invalid password” error and allow retry/cancel.

Important: the password should only ever exist in memory for the duration of the open attempt.
It must not be persisted in:

- app logs
- crash reports
- recent-file metadata
- autosave/backup caches

---

## Error semantics

Encrypted workbook handling should distinguish at least these cases:

1. **Password required** (encrypted workbook detected, but no password was provided)
   - Surface as: `PasswordRequired`
   - UI action: prompt user for password and retry.

2. **Invalid password** (password provided, but key verification fails)
   - Surface as: `InvalidPassword`
   - UI action: allow retry; do not treat as a “corrupt file” error.

3. **Unsupported encryption scheme** (recognized encrypted container, but scheme not implemented)
   - Surface as: `Error::UnsupportedOoxmlEncryption { version_major, version_minor }` (for OOXML),
     and/or a future more specific “unsupported scheme” error once we plumb lower-level crypto
     errors through.
   - UI action: explain limitation and suggest re-saving without encryption in Excel.

4. **Decrypted payload is not a recognized workbook package** (decryption succeeded, but the decrypted
   bytes do not appear to be a supported Excel workbook)
   - Surface as: `Error::OpenXlsx { .. }` / `Error::OpenXlsb { .. }` / other parse errors (depending
     on which reader it routes to).
   - UI action: treat as “file corrupted or unsupported” (if the password is known-good, the file
     may not be an Excel workbook or the encrypted wrapper may be malformed).

5. **Corrupt encrypted wrapper** (missing streams, malformed `EncryptionInfo`, truncated payload)
   - Surface as: a dedicated “corrupt encrypted container” error (future); today this may surface
     as a generic parse/IO error depending on where it fails.

These distinctions matter for UX and telemetry: “needs password” is a normal user workflow, while
“unsupported scheme” is an engineering coverage gap.

Current behavior in `formula-io`:

- Encrypted OOXML wrappers surface:
  - `Error::UnsupportedOoxmlEncryption` (unrecognized `EncryptionInfo` version)
  - With `formula-io/encrypted-workbooks` enabled:
    - `Error::PasswordRequired` (no password provided)
    - `Error::InvalidPassword` (wrong password/verifier mismatch *or* Agile `dataIntegrity` mismatch)
    - Note: encrypted `.xlsb` containers decrypt to a ZIP package containing `xl/workbook.bin` and are
      opened as `Workbook::Xlsb`.
  - Without that feature: `Error::UnsupportedEncryption` (encrypted OOXML decryption is not enabled in
    this build)
- Legacy `.xls` encryption (`FILEPASS`) is surfaced as:
  - `Error::EncryptedWorkbook` via `open_workbook(..)` / `open_workbook_model(..)` (no password support),
    and
  - `Error::PasswordRequired` / `Error::InvalidPassword` via `open_workbook_with_password(..)` /
    `open_workbook_model_with_password(..)` (attempts the decrypting `.xls` importer; XOR, RC4
    “standard”, RC4 CryptoAPI).

### Mapping to existing Rust error types

Lower-level crypto/decryption code already has more granular error variants. When wiring password
support through `formula-io`, we should preserve these distinctions rather than collapsing them
back into a generic “encrypted workbook” error:

- `formula_io::Error::PasswordRequired { .. }` → **Password required**
- `formula_io::Error::InvalidPassword { .. }` → **Invalid password**
- `formula_io::Error::UnsupportedOoxmlEncryption { .. }` → **Unsupported encryption scheme**
- `formula_io::Error::UnsupportedEncryption { .. }` → **Unsupported encryption scheme**
- `formula_io::Error::EncryptedWorkbook { .. }` → **Password required (legacy `.xls` encryption / FILEPASS)**

- `formula_xlsx::offcrypto::OffCryptoError::WrongPassword` → **Invalid password**
- `formula_xlsx::offcrypto::OffCryptoError::IntegrityMismatch` → **Invalid password** *or* **corrupt file**
  - UX should not claim “file is corrupted” with certainty; treat as “password incorrect or file
    corrupted”.
- `formula_xlsx::offcrypto::OffCryptoError::UnsupportedEncryptionVersion { .. }` and
  `Unsupported*` variants → **Unsupported encryption scheme**

- `formula_offcrypto::OffcryptoError::InvalidPassword` → **Invalid password**
- `formula_offcrypto::OffcryptoError::UnsupportedKeyEncryptor { available }` → **Unsupported encryption scheme** (no password key-encryptor present)
- `formula_offcrypto::OffcryptoError::UnsupportedVersion { .. }` and `UnsupportedAlgorithm(..)` → **Unsupported encryption scheme**
- `formula_offcrypto::OffcryptoError::InvalidEncryptionInfo { .. }` / `Truncated { .. }` → **Corrupt encrypted wrapper**

- `formula_io::offcrypto::EncryptedPackageError::*` → **Corrupt file / invalid encrypted wrapper**
  - e.g. `StreamTooShort`, `CiphertextLenNotBlockAligned`, `DecryptedTooShort`

- `formula_xls::ImportError::EncryptedWorkbook` → **Password required**
- `formula_xls::ImportError::InvalidPassword` → **Invalid password**
- `formula_xls::ImportError::UnsupportedEncryption(..)` → **Unsupported encryption scheme**
- `formula_xls::ImportError::Decrypt(..)` → **Corrupt encrypted wrapper**

---

## Saving / round-trip limitations

Opening an encrypted workbook inherently produces a **decrypted in-memory representation**.

### Default save behavior

Unless the caller explicitly re-wraps the output workbook with a new `EncryptionInfo` +
`EncryptedPackage` (OOXML) (or re-emits BIFF encryption structures for legacy `.xls`), a save
operation will write a **decrypted** file.

### Preserving OOXML encryption on save (encrypted `.xlsx`/`.xlsm`/`.xlsb`)

With the `formula-io/encrypted-workbooks` feature enabled, `formula-io` provides an opt-in API to
round-trip Office-encrypted OOXML **OLE/CFB** containers while preserving extra OLE metadata streams:

- Open + capture preserved OLE entries:
  `open_workbook_with_password_and_preserved_ole(path, Some(password))`
- Save with encryption preserved:
  `OpenedWorkbookWithPreservedOle::save_preserving_encryption(out_path, password)`

Notes:

- This path re-wraps the decrypted OOXML package with freshly generated `EncryptionInfo` +
  `EncryptedPackage` streams using `crates/formula-office-crypto`.
  - Today, `formula-io` uses `formula_office_crypto::EncryptOptions::default()` for output, which
    means the saved file is typically re-encrypted as **Agile (4.4)** (even if the input used Standard).
- The password is required at save time because `OpenedWorkbookWithPreservedOle` does not store it.
- This path can also re-encrypt and save `.xlsb` workbooks (use an `.xlsb` output extension if you
  want to keep the binary format).
- Legacy `.xls` `FILEPASS` encryption is not preserved on save.

### Desktop app behavior

- The desktop save command supports writing an **encrypted** `.xlsx`/`.xlsm` when a `password` is
  provided (it encrypts the plaintext OOXML package into an OLE/CFB wrapper via
  `formula_office_crypto::encrypt_package_to_ole` / `OpenedWorkbookWithPreservedOle::save_preserving_encryption`).
- Opening encrypted `.xlsb` is supported, but the decrypted payload is not backed by an on-disk XLSB
  ZIP package (`origin_xlsb_path`), so the desktop app currently cannot save back to `.xlsb` and
  forces “Save As” (save as `.xlsx` instead).
- Saving encrypted `.xlsb` is currently rejected (save as `.xlsx` instead).

Product/UX mitigations (if the caller is *not* using the preserve-encryption save path):

- Warn before saving if the origin workbook was encrypted.
- Use “Save As” flows that default to a new filename/extension to make the change explicit.

---

## Security notes (handling decrypted bytes safely)

When implementing (or calling) encrypted-workbook support:

- Prefer **in-memory decryption**. Avoid writing decrypted workbook bytes to disk.
  - If a spill-to-disk fallback is unavoidable (large files), prefer OS mechanisms that keep data
    off persistent storage (e.g. `memfd` on Linux) and ensure secure deletion semantics.
- Minimize copies of plaintext data and secrets:
  - Keep decrypted buffers scoped and drop them as soon as parsing is complete.
  - Consider `zeroize`-style clearing of password/key material when practical.
- Treat decrypted content as sensitive:
  - Do not include decrypted snippets in logs/telemetry.
  - Be careful with crash dumps and “upload file for support” tooling.

---

## Test fixtures in this repo

- Encrypted/password-protected OOXML workbook fixtures live under `fixtures/encrypted/ooxml/` (for
  example `.xlsx` and `.xlsm`). The repo currently includes:
  - `fixtures/encrypted/ooxml/plaintext.xlsx` (unencrypted ZIP/OPC workbook used as the known-good plaintext)
  - `fixtures/encrypted/ooxml/plaintext-excel.xlsx` (unencrypted ZIP/OPC workbook produced by Microsoft Excel; used to exercise additional real-world ZIP/part variations)
  - `fixtures/encrypted/ooxml/plaintext-large.xlsx` (unencrypted ZIP/OPC workbook used to exercise multi-segment decryption; intentionally > 4096 bytes)
  - `fixtures/encrypted/ooxml/agile.xlsx` (Agile encryption; `EncryptionInfo` 4.4)
  - `fixtures/encrypted/ooxml/agile-large.xlsx` (Agile encryption; `EncryptionInfo` 4.4; decrypts to `plaintext-large.xlsx`)
  - `fixtures/encrypted/ooxml/standard.xlsx` (Standard encryption; `EncryptionInfo` 3.2)
  - `fixtures/encrypted/ooxml/standard-4.2.xlsx` (Standard encryption; `EncryptionInfo` 4.2)
  - `fixtures/encrypted/ooxml/standard-unicode.xlsx` (Standard encryption; `EncryptionInfo` 4.2; Unicode password `pässwörd🔒` in NFC form; decrypts to `plaintext.xlsx`)
  - `fixtures/encrypted/ooxml/standard-rc4.xlsx` (Standard encryption; `EncryptionInfo` 3.2; RC4 CryptoAPI)
  - `fixtures/encrypted/ooxml/standard-large.xlsx` (Standard encryption; `EncryptionInfo` 3.2; decrypts to `plaintext-large.xlsx`)
  - `fixtures/encrypted/ooxml/agile-empty-password.xlsx` (Agile encryption; `EncryptionInfo` 4.4; empty password `""`)
  - `fixtures/encrypted/ooxml/agile-unicode.xlsx` (Agile encryption; `EncryptionInfo` 4.4; Unicode password `pässwörd` in NFC normalization form)
  - `fixtures/encrypted/ooxml/agile-unicode-excel.xlsx` (Agile encryption; `EncryptionInfo` 4.4; Unicode password `pässwörd🔒` in NFC form; decrypts to `plaintext-excel.xlsx`)
  - `fixtures/encrypted/ooxml/plaintext-basic.xlsm` (unencrypted ZIP/OPC macro-enabled workbook used as the known-good `.xlsm` plaintext)
  - `fixtures/encrypted/ooxml/basic-password.xlsm` (Agile encryption; `EncryptionInfo` 4.4; password `password`; macro-enabled workbook)
  - `fixtures/encrypted/ooxml/agile-basic.xlsm` (Agile encryption; `EncryptionInfo` 4.4; decrypts to `plaintext-basic.xlsm`)
  - `fixtures/encrypted/ooxml/standard-basic.xlsm` (Standard encryption; `EncryptionInfo` 3.2; decrypts to `plaintext-basic.xlsm`)
  See `fixtures/encrypted/ooxml/README.md` for more fixture details.
  These files are OLE/CFB wrappers (not ZIP/OPC), so they must not live under `fixtures/xlsx/`
  where the round-trip corpus is enumerated via `xlsx-diff::collect_fixture_paths`.
- Additional “real-world” encrypted workbook fixtures (including encrypted `.xlsb` and legacy
  encrypted `.xls`) live under `fixtures/encrypted/`; see `fixtures/encrypted/README.md`.
- Encrypted legacy `.xls` fixtures for `formula-xls` tests live under:
  - `crates/formula-xls/tests/fixtures/encrypted/` (deterministic, test-generated), and
  - `crates/formula-xls/tests/fixtures/encrypted_rc4_cryptoapi_boundary.xls` (Microsoft Excel-produced;
    exercises RC4 CryptoAPI legacy FILEPASS layout + 1024-byte rekey boundary behavior).
- Encryption-focused tests reference these fixtures explicitly (they are not part of the ZIP/OPC
  round-trip corpus). See `fixtures/encrypted/ooxml/README.md` for the canonical list, passwords,
  provenance, and test references.
- `crates/formula-io/tests/encrypted_ooxml.rs` (and `encrypted_ooxml_fixtures.rs`) asserts that
  opening these fixtures without a password surfaces an error mentioning
  encryption/password protection (guards the “password required” UX path).
- End-to-end decryption (including empty-password, Unicode-password, macro-enabled `.xlsm`, and multi-segment coverage) is exercised by
  `crates/formula-io/tests/encrypted_ooxml_decrypt.rs` and
  `crates/formula-xlsx/tests/encrypted_ooxml_decrypt.rs`.
- Some encryption coverage is exercised with **synthetic** containers generated directly in tests
  (for example `crates/formula-io/tests/encrypted_xls.rs`, plus a synthetic encrypted OOXML wrapper
  in `crates/formula-io/tests/encrypted_ooxml.rs`).

---

## References (specs)

- **MS-OFFCRYPTO** — Office Document Cryptography Structure Specification  
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
- **MS-CFB** — Compound File Binary File Format (OLE Structured Storage)  
  https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
- **MS-XLS** — Excel Binary File Format (`FILEPASS`, BIFF globals)  
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/
- **MS-XLSX** — Office Open XML SpreadsheetML Package Structure  
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsx/
