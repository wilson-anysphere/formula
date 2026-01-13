# Encrypted / Password-Protected Excel Workbooks

This document covers **file encryption** (a password is required to *open* the file) and how it
differs from Excel’s **workbook/worksheet protection** features (a password is required to *edit*
certain parts of the workbook, but the file contents are not encrypted).

Formula’s goal is to open encrypted spreadsheets when possible, surface **actionable** errors when
not, and avoid security pitfalls (like accidentally persisting decrypted bytes to disk).

## Related docs

- [`docs/office-encryption.md`](./office-encryption.md) — maintainer-level reference (supported
  parameter subsets, KDF nuances, writer defaults).
- [`docs/21-offcrypto.md`](./21-offcrypto.md) — short MS-OFFCRYPTO overview focused on “what the file
  looks like”, scheme detection, and `formula-io` password APIs.
- [`docs/offcrypto-standard-encryptedpackage.md`](./offcrypto-standard-encryptedpackage.md) —
  Standard/CryptoAPI AES `EncryptedPackage` segmenting + IV derivation notes.

## Status (current behavior vs intended behavior)

**Current behavior (in this repo today):**

- Encrypted workbooks are **detected** and rejected with a clear error:
  - Encrypted OOXML (`EncryptionInfo` + `EncryptedPackage`) yields `formula_io::Error::PasswordRequired`
    (and `UnsupportedOoxmlEncryption` for unknown versions).
  - Legacy `.xls` with `FILEPASS` yields `formula_io::Error::EncryptedWorkbook` unless the caller
    uses the `.xls` importer with a password (see below).
  - The desktop app surfaces an “encrypted workbook not supported” message.
- Password-based OOXML decryption is not yet implemented end-to-end in `formula-io`.
  - The default open path (`open_workbook` / `open_workbook_model`) has no password parameter.
  - `open_workbook_with_password` / `open_workbook_model_with_password` exist so callers can
    surface **InvalidPassword** vs **PasswordRequired** in UX flows, but they do not decrypt yet.
  - Note: the legacy `.xls` importer (`formula-xls`) exposes `import_xls_path_with_password(...)`
    for a subset of BIFF8 encryption (see below).

**Intended behavior (when decryption + password plumbing is implemented):**

- Support opening Excel-encrypted workbooks without writing decrypted bytes to disk.
- Distinguish “password required” vs “invalid password” vs “unsupported encryption scheme”
  (see [Error semantics](#error-semantics)).

### Support matrix (current vs planned)

| File type | Encryption marker | Schemes (common) | Current behavior | Planned/target behavior |
|---|---|---|---|---|
| `.xlsx` / `.xlsm` / `.xlsb` (OOXML) | OLE/CFB streams `EncryptionInfo` + `EncryptedPackage` | Agile (4.4), Standard (minor=2; commonly 3.2) | `PasswordRequired` (and `UnsupportedOoxmlEncryption` for unknown versions; `open_workbook_with_password` surfaces `InvalidPassword`) | Decrypt + open; surface `PasswordRequired` / `InvalidPassword` / `UnsupportedOoxmlEncryption` (or a more specific “unsupported scheme” error) |
| `.xls` (BIFF) | BIFF `FILEPASS` record in workbook stream | XOR, RC4, CryptoAPI | `formula-io`: detect + `Error::EncryptedWorkbook` (no password plumbing yet). `formula-xls`: supports BIFF8 RC4 CryptoAPI when a password is provided. | Plumb password into `formula-io` and expand scheme coverage as needed (see [Legacy `.xls` encryption](#legacy-xls-encryption-biff-filepass)) |

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
| **Standard** | `*.2` (`versionMinor == 2`, `versionMajor ∈ {2,3,4}` in the wild; commonly `3.2`) | CryptoAPI-style header/verifier structures (binary). |
| **Agile** | `4.4` | XML-based encryption descriptor. |

Implementation notes:

- `crates/formula-io/src/bin/ooxml-encryption-info.rs` prints a one-line scheme/version summary
  based on the `(majorVersion, minorVersion)` header (e.g. `3.2` → Standard, `4.4` → Agile).
  Run it with:

  ```bash
  bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- path/to/encrypted.xlsx
  ```
- `crates/formula-offcrypto` parses `EncryptionInfo` for both:
  - Standard (CryptoAPI; `versionMinor == 2`) header + verifier structures, and
  - Agile (4.4) XML (password key-encryptor subset),
  and implements Standard password→key derivation + verifier checks.
- `crates/formula-io/src/offcrypto/encrypted_package.rs` decrypts the Standard/CryptoAPI
  `EncryptedPackage` stream once you have the derived file key and verifier salt.

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

- **8 bytes**: `original_package_size` (`u64`, little-endian)
- remaining bytes: encrypted OPC package data (ciphertext + padding)

After decrypting the ciphertext, the plaintext bytes should be truncated to `original_package_size`
to recover the real workbook package (a normal ZIP/OPC archive for `.xlsx`/`.xlsm`, or a ZIP/OPC
archive containing `xl/workbook.bin` for `.xlsb`).

The encryption *mode* differs by scheme:

- **Standard (CryptoAPI; `versionMinor == 2`):** AES-CBC over **0x1000 (4096)-byte plaintext segments** with a
  per-segment IV. Segment `i` uses `IV = SHA1(verifierSalt || LE32(i))[0..16]` and is decrypted
  independently; the concatenated plaintext is truncated to `original_package_size` (see
  `docs/offcrypto-standard-encryptedpackage.md`).
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
6. **Hand off to the normal workbook readers**
   - Once decrypted, `EncryptedPackage` yields the plaintext OPC ZIP. Route that ZIP through the
     existing `.xlsx`/`.xlsm`/`.xlsb` readers as if it were an unencrypted file.

Security requirements for this flow:

- **No plaintext to disk:** decrypt in memory and pass bytes/readers down the stack.
- **Bound resource use:** the decrypted package size is encoded in `EncryptedPackage`; enforce a
  maximum size and handle corrupt/truncated streams defensively.

### Implementation pointers (Formula code)

Useful entrypoints when working on encrypted workbook support:

- **Encryption detection / classification (OLE vs ZIP vs BIFF):**
  - `crates/formula-io/src/lib.rs`:
    - `detect_workbook_encryption`
    - `WorkbookEncryptionKind`
    - error surfacing:
      - OOXML encrypted wrapper: `Error::PasswordRequired` / `Error::InvalidPassword` /
        `Error::UnsupportedOoxmlEncryption`
      - legacy `.xls` `FILEPASS`: `Error::EncryptedWorkbook` (until password plumbing is added)
- **Standard (CryptoAPI) helpers:**
  - End-to-end decrypt (OLE wrapper → decrypted ZIP bytes; Agile + Standard):
    `crates/formula-office-crypto`
  - Parse `EncryptionInfo`, derive/verify password key:
    `crates/formula-offcrypto`
  - Decrypt `EncryptedPackage` (given a derived key + verifier salt):
    `crates/formula-io/src/offcrypto/encrypted_package.rs`
  - Segment framing helper (size prefix + 0x1000 segment boundaries):
    `crates/formula-offcrypto/src/encrypted_package.rs`
- **Agile encryption primitives (password hash / key+IV derivation):**
  - `crates/formula-xlsx/src/offcrypto/crypto.rs`
- **Standard (CryptoAPI) RC4 algorithm writeup (KDF, 0x200 block size, test vectors):**
  - `docs/offcrypto-standard-cryptoapi-rc4.md`

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

- `formula_xls::import_xls_path(...)` (no password) returns `ImportError::EncryptedWorkbook` when
  `FILEPASS` is present.
- `formula_xls::import_xls_path_with_password(...)` supports **BIFF8 RC4 CryptoAPI**
  (`wEncryptionType=0x0001`, `wEncryptionSubType=0x0002`). Other `FILEPASS` variants are currently
  treated as unsupported.

---

## Public API: supplying passwords

### `formula-io` API

Today, the `formula-io` crate:

- **Detects** encrypted workbooks via `detect_workbook_encryption(...)`.
- For encrypted OOXML containers (`EncryptionInfo` + `EncryptedPackage`), `open_workbook(...)` /
  `open_workbook_model(...)` return `Error::PasswordRequired` (or `Error::UnsupportedOoxmlEncryption`
  when the `EncryptionInfo` version is unknown).
  - `open_workbook_with_password(...)` / `open_workbook_model_with_password(...)` accept an
    optional password and return `Error::InvalidPassword` when one is supplied (until decryption is
    implemented).
- For legacy `.xls` with BIFF `FILEPASS`, `formula-io` currently returns `Error::EncryptedWorkbook`
  (password plumbing is not integrated at the `formula-io` layer yet).

If you specifically need to open an **encrypted legacy `.xls`** workbook *today*, you can bypass
`formula-io` and use the `.xls` importer directly (limited to BIFF8 RC4 CryptoAPI):

```rust
use formula_io::xls::import_xls_path_with_password;

let imported = import_xls_path_with_password("book.xls", "password")?;
let workbook_model = imported.workbook;
```

Example (detection + UX routing):

```rust
use formula_io::{
    detect_workbook_encryption,
    open_workbook,
    open_workbook_with_password,
    Error,
    WorkbookEncryptionKind,
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
        // Prompt the user for a password, then retry with it:
        let password = "user-input-password";
        let _ = open_workbook_with_password(path, Some(password))?;
    }
    Err(Error::InvalidPassword { .. }) => {
        // A password was provided but is incorrect (once decryption is implemented).
    }
    Err(Error::UnsupportedOoxmlEncryption { .. }) => {
        // Encrypted OOXML, but the EncryptionInfo version/scheme isn't supported.
    }
    Err(Error::EncryptedWorkbook { .. }) => {
        // Legacy `.xls` encryption (BIFF `FILEPASS`) or other encrypted container.
        // Use `detect_workbook_encryption` to classify further if needed:
        let _kind = detect_workbook_encryption(path)?
            .map(|info| info.kind)
            .unwrap_or(WorkbookEncryptionKind::UnknownOleEncrypted);
    }
    Err(other) => return Err(other),
}
```

The intended end-state is that `formula-io` will actually **decrypt and open** encrypted OOXML
workbooks when a correct password is provided (and optionally verify Agile `dataIntegrity`).
Today, the password-aware entrypoints exist primarily to enable better UX/error handling
(`PasswordRequired` vs `InvalidPassword`) while decryption support is still landing.

#### API notes

- Passwords are treated as **UTF-8 strings at the API boundary** and encoded internally according
  to the relevant spec requirements (typically UTF-16LE for key derivation).
- Callers should avoid logging passwords or embedding them in error messages.
- `detect_workbook_encryption` can be used to decide whether to prompt for a password before
  attempting a full open (see below).

#### Preflight detection (optional)

Callers that want to decide whether to prompt for a password *before* attempting a full open can
use `detect_workbook_encryption`:

```rust
use formula_io::{detect_workbook_encryption, WorkbookEncryptionKind};

if let Some(info) = detect_workbook_encryption("book.xlsx")? {
    match info.kind {
        WorkbookEncryptionKind::OoxmlOleEncryptedPackage => {
            // Encrypted OOXML wrapper (`EncryptionInfo` + `EncryptedPackage`).
        }
        WorkbookEncryptionKind::XlsFilepass => {
            // Legacy `.xls` workbook stream contains BIFF `FILEPASS`.
        }
        WorkbookEncryptionKind::UnknownOleEncrypted => {
            // Some other encrypted OLE container (treat as encrypted/unsupported).
        }
    }
}
```

### Desktop app flow (IPC + password prompt)

In the desktop app, the file-open path is interactive. The intended flow is:

1. Frontend requests open: `openWorkbook({ path })`
2. Backend attempts to open without a password.
3. If the backend returns **PasswordRequired**, the frontend shows a password prompt.
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

4. **Corrupt encrypted wrapper** (missing streams, malformed `EncryptionInfo`, truncated payload)
   - Surface as: a dedicated “corrupt encrypted container” error (future); today this may surface
     as a generic parse/IO error depending on where it fails.

These distinctions matter for UX and telemetry: “needs password” is a normal user workflow, while
“unsupported scheme” is an engineering coverage gap.

Current behavior in `formula-io`:

- Encrypted OOXML wrappers already distinguish:
  - `Error::PasswordRequired` (no password provided)
  - `Error::InvalidPassword` (password provided; decryption not yet implemented, so this is a UX
    signal rather than a verified cryptographic check)
  - `Error::UnsupportedOoxmlEncryption` (unrecognized `EncryptionInfo` version)
- Legacy `.xls` encryption (`FILEPASS`) is still surfaced as `Error::EncryptedWorkbook` at the
  `formula-io` layer unless callers use `formula_xls::import_xls_path_with_password`.

### Mapping to existing Rust error types

Lower-level crypto/decryption code already has more granular error variants. When wiring password
support through `formula-io`, we should preserve these distinctions rather than collapsing them
back into a generic “encrypted workbook” error:

- `formula_io::Error::PasswordRequired { .. }` → **Password required**
- `formula_io::Error::InvalidPassword { .. }` → **Invalid password**
- `formula_io::Error::UnsupportedOoxmlEncryption { .. }` → **Unsupported encryption scheme**
- `formula_io::Error::EncryptedWorkbook { .. }` → **Encrypted workbook (unsupported at this layer)**

- `formula_xlsx::offcrypto::OffCryptoError::WrongPassword` → **Invalid password**
- `formula_xlsx::offcrypto::OffCryptoError::IntegrityMismatch` → **Invalid password** *or* **corrupt file**
  - UX should not claim “file is corrupted” with certainty; treat as “password incorrect or file
    corrupted”.
- `formula_xlsx::offcrypto::OffCryptoError::UnsupportedEncryptionVersion { .. }` and
  `Unsupported*` variants → **Unsupported encryption scheme**

- `formula_offcrypto::OffcryptoError::InvalidPassword` → **Invalid password**
- `formula_offcrypto::OffcryptoError::UnsupportedVersion { .. }` and `UnsupportedAlgorithm(..)` → **Unsupported encryption scheme**
- `formula_offcrypto::OffcryptoError::InvalidEncryptionInfo { .. }` / `Truncated { .. }` → **Corrupt encrypted wrapper**

- `formula_io::offcrypto::EncryptedPackageError::*` → **Corrupt file / invalid encrypted wrapper**
  - e.g. `StreamTooShort`, `CiphertextLenNotBlockAligned`, `DecryptedTooShort`

- `formula_xls::DecryptError::WrongPassword` → **Invalid password**
- `formula_xls::DecryptError::UnsupportedEncryption` → **Unsupported encryption scheme**
- `formula_xls::DecryptError::InvalidFormat(..)` → **Corrupt encrypted wrapper**

---

## Saving / round-trip limitations

Opening an encrypted workbook inherently produces a **decrypted in-memory representation**.

Unless Formula explicitly re-wraps the output workbook with a new `EncryptionInfo` +
`EncryptedPackage` (OOXML) or re-emits BIFF encryption structures (`.xls`), a “save” operation will
write a **decrypted** file.

Current limitation to document and handle in product:

- **Editing + saving an encrypted workbook may drop encryption** unless/until re-encryption is
  implemented.

Mitigations:

- UI warning before saving if the origin workbook was encrypted.
- “Save As” flows that default to a new filename/extension to make the change explicit.
- Optional support for “re-encrypt on save” once implemented.

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
  example `.xlsx`, `.xlsm`, `.xlsb`). The repo currently includes:
  - `fixtures/encrypted/ooxml/agile.xlsx` (Agile encryption; `EncryptionInfo` 4.4)
  - `fixtures/encrypted/ooxml/standard.xlsx` (Standard encryption; `EncryptionInfo` 3.2)
  See `fixtures/encrypted/ooxml/README.md` for more fixture details.
  These files are OLE/CFB wrappers (not ZIP/OPC), so they must not live under `fixtures/xlsx/`
  where the round-trip corpus is enumerated via `xlsx-diff::collect_fixture_paths`.
- `crates/formula-io/tests/encrypted_ooxml.rs` asserts that opening these fixtures without a
  password surfaces an error mentioning encryption/password protection (guards the “password
  required” UX path).
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
