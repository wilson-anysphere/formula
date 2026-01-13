# Encrypted / Password-Protected Excel Workbooks

This document covers **file encryption** (a password is required to *open* the file) and how it
differs from Excel’s **workbook/worksheet protection** features (a password is required to *edit*
certain parts of the workbook, but the file contents are not encrypted).

Formula’s goal is to open encrypted spreadsheets when possible, surface **actionable** errors when
not, and avoid security pitfalls (like accidentally persisting decrypted bytes to disk).

## Status (current behavior vs intended behavior)

**Current behavior (in this repo today):**

- Encrypted workbooks are **detected** and rejected with a clear error:
  - `formula-io` returns `formula_io::Error::EncryptedWorkbook`.
  - The desktop app surfaces an “encrypted workbook not supported” message.
- Password-based decryption is not yet wired into the public open APIs, so callers cannot supply a
  password to open an encrypted workbook.

**Intended behavior (when decryption + password plumbing is implemented):**

- Support opening Excel-encrypted workbooks without writing decrypted bytes to disk.
- Distinguish “password required” vs “invalid password” vs “unsupported encryption scheme”
  (see [Error semantics](#error-semantics)).

### Support matrix (current vs planned)

| File type | Encryption marker | Schemes (common) | Current behavior | Planned/target behavior |
|---|---|---|---|---|
| `.xlsx` / `.xlsm` / `.xlsb` (OOXML) | OLE/CFB streams `EncryptionInfo` + `EncryptedPackage` | Agile (4.4), Standard (3.2) | Detect + return `Error::EncryptedWorkbook` | Decrypt + open; surface `PasswordRequired` / `InvalidPassword` / `UnsupportedEncryptionScheme` |
| `.xls` (BIFF) | BIFF `FILEPASS` record in workbook stream | XOR, RC4, CryptoAPI | Detect + return `Error::EncryptedWorkbook` | Decrypt + open (scope TBD; see [Legacy `.xls` encryption](#legacy-xls-encryption-biff-filepass)) |

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
| **Standard** | `3.2` (versionMinor = `2`) | CryptoAPI-style header/verifier structures (binary). |
| **Agile** | `4.4` | XML-based encryption descriptor. |

Implementation notes:

- Formula’s Standard parser (`crates/formula-offcrypto`) treats **`versionMinor == 2`** (and
  `versionMajor ∈ {2,3,4}`) as Standard encryption; Excel commonly emits `3.2`.
- Formula includes a small helper binary that prints the one-line scheme/version for an encrypted
  OOXML file: `crates/formula-io/src/bin/ooxml-encryption-info.rs`.

### Supported OOXML encryption schemes

The `EncryptionInfo` stream encodes one of the schemes defined in **[MS-OFFCRYPTO]**. In practice,
Excel-produced encrypted workbooks primarily use:

- **Agile Encryption** (modern; Office 2010+)
- **Standard Encryption** (older; Office 2007 era)

Formula’s encrypted-workbook support targets these two schemes:

- **Agile** (`EncryptionInfo` version 4.4; XML-based descriptor inside `EncryptionInfo`)
- **Standard** (`EncryptionInfo` version 3.2; CryptoAPI-style header/verifier)

Everything else should fail with a specific “unsupported encryption scheme” error (see
[Error semantics](#error-semantics)).

### High-level decrypt/open algorithm (OOXML)

At a high level, opening a password-encrypted OOXML workbook is:

1. **Detect the OLE wrapper**
   - Open the file as a CFB container.
   - Confirm `EncryptionInfo` + `EncryptedPackage` streams exist.
2. **Read and classify `EncryptionInfo`**
   - Parse the `(major, minor)` version header.
   - **Standard (3.2 / minor=2):** parse the binary CryptoAPI header + verifier structures.
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
    - `Error::EncryptedWorkbook` (legacy “not supported” umbrella error)
- **Standard (CryptoAPI) `EncryptedPackage` AES-ECB decrypt helper (given a derived key):**
  - `crates/formula-offcrypto/src/lib.rs`
- **Agile encryption primitives (password hash / key+IV derivation):**
  - `crates/formula-xlsx/src/offcrypto/crypto.rs`

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

Current behavior: Formula detects `FILEPASS` during `.xls` import and returns an
“encrypted workbook not supported” error rather than attempting decryption.

---

## Public API: supplying passwords

### `formula-io` API

Today, the `formula-io` crate:

- **Detects** encrypted workbooks via `detect_workbook_encryption(...)`.
- Returns `Error::EncryptedWorkbook` from `open_workbook(...)` / `open_workbook_model(...)` for
  password-encrypted workbooks (OOXML OLE wrapper and legacy `.xls` `FILEPASS`).

Example (detection + UX routing):

```rust
use formula_io::{
    detect_workbook_encryption, open_workbook, Error, WorkbookEncryptionKind,
};

let path = "book.xlsx";
match open_workbook(path) {
    Ok(workbook) => {
        // Opened normally.
        let _ = workbook;
    }
    Err(Error::EncryptedWorkbook { .. }) => {
        // Prompt the user for a password (or instruct them to remove encryption in Excel).
        let info = detect_workbook_encryption(path)?
            .expect("encrypted workbook should be classified");

        match info.kind {
            WorkbookEncryptionKind::OoxmlOleEncryptedPackage => {
                // Encrypted OOXML (`EncryptionInfo` + `EncryptedPackage`).
            }
            WorkbookEncryptionKind::XlsFilepass => {
                // Encrypted legacy `.xls` (BIFF `FILEPASS`).
            }
            WorkbookEncryptionKind::UnknownOleEncrypted => {
                // OLE container appears encrypted but doesn't match known patterns.
            }
        }
    }
    Err(other) => return Err(other),
}
```

The intended end-state is that these same `formula-io` entrypoints will accept a user-provided
password (and optionally an “Agile HMAC verify” toggle) and return a normal workbook on success.
That plumbing is not implemented yet.

Notes:

- Passwords are treated as **UTF-8 strings at the API boundary** and encoded internally according
  to the relevant spec requirements (typically UTF-16LE for key derivation).
- Callers should avoid logging passwords or embedding them in error messages.

This `*_with_options` API is **planned**; today `formula-io` exposes `open_workbook` and
`open_workbook_model` only.

If you call the current APIs on an encrypted workbook, they will return
`Error::EncryptedWorkbook` until password support is integrated.

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
   - Surface as: `UnsupportedEncryptionScheme { scheme: … }`
   - UI action: explain limitation and suggest re-saving without encryption in Excel.

4. **Corrupt encrypted wrapper** (missing streams, malformed `EncryptionInfo`, truncated payload)
   - Surface as: `InvalidEncryptedContainer` / `CorruptFile`

These distinctions matter for UX and telemetry: “needs password” is a normal user workflow, while
“unsupported scheme” is an engineering coverage gap.

Until password-based decryption is implemented, Formula will generally surface a single
“encrypted workbook not supported” error instead of distinguishing the cases above.

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
  example `.xlsx`, `.xlsm`, `.xlsb`).
  These files are OLE/CFB wrappers (not ZIP/OPC), so they must not live under `fixtures/xlsx/`
  where the round-trip corpus is enumerated via `xlsx-diff::collect_fixture_paths`.
- Some encryption coverage is exercised with **synthetic** containers generated directly in tests
  (for example `crates/formula-io/tests/encrypted_ooxml.rs` and
  `crates/formula-io/tests/encrypted_xls.rs`).

---

## References (specs)

- **MS-OFFCRYPTO** — Office Document Cryptography Structure Specification  
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
- **MS-XLS** — Excel Binary File Format (`FILEPASS`, BIFF globals)  
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/
- **MS-XLSX** — Office Open XML SpreadsheetML Package Structure  
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsx/
