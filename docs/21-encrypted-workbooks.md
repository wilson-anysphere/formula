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

---

## Public API: supplying passwords

### `formula-io` API

When password-based decryption is implemented, callers will supply an *open password* via
`OpenOptions.password`:

```rust
use formula_io::{open_workbook_with_options, OpenOptions};

let workbook = open_workbook_with_options(
    "book.xlsx",
    OpenOptions {
        password: Some("correct horse battery staple".into()),
        ..Default::default()
    },
)?;
```

Notes:

- Passwords are treated as **UTF-8 strings at the API boundary** and encoded internally according
  to the relevant spec requirements (typically UTF-16LE for key derivation).
- Callers should avoid logging passwords or embedding them in error messages.

If you call the current APIs (`open_workbook` / `open_workbook_model`) on an encrypted workbook,
they will return `Error::EncryptedWorkbook` until password support is integrated.

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

- Encrypted/password-protected OOXML workbook fixtures live under `fixtures/encrypted/ooxml/`.
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
