# OOXML Password Decryption (MS-OFFCRYPTO Agile 4.4)

This document is a **developer-facing reference** for how Formula handles **Excel “Encrypt with
Password”** (password-to-open) OOXML files (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xlam`, `.xlsb`).

It is intentionally specific about:

- what we **support** vs **reject**
- which bytes are authenticated (HMAC target)
- which salts/IVs are used where (common source of subtle bugs)

The goal is to prevent regressions where “almost right” implementations pass basic tests but fail on
real Excel-produced files.

## Container format: OLE/CFB wrapper

Password-encrypted OOXML files are **not ZIP files on disk**. Excel stores them as an **OLE/CFB
(Compound File Binary)** container with (at least) these streams:

- `EncryptionInfo` — encryption parameters (version header + XML descriptor for Agile)
- `EncryptedPackage` — the encrypted bytes of the real workbook package

Spec: MS-OFFCRYPTO “Encrypted OOXML File” / “EncryptionInfo Stream” / “EncryptedPackage Stream”.

## Where this lives in Formula (implementation map)

This doc is intentionally “close to the metal”. Helpful entrypoints in this repo:

- **User-facing detection / error semantics:** `crates/formula-io/src/lib.rs`
  - `Error::PasswordRequired`
  - `Error::InvalidPassword`
  - `Error::UnsupportedOoxmlEncryption`
  - Note: with the `formula-io` crate feature **`encrypted-workbooks`** enabled, the password-aware
    open APIs (`open_workbook_with_password`, `open_workbook_model_with_password`) can also decrypt
    and open Agile (4.4) encrypted `.xlsx`/`.xlsm` in memory (via the `formula-xlsx` decryptor).
- **Agile (4.4) reference decryptor (includes `dataIntegrity` HMAC verification):**
  `crates/formula-xlsx/src/offcrypto/*`
- **End-to-end decrypt helpers + Agile writer (OLE wrapper → decrypted ZIP bytes):**
  `crates/formula-office-crypto`
  - Note: `formula-office-crypto` also validates `dataIntegrity` (HMAC) and returns
    `OfficeCryptoError::IntegrityCheckFailed` on mismatch.
- **MS-OFFCRYPTO parsing + low-level building blocks:** `crates/formula-offcrypto`
  - `parse_encryption_info`, `inspect_encryption_info` (`crates/formula-offcrypto/src/lib.rs`)
  - Agile password verifier + secret key (`crates/formula-offcrypto/src/agile.rs`)
  - Agile `EncryptedPackage` segment decryption + IV derivation
    (`crates/formula-offcrypto/src/encrypted_package.rs`, `agile_decrypt_package`)

## Supported vs unsupported (current scope)

Formula’s OOXML decryption support is intentionally scoped to the dominant real-world scheme Excel
uses today: **Agile Encryption (version 4.4) with password-based key encryption**.

### Supported

| Area | Supported |
|------|-----------|
| `EncryptionInfo` version | **Agile** `major=4, minor=4` |
| Package cipher (`keyData`) | **AES** + **CBC** (`cipherAlgorithm="AES"`, `cipherChaining="ChainingModeCBC"`) |
| Package key sizes | 128/192/256-bit (`keyBits` 128/192/256) |
| Hash algorithms | `SHA1`, `SHA256`, `SHA384`, `SHA512` (case-insensitive) |
| Key encryptor | **Password** key-encryptor only (`uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"`) |
| Integrity | `dataIntegrity` HMAC verification (implemented in `crates/formula-xlsx::offcrypto` and `crates/formula-office-crypto`; algorithm documented below) |

### Explicitly unsupported (hard errors)

These inputs should fail with actionable “unsupported” errors (not “corrupt file”):

- **Certificate key-encryptor** (`…/keyEncryptor/certificate`) and other non-password key encryptors
- **Non-CBC chaining** (e.g. CFB, ECB) for either the package (`keyData`) or the password key-encryptor
- **Non-AES ciphers**
- **Extensible Encryption** and other `EncryptionInfo` versions besides `4.4`

#### Important implementation note (per-crate validation)

The MS-OFFCRYPTO XML descriptor contains many parameters. Different Formula crates enforce different
levels of validation:

- `crates/formula-xlsx::offcrypto` and `crates/formula-office-crypto` validate:
  - `cipherAlgorithm == AES`
  - `cipherChaining == ChainingModeCBC`
  - password key-encryptor present (and will reject certificate-only encryption with an explicit
    “unsupported key encryptor” error)
- `crates/formula-offcrypto`’s Agile parser currently **ignores** `cipherAlgorithm` /
  `cipherChaining` attributes (it assumes AES-CBC) and will treat missing password key-encryptor
  data as a structural error (e.g. “missing password `<encryptedKey>` element”).

If you add new decryption entrypoints on top of `formula-offcrypto`, make sure to validate the
declared cipher/chaining/key-encryptor parameters so unsupported files don’t get misreported as
“corrupt” or “wrong password”.

## The `EncryptionInfo` stream (Agile 4.4)

The `EncryptionInfo` stream begins with an 8-byte `EncryptionVersionInfo` header:

```text
u16 majorVersion (LE)
u16 minorVersion (LE)
u32 flags        (LE)
```

For Agile encryption this is typically `4.4` with `flags=0x00000040`, followed by an XML document
(`<?xml …?><encryption …>…</encryption>`).

See also: `crates/formula-io/src/bin/ooxml-encryption-info.rs` (a small helper that prints the
version header and sniffs the XML root tag).

Example:

```bash
bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- path/to/encrypted.xlsx
```

```text
Agile (4.4) flags=0x00000040 xml_root=encryption
```

Note: the spec says the XML follows immediately after the 8-byte version header, but some
non-Excel producers have been observed to include extra bytes (e.g. a 4-byte length prefix or
UTF-16LE XML). `crates/formula-offcrypto` includes conservative heuristics to handle these cases
(`parse_agile_encryption_info_xml`).

### Key point: there are *two* parameter sets

The Agile XML descriptor contains two different “parameter sets” that are easy to mix up:

1. `<keyData …>` — parameters for **encrypting the workbook package** (`EncryptedPackage`)
2. `<p:encryptedKey …>` — parameters for **deriving keys from the user password** and decrypting the
   **package key** + verifier fields

Representative shape (attributes elided):

```xml
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData ... saltValue="..."/>
  <dataIntegrity encryptedHmacKey="..." encryptedHmacValue="..."/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey ... saltValue="..."
                      encryptedVerifierHashInput="..."
                      encryptedVerifierHashValue="..."
                      encryptedKeyValue="..."/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>
```

When debugging decryption bugs, always ask:

- “Am I using the `keyData` salt/hash, or the `p:encryptedKey` salt/hash?”

They are **not interchangeable**.

## Password KDF and package key decryption

The password processing pipeline is:

1. **Hash the password** (salt + spin count) using the password key-encryptor parameters
2. **Derive 3 AES keys** from that hash (verifier input, verifier value, package key)
3. **Decrypt verifier fields** and compare (WrongPassword vs continue)
4. **Decrypt `encryptedKeyValue`** → yields the **package key** used for `EncryptedPackage`

These primitives (password hashing, verifier checks, and secret-key extraction) are implemented in:

- `crates/formula-offcrypto/src/agile.rs`

### Password hashing (Agile)

MS-OFFCRYPTO defines the password hash as:

```text
pw = UTF-16LE(password)          (no BOM, no terminator)
H0 = Hash(saltValue || pw)
Hi = Hash(LE32(i) || H(i-1))     for i in 0..spinCount-1
```

Notes:

- The `saltValue` and `spinCount` come from `<p:encryptedKey ...>`, not `<keyData>`.
- “Hash” is the algorithm named by `p:encryptedKey/@hashAlgorithm`.
- `password` is taken **as-is** from the caller. No Unicode normalization is applied before UTF-16LE
  encoding, so different normalization forms (e.g. NFC vs NFD) will yield different hashes and
  therefore different derived keys. See the in-repo Unicode-password fixture
  `fixtures/encrypted/ooxml/agile-unicode.xlsx` and the regression tests in
  `crates/formula-io/tests/encrypted_ooxml_decrypt.rs`.

### Password edge cases (do not accidentally change semantics)

- **Empty password is valid**: Excel can encrypt a workbook with an empty open password (`""`).
  Treat this as a real password value (UTF-16LE empty byte string), distinct from “no password was
  provided”.
  - In `formula-io`, this typically means callers must distinguish `None` (prompt user) vs
    `Some("")` (attempt empty-password decrypt).
- **Unicode normalization matters**: the KDF operates on the exact UTF-16LE byte sequence. Different
  Unicode normalization forms (e.g. NFC vs NFD) produce different derived keys, even if the
  displayed password looks identical.
  - This is a UI/input policy decision: Formula’s crypto expects the *exact* password string that
    was used at encryption time.
  - For debugging a “wrong password” report with a Unicode password, try NFC vs NFD variants before
    assuming the file is corrupted.

### Verifier check (wrong password vs continue)

To distinguish a wrong password from “corrupt/unsupported file”, Excel-style Agile encryption stores
an encrypted verifier:

- decrypt `encryptedVerifierHashInput`
- decrypt `encryptedVerifierHashValue`
- compute `Hash(verifierHashInput)` and compare to `verifierHashValue`

If this comparison fails, we should return a **wrong password** error:

- at the decryption layer: `formula_offcrypto::OffcryptoError::InvalidPassword`
- at the API layer: `formula_io::Error::InvalidPassword`

### IV usage for verifier/key fields (important)

For the password key-encryptor fields (`encryptedVerifierHashInput`, `encryptedVerifierHashValue`,
`encryptedKeyValue`), the IV used by Excel-compatible implementations is the `p:encryptedKey/@saltValue`
(padded/truncated to the AES block size).

Do **not** reuse the per-segment IV logic from `EncryptedPackage` here.

Compatibility note: some non-Excel producers appear to derive the IV as
`IV = Truncate(Hash(saltValue || blockKey), blockSize)` (similar to other Agile IV derivations).
To maximize real-world compatibility, `crates/formula-xlsx::offcrypto` will try both strategies
(treating a verifier mismatch as a signal to retry with the alternative IV derivation). In
contrast, `crates/formula-office-crypto` uses `saltValue` directly as the IV.

## Decrypting `EncryptedPackage`

The `EncryptedPackage` stream layout (on disk / as stored) is:

```text
8B   original_package_size (u64 little-endian, *not encrypted*)
...  ciphertext bytes (AES-CBC, block-aligned)
```

Decryption rules:

- Plaintext is processed in **4096-byte segments** (except the last).
- Each segment is AES-CBC-encrypted **independently** using:
  - the **package key** (from `encryptedKeyValue`)
  - a **per-segment IV** derived from `keyData/@saltValue` and the segment index

### Segment IV derivation (the common pitfall)

For segment `i` (0-based), derive:

```text
iv_i = Truncate(keyData/@blockSize, Hash(keyData/@saltValue || LE32(i)))
```

where `Hash` is `keyData/@hashAlgorithm`.

This is implemented in `crates/formula-offcrypto/src/encrypted_package.rs`
(`agile_decrypt_package`).

**Gotcha:** it is easy to accidentally use `p:encryptedKey/@saltValue` here. That will “decrypt” to
garbage that often looks like a corrupt ZIP.

### Truncation (padding is not PKCS#7)

Do not rely on PKCS#7 unpadding. The ciphertext is block-aligned, and the plaintext semantic length
is the `original_package_size` stored in the first 8 bytes. After decrypting segments, truncate the
output to that length.

## Integrity verification (`dataIntegrity` HMAC)

Agile encryption can include a package-level integrity check via:

```xml
<dataIntegrity encryptedHmacKey="..." encryptedHmacValue="..."/>
```

This is **not optional** for correctness when we want good error semantics:

- “password wrong” should not surface as “ZIP is corrupt”
- file tampering/corruption should be detected even if the decrypted ZIP happens to parse

### What bytes are authenticated (critical)

The HMAC is computed over the **raw bytes of the `EncryptedPackage` stream as stored**, i.e.:

> **HMAC target = `EncryptedPackage` stream bytes = header (u64 size) + ciphertext + padding**

Notably:

- It is **not** the HMAC of the decrypted ZIP bytes.
- It is **not** the HMAC of “ciphertext excluding the 8-byte header”.
- It includes any final block padding bytes present in the `EncryptedPackage` stream.

This detail has been the source of prior incorrect implementations.

Compatibility note: some non-Excel producers have been observed to compute `dataIntegrity` over the
**decrypted package bytes** (plaintext ZIP) instead of the `EncryptedPackage` stream bytes. For
Excel parity, new implementations should follow the spec (authenticate the stream bytes). However,
`crates/formula-xlsx::offcrypto` is permissive and will accept either target, while
`crates/formula-office-crypto` requires the spec/Excel behavior (stream bytes).

### High-level integrity algorithm (Agile)

1. Obtain the **package key** (by password verification + decrypting `encryptedKeyValue`)
2. Decrypt `encryptedHmacKey` and `encryptedHmacValue` using the **package key**
   - IVs are derived from `keyData/@saltValue` and constant blocks:
     - HMAC key block: `5F B2 AD 01 0C B9 E1 F6`
     - HMAC value block: `A0 67 7F 02 B2 2C 84 33`
3. Compute:

```text
actual = HMAC(key = hmacKey, hash = keyData/@hashAlgorithm, data = EncryptedPackageStreamBytes)
```

4. Compare `actual` to the decrypted `hmacValue`
   - mismatch ⇒ treat as an **integrity failure** (ideally surfaced distinctly; `formula-io` currently
     maps this to `Error::InvalidPassword` for UX)

## Common errors and what they mean

The Agile decryption errors are designed to be actionable. The most important distinctions:

| Error | Meaning | Typical user action |
|------|---------|---------------------|
| `formula_io::Error::PasswordRequired` | Encrypted OOXML detected, but no password provided | Prompt for password |
| `formula_io::Error::InvalidPassword` | Wrong password **or** integrity mismatch (Agile HMAC); callers should treat this as “password incorrect or file corrupted/tampered”. | Retry password; if persistent, treat as corrupted/tampered |
| `formula_offcrypto::OffcryptoError::InvalidPassword` / `formula_xlsx::offcrypto::OffCryptoError::WrongPassword` | Password verifier mismatch | Retry password |
| `formula_xlsx::offcrypto::OffCryptoError::IntegrityMismatch` | HMAC mismatch (tampering/corruption) | Re-download file; if persistent, treat as corrupted |
| `formula_office_crypto::OfficeCryptoError::IntegrityCheckFailed` | HMAC mismatch (tampering/corruption) | Re-download file; if persistent, treat as corrupted |
| `formula_io::Error::UnsupportedOoxmlEncryption` / `formula_offcrypto::OffcryptoError::UnsupportedVersion` | `EncryptionInfo` version not recognized | Re-save without encryption; or add support |
| `formula_offcrypto::OffcryptoError::UnsupportedEncryption { encryption_type: ... }` | Encryption type known but not supported by selected decrypt mode | Use correct decrypt mode / add support |
| `formula_xlsx::offcrypto::OffCryptoError::UnsupportedKeyEncryptor { .. }` | File is encrypted, but only a non-password key-encryptor (e.g. certificate) is present | Re-save using password encryption; or add key-encryptor support |
| `formula_xlsx::offcrypto::OffCryptoError::{UnsupportedCipherAlgorithm, UnsupportedCipherChaining, UnsupportedHashAlgorithm}` | Cipher/chaining/hash params are not in our supported subset | Re-save with default Excel encryption settings; or add support |
| XML/structure errors (`InvalidEncryptionInfo`, base64 decode, ciphertext alignment) | Malformed/corrupt encrypted wrapper | Treat as corrupted file |

Note: at the `formula-io` API boundary we currently expose only `PasswordRequired` / `InvalidPassword`
for password-to-open OOXML. Integrity failures from deeper layers (HMAC mismatch) may be mapped to
`formula_io::Error::InvalidPassword`. If you need to distinguish “wrong password” vs “integrity
failure”, use `crates/formula-xlsx::offcrypto` or `crates/formula-office-crypto` directly (or plumb a
dedicated error variant through `formula-io`).

For the exact user-facing strings, see:

- `crates/formula-io/src/lib.rs` (`formula_io::Error`)
- `crates/formula-offcrypto/src/lib.rs` (`formula_offcrypto::OffcryptoError`)
- `crates/formula-xlsx/src/offcrypto/error.rs` (`formula_xlsx::offcrypto::OffCryptoError`)
- `crates/formula-office-crypto/src/error.rs` (`formula_office_crypto::OfficeCryptoError`)

### Example message strings (today)

Strings change over time, but these are the common ones you’ll see in logs/bug reports:

- Wrong password (Agile verifier mismatch):  
  `wrong password for encrypted workbook (verifier mismatch)`
- Integrity failure (HMAC mismatch):  
  `encrypted workbook integrity check failed (HMAC mismatch); the file may be corrupted or the password is incorrect`

## Debugging workflow (triage cookbook)

When investigating a bug report (“Formula can’t open my password-protected `.xlsx`”), it helps to
follow a consistent flow:

1. **Classify the file quickly**
   - Run:

     ```bash
     bash scripts/cargo_agent.sh run -p formula-io --bin ooxml-encryption-info -- path/to/file.xlsx
     ```

   - Results:
     - `Agile (4.4)` ⇒ this document applies.
     - `Standard (*.2)` ⇒ see `docs/offcrypto-standard-cryptoapi.md` +
       `docs/offcrypto-standard-encryptedpackage.md`.
     - `Extensible (*.3)` or `Unknown` ⇒ unsupported today.

2. **Validate the password with an external oracle**
   - Using Python `msoffcrypto-tool`:

     ```bash
     msoffcrypto-tool -p 'password' encrypted.xlsx decrypted.zip
     unzip -l decrypted.zip | head
     ```

   - If `msoffcrypto-tool` can’t decrypt with the user’s password, the issue is likely the password
     itself (or file corruption), not Formula.

3. **Compare Formula’s decryptors**
   - `crates/formula-xlsx::offcrypto` and `crates/formula-office-crypto` both implement Agile
     decryption + `dataIntegrity` validation.
   - `crates/formula-offcrypto` provides parsing and low-level building blocks, but is intentionally
     not yet a full end-to-end Agile decryptor.

4. **Interpret errors**
   - `WrongPassword` / `InvalidPassword`: verifier mismatch ⇒ password wrong (or normalization
     mismatch).
   - `IntegrityMismatch` / `IntegrityCheckFailed`: HMAC mismatch ⇒ tampering/corruption, or (if the
     password is known correct) an implementation bug. Re-check:
     - HMAC target bytes are the **raw `EncryptedPackage` stream bytes** (including the 8-byte
       size header and any padding).
     - Segment IV derivation uses `keyData/@saltValue` + LE32(segment_index).

## Test oracles / cross-validation

When validating Agile encryption/decryption correctness, we rely on external “known good”
implementations as oracles:

- **`msoffcrypto-tool`** (Python): can decrypt Excel-encrypted OOXML OLE containers.
  - Example: `msoffcrypto-tool -p 'password' encrypted.xlsx decrypted.zip`
- **`ms-offcrypto-writer`** (Rust): can generate Agile-encrypted OOXML containers compatible with
  Excel.
  - Notable real-world quirk: Excel (and `ms-offcrypto-writer`) use an HMAC key whose length matches
    the hash digest length (e.g. **64 bytes for SHA-512**) even when `keyData/@saltSize` is 16.

These tools are useful for answering “is our implementation wrong, or is the file weird?” quickly.

Repository fixtures:

- `fixtures/encrypted/ooxml/` contains synthetic encrypted OOXML workbooks (Agile + Standard, plus
  empty-password + Unicode-password cases). See `fixtures/encrypted/ooxml/README.md` for passwords,
  provenance, and regeneration notes.

## References (MS-OFFCRYPTO sections)

Microsoft Open Specifications (entrypoint):

- MS-OFFCRYPTO — Office Document Cryptography Structure  
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/

Deep links used frequently during development:

- EncryptionInfo Stream (Agile header + XML)  
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/87020a34-e73f-4139-99bc-bbdf6cf6fa55
- EncryptedPackage Stream  
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/b60c8b35-2db2-4409-8710-59d88a793f83
- Agile password verification / key derivation  
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/a57cb947-554f-4e5e-b150-3f2978225e92
- Data integrity (HMAC)  
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/63d9c262-82b9-4fa3-a06d-d087b93e3b00
