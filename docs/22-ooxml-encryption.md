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
    open APIs (`open_workbook_with_password`, `open_workbook_model_with_password`) can decrypt and
    open Agile (4.4) encrypted `.xlsx`/`.xlsm`/`.xlsb` in memory (via the `formula-xlsx` decryptor).
    (Standard/CryptoAPI `minor=2` encrypted workbooks are a different scheme; see
    `docs/offcrypto-standard-cryptoapi.md`.)
- **Streaming decrypt reader (does not validate `dataIntegrity` HMAC):**
  - `crates/formula-io/src/encrypted_ooxml.rs`
  - `crates/formula-io/src/encrypted_package_reader.rs`
  - Verifies the password and unwraps the package key, then decrypts `EncryptedPackage` on demand as
    a `Read + Seek` stream.
- **Agile (4.4) reference decryptor (includes `dataIntegrity` HMAC verification when present):**
  `crates/formula-xlsx/src/offcrypto/*`
- **End-to-end decrypt helpers + Agile writer (OLE wrapper → decrypted ZIP bytes):**
  `crates/formula-office-crypto`
  - Note: `formula-office-crypto` validates `dataIntegrity` (HMAC) when present and returns
    `OfficeCryptoError::IntegrityCheckFailed` on mismatch. If the `<dataIntegrity>` element is
    missing, it will still decrypt but cannot validate integrity.
- **MS-OFFCRYPTO parsing + decrypt helpers + low-level building blocks:** `crates/formula-offcrypto`
  - `parse_encryption_info`, `inspect_encryption_info` (`crates/formula-offcrypto/src/lib.rs`)
  - End-to-end decrypt helpers:
    - `decrypt_encrypted_package` (given `EncryptionInfo` + `EncryptedPackage` stream bytes;
      integrity verification is optional via `DecryptOptions.verify_integrity`)
    - `decrypt_ooxml_from_ole_bytes` (given raw OLE/CFB bytes; does not currently verify
      `dataIntegrity`)
  - Agile password verifier + secret key:
    - public helpers (`agile_verify_password`, `agile_secret_key`) in `crates/formula-offcrypto/src/lib.rs`
    - verifier/key primitives in `crates/formula-offcrypto/src/agile.rs`
  - Agile `EncryptedPackage` segment decryption + IV derivation
    (`crates/formula-offcrypto/src/encrypted_package.rs`, `agile_decrypt_package`)

## Supported vs unsupported (Agile 4.4 scope)

This document describes Formula’s support for the dominant modern scheme Excel uses today:
**Agile Encryption (version 4.4)** with password-based key encryption.

Standard/CryptoAPI (minor=2) encryption is a different scheme; see
`docs/offcrypto-standard-cryptoapi.md` and `docs/offcrypto-standard-encryptedpackage.md`.

### Supported

| Area | Supported |
|------|-----------|
| `EncryptionInfo` version | **Agile** `major=4, minor=4` |
| Package cipher (`keyData`) | **AES** + **CBC** (`cipherAlgorithm="AES"`, `cipherChaining="ChainingModeCBC"`) |
| Package key sizes | 128/192/256-bit (`keyBits` 128/192/256) |
| Hash algorithms | `SHA1`, `SHA256`, `SHA384`, `SHA512` (case-insensitive) |
| Key encryptor | **Password** key-encryptor only (`uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"`) |
| Integrity | `dataIntegrity` HMAC verification **when `<dataIntegrity>` is present** (algorithm documented below). Some real-world producers omit the `<dataIntegrity>` element; in that case `crates/formula-xlsx::offcrypto` and `crates/formula-office-crypto` will still decrypt `EncryptedPackage` but **skip integrity verification** (and `decrypt_agile_encrypted_package_with_warnings` can report `OffCryptoWarning::MissingDataIntegrity`). `formula-io`’s streaming decrypt reader does not currently validate `dataIntegrity`. |

### Explicitly unsupported (hard errors)

These inputs should fail with actionable “unsupported” errors (not “corrupt file”):

- **Certificate key-encryptor** (`…/keyEncryptor/certificate`) and other non-password key encryptors
- **Non-CBC chaining** (e.g. CFB, ECB) for either the package (`keyData`) or the password key-encryptor
- **Non-AES ciphers**
- **Extensible Encryption** and other `EncryptionInfo` versions besides `4.4`

#### Important implementation note (per-crate validation)

The MS-OFFCRYPTO XML descriptor contains many parameters. Different Formula crates enforce different
levels of validation:

- `crates/formula-xlsx::offcrypto` validates:
  - `cipherAlgorithm == AES`
  - `cipherChaining == ChainingModeCBC`
  - password key-encryptor present (certificate-only encryption fails with
    `formula_xlsx::offcrypto::OffCryptoError::UnsupportedKeyEncryptor { .. }`)
- `crates/formula-office-crypto` validates:
  - `cipherAlgorithm == AES`
  - `cipherChaining == ChainingModeCBC`
  - password key-encryptor present (certificate-only encryption is rejected; today this surfaces as
    `formula_office_crypto::OfficeCryptoError::InvalidFormat("missing password keyEncryptor")`)
- `crates/formula-offcrypto`’s Agile `EncryptionInfo` parser is also strict about the dominant
  Excel scheme:
  - it validates `keyData.cipherAlgorithm == AES` / `keyData.cipherChaining == ChainingModeCBC`
    (and the same attributes on the password `encryptedKey` element), returning
    `formula_offcrypto::OffcryptoError::UnsupportedAlgorithm(..)` on mismatch.
  - if **no password key-encryptor** is present (e.g. the file only contains certificate
    key-encryptors), it returns
    `formula_offcrypto::OffcryptoError::UnsupportedKeyEncryptor { available: Vec<String> }`.
    (If a password key-encryptor is present but its required `<encryptedKey>` data is missing, that
    remains a structural `InvalidEncryptionInfo`.)

If you add new decryption entrypoints on top of `formula-offcrypto`, make sure to validate the
declared cipher/chaining/key-encryptor parameters (or preserve `formula-offcrypto`’s
`UnsupportedAlgorithm` / `UnsupportedKeyEncryptor { available }` errors) so unsupported files don’t
get misreported as “corrupt” or “wrong password”.

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

Compatibility note: while Excel typically encodes the three password key-encryptor ciphertext blobs
(`encryptedVerifierHashInput`, `encryptedVerifierHashValue`, `encryptedKeyValue`) as **attributes**
on `<p:encryptedKey>`, MS-OFFCRYPTO also permits them to appear as **child elements** with base64
text content (e.g. `<p:encryptedKeyValue>...</p:encryptedKeyValue>`). Formula’s Agile parsers
tolerate either representation and prefer attribute values when both are present.

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

### Key derivation (`TruncateHash`) for verifier/key/HMAC purposes

Agile derives different AES keys from the iterated password hash by hashing with a purpose-specific
`blockKey` and then applying MS-OFFCRYPTO’s `TruncateHash`:

```text
derived = Hash(iterated_password_hash || blockKey)
key     = TruncateHash(derived, keyLenBytes)
```

`TruncateHash` detail that matters for interoperability:

- If `keyLenBytes <= hashLen`, take the prefix.
- If `keyLenBytes > hashLen`, **pad the digest with `0x36` bytes** to reach `keyLenBytes` (this
  matches Excel and `msoffcrypto-tool`).

This is easy to miss when `hashAlgorithm="SHA1"` (20-byte digest) but `keyBits="192"`/`"256"` (24/32
bytes).

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
  - See the in-repo Unicode-password fixtures:
    - `fixtures/encrypted/ooxml/agile-unicode.xlsx` (`pässwörd`, NFC)
    - `fixtures/encrypted/ooxml/agile-unicode-excel.xlsx` (`pässwörd🔒`, NFC, includes non-BMP emoji)
    - `fixtures/encrypted/ooxml/standard-unicode.xlsx` (`pässwörd🔒`, NFC, includes non-BMP emoji)
    and the regression tests in `crates/formula-io/tests/encrypted_ooxml_decrypt.rs`.
- **Whitespace is significant**: do not trim. If the password was set with leading/trailing
  whitespace, that whitespace is part of the UTF-16LE byte sequence used by the KDF, and a trimmed
  password will fail verification. (See `crates/formula-io/tests/encrypted_unicode_passwords.rs`.)

### Agile `blockKey` constants (must match spec exactly)

Agile uses a handful of fixed 8-byte `blockKey` constants. They are **not** derived from the file
and are easy to mistype; incorrect constants often yield “almost works” implementations that fail
on real Excel-produced files.

The canonical values (MS-OFFCRYPTO) are:

```text
VERIFIER_HASH_INPUT_BLOCK = FE A7 D2 76 3B 4B 9E 79
VERIFIER_HASH_VALUE_BLOCK = D7 AA 0F 6D 30 61 34 4E
KEY_VALUE_BLOCK           = 14 6E 0B E7 AB AC D0 D6
HMAC_KEY_BLOCK            = 5F B2 AD 01 0C B9 E1 F6
HMAC_VALUE_BLOCK          = A0 67 7F 02 B2 2C 84 33
```

Code references:

- `crates/formula-xlsx/src/offcrypto/crypto.rs`
- `crates/formula-office-crypto/src/agile.rs`

### Verifier check (wrong password vs continue)

To distinguish a wrong password from “corrupt/unsupported file”, Excel-style Agile encryption stores
an encrypted verifier:

- decrypt `encryptedVerifierHashInput`
- decrypt `encryptedVerifierHashValue`
- compute `Hash(verifierHashInput)` and compare to `verifierHashValue`

Notes:

- `encryptedVerifierHashValue` decrypts to a buffer that is typically **AES-block padded**
  (e.g. SHA-1’s 20-byte digest stored in a 32-byte decrypted buffer). Compare only the digest prefix
  (`hashSize` bytes / hash output length), not the full decrypted buffer.

If this comparison fails, we should return a **wrong password** error:

- at the decryption layer: `formula_offcrypto::OffcryptoError::InvalidPassword`
- at the API layer: `formula_io::Error::InvalidPassword`

### IV usage for verifier/key fields (important)

For the password key-encryptor fields (`encryptedVerifierHashInput`, `encryptedVerifierHashValue`,
`encryptedKeyValue`), the IV used by Excel-compatible implementations is:

> `IV = p:encryptedKey/@saltValue[..blockSize]`

i.e. the first AES block of `p:encryptedKey/@saltValue` (the attribute must be at least `blockSize`
bytes long).

Do **not** reuse the per-segment IV logic from `EncryptedPackage` here.

Compatibility note: some non-Excel producers appear to derive the IV as
`IV = Truncate(Hash(saltValue || blockKey), blockSize)` (similar to other Agile IV derivations).
To maximize real-world compatibility:

- `crates/formula-xlsx::offcrypto` and `crates/formula-office-crypto` will try both strategies in
  their main decrypt paths (treating a verifier mismatch as a signal to retry with the alternative
  IV derivation).
- `crates/formula-offcrypto` also tries both strategies anywhere it can validate the verifier
  hashes (e.g. `decrypt_encrypted_package` for Agile, and the `agile_verify_password` /
  `agile_secret_key` helpers). When verifier fields are missing, it falls back to the common
  `IV = saltValue[..blockSize]` behavior.

## Decrypting `EncryptedPackage`

The `EncryptedPackage` stream layout (on disk / as stored) is:

```text
8B   original_package_size (8-byte plaintext size prefix; see note below)
...  ciphertext bytes (AES-CBC, block-aligned)
```

Compatibility note: while the spec describes this as a `u64le`, some producers/libraries treat the
8-byte prefix as `u32 totalSize` + `u32 reserved` (often 0). To avoid truncation or “huge size”
misreads, parse it as `lo=u32le(bytes[0..4])`, `hi=u32le(bytes[4..8])`, then
`size = lo as u64 | ((hi as u64) << 32)`.

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

In Formula, validating this HMAC (when the metadata is available) is important for good error
semantics:

- “password wrong” should not surface as “ZIP is corrupt”
- file tampering/corruption should be detected even if the decrypted ZIP happens to parse

When `<dataIntegrity>` is absent, Formula can still decrypt the package, but cannot verify integrity
(and callers should treat the decrypted bytes as unauthenticated).

Implementation note:

- `crates/formula-xlsx::offcrypto` treats `<dataIntegrity>` as **optional**:
  - if present: Formula validates the HMAC as described below
  - if absent: Formula decrypts successfully but does **not** verify integrity (and
    `decrypt_agile_encrypted_package_with_warnings` can report `OffCryptoWarning::MissingDataIntegrity`)
- `crates/formula-office-crypto` currently requires `<dataIntegrity>` and treats a missing element as
  a malformed wrapper.
- `crates/formula-offcrypto` can validate `dataIntegrity` when decrypting with
  `formula_offcrypto::decrypt_encrypted_package` and `DecryptOptions.verify_integrity = true`
  (default: `false`).
- `formula-io`’s streaming decrypt reader does not currently validate `dataIntegrity`, so it may
  successfully open some inputs even when integrity metadata is missing or inconsistent.

### What bytes are authenticated (critical)

The HMAC is computed over the **raw bytes of the `EncryptedPackage` stream as stored**, i.e.:

> **HMAC target = `EncryptedPackage` stream bytes = header (8-byte size prefix) + ciphertext + padding**

Notably:

- It is **not** the HMAC of the decrypted ZIP bytes.
- It is **not** the HMAC of “ciphertext excluding the 8-byte header”.
- It includes any final block padding bytes present in the `EncryptedPackage` stream.

This detail has been the source of prior incorrect implementations.

Compatibility note: some non-Excel producers have been observed to compute `dataIntegrity` over
different byte ranges (e.g. ciphertext-only, or the decrypted plaintext ZIP) instead of the
spec/Excel target (`EncryptedPackage` stream bytes). To maximize interoperability with real-world
files (including our committed fixture corpus), Formula’s decryptors are currently permissive:

- `crates/formula-xlsx::offcrypto` accepts:
  - HMAC over the full `EncryptedPackage` stream bytes (8-byte size header + ciphertext + padding)
  - fallback: HMAC over the decrypted package bytes (plaintext ZIP)
- `crates/formula-office-crypto` accepts:
  - HMAC over the full `EncryptedPackage` stream bytes (8-byte size header + ciphertext + padding)
  - HMAC over ciphertext only (excludes the 8-byte size header)
  - HMAC over plaintext only (decrypted ZIP bytes)
  - HMAC over (8-byte size header + plaintext ZIP bytes)
- `crates/formula-offcrypto` accepts:
  - HMAC over the full `EncryptedPackage` stream bytes (8-byte size header + ciphertext + padding)
    when integrity verification is enabled (`DecryptOptions.verify_integrity = true`); it does not
    currently include fallback HMAC target variants.

Writers should follow MS-OFFCRYPTO/Excel (authenticate the stream bytes). If we ever decide to make
`formula-office-crypto` strict, update this section and the corresponding compatibility tests.

### High-level integrity algorithm (Agile)

1. Obtain the **package key** (by password verification + decrypting `encryptedKeyValue`)
2. Decrypt `encryptedHmacKey` and `encryptedHmacValue` using the **package key**
   - IVs are derived from `keyData/@saltValue` and constant blocks:
      - HMAC key block: `5F B2 AD 01 0C B9 E1 F6`
      - HMAC value block: `A0 67 7F 02 B2 2C 84 33`
   - The decrypted `hmacKey` / `hmacValue` buffers can include AES-CBC block padding; use only the
     digest-length prefix when computing/comparing.
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
| `formula_offcrypto::OffcryptoError::IntegrityCheckFailed` | HMAC mismatch (tampering/corruption) when integrity verification is enabled | Re-download file; if persistent, treat as corrupted |
| `formula_office_crypto::OfficeCryptoError::IntegrityCheckFailed` | HMAC mismatch (tampering/corruption) | Re-download file; if persistent, treat as corrupted |
| `formula_io::Error::UnsupportedOoxmlEncryption` / `formula_offcrypto::OffcryptoError::UnsupportedVersion` | `EncryptionInfo` version not recognized | Re-save without encryption; or add support |
| `formula_offcrypto::OffcryptoError::UnsupportedEncryption { encryption_type: ... }` | Encryption type known but not supported by selected decrypt mode | Use correct decrypt mode / add support |
| `formula_offcrypto::OffcryptoError::UnsupportedKeyEncryptor { available }` | Agile encryption is present, but no password key-encryptor is present (e.g. certificate-only protection) | Re-save using password encryption; or add key-encryptor support |
| `formula_offcrypto::OffcryptoError::UnsupportedAlgorithm(..)` | Agile cipher/chaining parameters are not in our supported subset (e.g. non-AES / non-CBC) | Re-save with default Excel encryption settings; or add support |
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
   - `crates/formula-offcrypto` provides MS-OFFCRYPTO parsing plus end-to-end decrypt helpers (e.g.
     `decrypt_encrypted_package`, `decrypt_ooxml_from_ole_bytes`).
     - Integrity verification is optional when using `decrypt_encrypted_package`
       (`DecryptOptions.verify_integrity`). Other helper APIs do not currently verify `dataIntegrity`.
     - It does not include the more permissive HMAC-target fallbacks found in the higher-level
       decryptors.

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
  empty-password + Unicode-password cases, macro-enabled `.xlsm` fixtures, and `*-large.xlsx`
  multi-segment fixtures). See `fixtures/encrypted/ooxml/README.md` for passwords, provenance, and
  regeneration notes.

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
