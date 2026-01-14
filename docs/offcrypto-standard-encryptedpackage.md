# MS-OFFCRYPTO Standard/CryptoAPI AES: `EncryptedPackage` decryption notes (AES-ECB)

This repo detects password-protected / encrypted OOXML workbooks as an **OLE/CFB** container with
`EncryptionInfo` + `EncryptedPackage` streams (MS-OFFCRYPTO).

High-level behavior in `formula-io`:

- Encrypted OOXML wrappers (`EncryptionInfo` + `EncryptedPackage`) are detected and surfaced via
  dedicated errors (`PasswordRequired` / `InvalidPassword` / `UnsupportedOoxmlEncryption`) so callers
  can prompt for a password and route “unsupported encryption” reports correctly.
- Without the `formula-io` cargo feature **`encrypted-workbooks`**, encrypted OOXML containers surface
  `Error::UnsupportedEncryption` (and `Error::UnsupportedOoxmlEncryption` for unknown/invalid
  `EncryptionInfo` versions).
- With **`encrypted-workbooks`** enabled:
  - `open_workbook(..)` / `open_workbook_model(..)` surface `Error::PasswordRequired` when no
    password is provided.
  - The password-aware entrypoints `open_workbook_with_password` /
    `open_workbook_model_with_password` can decrypt and open both **Agile (4.4)** and
    **Standard/CryptoAPI** (minor=2; commonly `3.2`/`4.2`) encrypted `.xlsx`/`.xlsm`.
    - Encrypted `.xlsb` currently surfaces `Error::UnsupportedEncryptedWorkbookKind { kind: "xlsb" }`.
  - `open_workbook_with_options` can also decrypt and open encrypted OOXML wrappers when a password
    is provided (typically returns `Workbook::Xlsx`; Standard AES may return `Workbook::Model`).
  - A streaming decrypt reader exists in `crates/formula-io/src/encrypted_ooxml.rs` +
    `crates/formula-io/src/encrypted_package_reader.rs`, but the high-level `open_workbook*` APIs
    currently decrypt to in-memory buffers.

Standard/CryptoAPI decryption primitives also exist in lower-level crates (notably
`crates/formula-offcrypto` and `crates/formula-office-crypto`). The high-level open path in
`formula-io` is still evolving (especially around streaming and cross-crate API consolidation); see
[`docs/21-encrypted-workbooks.md`](./21-encrypted-workbooks.md).

This document focuses on the `EncryptedPackage` stream itself, because the most common interop bugs
cluster around:

- the 8-byte plaintext length prefix,
- AES block alignment requirements,
- and **padding/truncation** (do **not** PKCS#7-unpad; always truncate to the declared size).

For Agile (4.4) OOXML decryption details (and `dataIntegrity` gotchas), see
[`docs/22-ooxml-encryption.md`](./22-ooxml-encryption.md).

See also:

- `docs/offcrypto-standard-cryptoapi.md` (Standard key derivation + verifier validation)
- `docs/offcrypto-standard-cryptoapi-rc4.md` (Standard CryptoAPI RC4 notes; different block size/keying)

## Implementation references in this repo

- **`formula-io` `EncryptedPackage` decryptors**:
  - `crates/formula-io/src/offcrypto/encrypted_package.rs`
    - `decrypt_encrypted_package_standard_aes_to_writer` (streaming Standard AES-ECB; no IV)
    - `decrypt_standard_encrypted_package_stream` (buffered; also attempts a non-standard segmented fallback when a salt is available)
- **`formula-offcrypto` Standard AES-ECB helper**:
  - `crates/formula-offcrypto/src/lib.rs`: `decrypt_encrypted_package_ecb`
- **More permissive Standard decryptor (handles additional variants)**:
  - `crates/formula-office-crypto/src/standard.rs`

## Normative spec references (MS-OFFCRYPTO)

* **`\\EncryptedPackage` stream layout**: MS-OFFCRYPTO **§2.3.4.4** “`\\EncryptedPackage` Stream”.
  * Defines `StreamSize` and notes the stream can be larger than `StreamSize` due to block padding.
* **Standard Encryption metadata (`\\EncryptionInfo`)**: MS-OFFCRYPTO **§2.3.4.5** “`\\EncryptionInfo`
  Stream (Standard Encryption)”.
* **Salt location/size**: MS-OFFCRYPTO **§2.3.4.7** “ECMA-376 Document Encryption Key Generation
  (Standard Encryption)”.
  * Salt is 16 bytes and stored in `EncryptionVerifier.Salt`.

## `EncryptedPackage` stream layout

`EncryptedPackage` is an OLE stream with:

1. `orig_size` (8-byte plaintext size prefix; `StreamSize` in the spec): decrypted package size in
   bytes.
2. `ciphertext: [u8]` (`EncryptedData` in the spec): encrypted bytes of the underlying OPC package
   (the `.xlsx` ZIP bytes), padded to a cipher block boundary.

Compatibility note: while MS-OFFCRYPTO describes the 8-byte prefix as a `u64le`, some
producers/libraries treat it as `u32 totalSize` + `u32 reserved` (often 0). To be compatible, parse
as two little-endian DWORDs and recombine:

```text
lo = u32le(prefix[0..4])
hi = u32le(prefix[4..8])
orig_size = lo as u64 | ((hi as u64) << 32)
```

Spec note (MS-OFFCRYPTO §2.3.4.4): the *physical* stream length can be **larger** than `orig_size`
because the encrypted data is padded to a cipher block boundary.

## AES decryption (baseline: AES-ECB, no IV)

Decrypt the ciphertext bytes (everything after the 8-byte size prefix) with **AES-ECB(key)**:

* AES key: derived from the password and `EncryptionInfo` (see `docs/offcrypto-standard-cryptoapi.md`).
* Mode: ECB (no IV).
* Padding: none at the crypto layer (ciphertext is block-aligned).
* Ciphertext length must be a **multiple of 16** bytes.

Implementation pointers:

* `crates/formula-offcrypto/src/lib.rs`: `decrypt_encrypted_package_ecb`
* `crates/formula-io/src/offcrypto/encrypted_package.rs`: `decrypt_encrypted_package_standard_aes_to_writer`

## Optional segmentation (0x1000)

Many decryptors process `EncryptedPackage` in **0x1000-byte (4096) plaintext segments** for
streaming/bounded memory:

```text
segmentSize = 0x1000   // 4096
aesBlock    = 16
```

For the **AES-ECB baseline**, segment boundaries are not cryptographically meaningful: ECB has no IV
and no chaining, so you may decrypt in any chunk sizes as long as they are multiples of 16 bytes.

### Non-standard segmented fallback (seen in some producers)

Some non-Excel producers encrypt `EncryptedPackage` as **0x1000-byte segments** using **AES-CBC**
with a per-segment IV derived from the verifier salt and segment index:

```text
iv_i = SHA1(salt || LE32(i))[0..16]
```

This is **not** the Excel-default Standard AES scheme (Excel uses AES-ECB), but it is common enough
that `formula-io`’s `decrypt_standard_encrypted_package_stream` attempts it as a fallback when a
salt is available.

## Padding + truncation (do not trust PKCS#7)

Do **not** treat decrypted bytes as PKCS#7 and “unpad”.

Instead:

1. Decrypt ciphertext (in the correct mode).
2. **Truncate the plaintext to `orig_size`**.

Example (real fixture in this repo):

* `fixtures/encrypted/ooxml/standard.xlsx` decrypts to a 3179-byte ZIP (`orig_size = 3179`), but the
  `EncryptedPackage` ciphertext bytes are padded to **4096 bytes** (plus the 8-byte size prefix) to
  accommodate OLE/producer quirks.
* Correct decryption therefore requires: decrypt all 4096 ciphertext bytes, then truncate the
  plaintext to 3179 bytes.

Rationale:

* MS-OFFCRYPTO §2.3.4.4 defines `StreamSize`/`orig_size` as authoritative for the unencrypted size,
  and explicitly allows the encrypted stream to be larger due to block-size padding.

## Validation rules (recommended)

These checks catch most corruption / truncation issues early:

* Stream length must be `>= 8` (need `orig_size` prefix).
* `orig_size` must fit in your address space / target type (e.g., `usize`).
* Let `ct = stream[8..]`. Require `ct.len() % 16 == 0` (AES block alignment).
* After decryption, if produced plaintext is `< orig_size`, treat as truncated/corrupt; otherwise
  truncate to `orig_size` and continue.
