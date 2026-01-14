# MS-OFFCRYPTO Standard/CryptoAPI AES: `EncryptedPackage` decryption notes

This repo detects password-protected / encrypted OOXML workbooks as an OLE/CFB container with
`EncryptionInfo` + `EncryptedPackage` streams.

In `formula-io`, attempting to open these files without a password will surface `Error::PasswordRequired`.
The password-aware helpers (`open_workbook_with_password` / `open_workbook_model_with_password`) can
surface `Error::InvalidPassword`. (End-to-end decryption is still being wired; see
[`docs/21-encrypted-workbooks.md`](./21-encrypted-workbooks.md) for current behavior and entrypoints.)

For unknown `EncryptionInfo` versions, `formula-io` may surface
`Error::UnsupportedOoxmlEncryption { version_major, version_minor }`.

Standard (CryptoAPI) decryption support exists in this repo (notably in
`crates/formula-office-crypto` and `crates/formula-offcrypto`), but not all `formula-io` open paths
plumb it end-to-end yet (see [`docs/21-encrypted-workbooks.md`](./21-encrypted-workbooks.md)).

The most common interoperability bugs are in the `EncryptedPackage` stream framing (the `u64` size
prefix) and **padding/truncation**. This note is meant as a compact developer reference.

## Normative spec references (MS-OFFCRYPTO)

* **`\\EncryptedPackage` stream layout**: MS-OFFCRYPTO **§2.3.4.4** “`\\EncryptedPackage` Stream”.
  * Defines `StreamSize` and notes the stream can be larger than `StreamSize` due to block padding.
* **Standard Encryption metadata (`\\EncryptionInfo`)**: MS-OFFCRYPTO **§2.3.4.5** “`\\EncryptionInfo`
  Stream (Standard Encryption)”.
* **Salt location/size**: MS-OFFCRYPTO **§2.3.4.7** “ECMA-376 Document Encryption Key Generation
  (Standard Encryption)”.
  * Salt is 16 bytes and stored in `EncryptionVerifier.Salt`.
Note: Agile encryption (4.4) uses 4096-byte segmenting + per-segment IVs. Standard/CryptoAPI AES
`EncryptedPackage` decryption is **not entirely uniform in the wild**: the baseline MS-OFFCRYPTO
scheme uses **AES-ECB** (no IV), but some producers use AES-CBC variants (segmented or stream-CBC).
For Agile details, see [`docs/22-ooxml-encryption.md`](./22-ooxml-encryption.md).

## `EncryptedPackage` stream layout

`EncryptedPackage` is an OLE stream with:

1. `orig_size: u64le` (`StreamSize` in the spec): **plaintext** (unencrypted) package size in bytes.
2. `ciphertext: [u8]` (`EncryptedData` in the spec): encrypted bytes of the underlying OPC package
   (the `.xlsx` ZIP bytes).

Spec note (MS-OFFCRYPTO §2.3.4.4): the *physical* stream length can be **larger** than `orig_size`
because the encrypted data is padded to a cipher block boundary.

## AES decryption (baseline + variants)

Baseline (MS-OFFCRYPTO / ECMA-376 Standard AES): decrypt the ciphertext bytes with **AES-ECB(key)**:

* AES key: derived from the password and `EncryptionInfo` (out of scope for this note).
* Mode: ECB.
* Padding: none (ciphertext is block-aligned).
* Ciphertext length (bytes after the `u64` prefix) must be a **multiple of 16** bytes.

Variants (observed in non-Excel tooling / compatibility cases):

- AES-CBC in 4096-byte segments with a per-segment IV (one common IV derivation is
  `IV = SHA1(salt || LE32(segmentIndex))[0..16]`).
- AES-CBC in 4096-byte segments with per-segment keys (block-indexed keys) and `IV=0`.
- AES-CBC as a single stream using `IV = salt`.

In this repo:

- `crates/formula-offcrypto`’s Standard decryptor uses AES-ECB for `EncryptedPackage`.
- `crates/formula-office-crypto` tries a small set of Standard AES-CBC layouts to maximize
  compatibility.
- `crates/formula-io/src/offcrypto/encrypted_package.rs` implements a CBC-segmented helper (IV derived
  from salt + segment index).

## Padding + truncation (do not trust PKCS#7)

Do **not** treat decrypted bytes as PKCS#7 and “unpad”.

Instead:

1. Decrypt all ciphertext blocks.
2. **Truncate the plaintext to `orig_size`**.

Rationale:

* MS-OFFCRYPTO §2.3.4.4 defines `StreamSize`/`orig_size` as authoritative for the unencrypted size,
  and explicitly allows the encrypted stream to be larger due to block-size padding.
* MS-OFFCRYPTO §2.3.4.15 notes the final data block is padded to the cipher block size (and padding
  bytes can be arbitrary).

## Validation rules (recommended)

These checks catch most corruption / truncation issues early:

* Stream length must be `>= 8` (need `orig_size` prefix).
* `orig_size` must fit in your address space / target type (e.g., `usize`).
* Let `ct = stream[8..]`. Require `ct.len() % 16 == 0` (AES block alignment).
* After decryption, if produced plaintext is `< orig_size`, treat as truncated/corrupt; otherwise
  truncate to `orig_size` and continue.
