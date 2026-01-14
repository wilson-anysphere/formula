# MS-OFFCRYPTO Standard/CryptoAPI AES: `EncryptedPackage` decryption notes

This repo detects password-protected / encrypted OOXML workbooks as an **OLE/CFB** container with
`EncryptionInfo` + `EncryptedPackage` streams (MS-OFFCRYPTO).

High-level behavior in `formula-io`:

- Encryption is always detected and surfaced via dedicated errors (`PasswordRequired` /
  `InvalidPassword` / `UnsupportedOoxmlEncryption`) so callers can prompt for a password and route
  “unsupported encryption” reports correctly.
- With the `formula-io` cargo feature **`encrypted-workbooks`** enabled:
  - The legacy password-aware entrypoints `open_workbook_with_password` /
    `open_workbook_model_with_password` can decrypt **Agile (4.4)** encrypted `.xlsx`/`.xlsm`/`.xlsb`,
    but still treat **Standard/CryptoAPI** as `PasswordRequired`/`InvalidPassword` for UX
    compatibility.
  - `open_workbook_with_options` uses a **streaming decryptor** and can open some **Standard (3.2)**
    and **Agile (4.4)** encrypted `.xlsx`/`.xlsm` as `Workbook::Model` (note: this streaming path does
    not yet validate Agile `dataIntegrity` HMAC).

Standard/CryptoAPI decryption primitives also exist in lower-level crates (notably
`crates/formula-offcrypto` and `crates/formula-office-crypto`), but the full open-path plumbing is
still converging (see [`docs/21-encrypted-workbooks.md`](./21-encrypted-workbooks.md)).

This document focuses on the `EncryptedPackage` stream itself, because the most common interop bugs
cluster around:

- the 8-byte plaintext length prefix,
- **cipher mode mismatches** (ECB vs CBC-style variants),
- and **padding/truncation** (do **not** PKCS#7-unpad; always truncate to the declared size).

For Agile (4.4) OOXML decryption details (and `dataIntegrity` gotchas), see
[`docs/22-ooxml-encryption.md`](./22-ooxml-encryption.md).

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

1. `orig_size: u64le` (`StreamSize` in the spec): **plaintext** (unencrypted) package size in bytes.
2. `ciphertext: [u8]` (`EncryptedData` in the spec): encrypted bytes of the underlying OPC package
   (the `.xlsx` ZIP bytes).

Spec note (MS-OFFCRYPTO §2.3.4.4): the *physical* stream length can be **larger** than `orig_size`
because the encrypted data is padded to a cipher block boundary.

## Decrypting `EncryptedPackage`: observed Standard/CryptoAPI AES variants

The Standard/CryptoAPI header does not reliably communicate “which AES mode” was used for
`EncryptedPackage`. In this repo we have encountered at least:

- **AES-ECB** (no IV), implemented by `crates/formula-offcrypto` (see also
  `fixtures/encrypted/ooxml/standard.xlsx`).
- **Segmented AES-CBC variants**, implemented by the streaming decryptor in
  `crates/formula-io/src/encrypted_package_reader.rs` (and by some lower-level helpers).

If you are implementing Standard decryption for real-world compatibility, expect to need **mode
fallback**.

### Variant A: AES-ECB (no IV)

Decrypt the ciphertext bytes (everything after the `u64` prefix) with **AES-ECB(key)**:

* AES key: derived from the password and `EncryptionInfo` (out of scope for this note).
* Mode: ECB.
* Padding: none at the crypto layer (ciphertext is block-aligned).
* Ciphertext length must be a **multiple of 16** bytes.

Implementation pointer: `crates/formula-offcrypto/src/lib.rs` (`decrypt_standard_only`).

### Variant B: segmented AES-CBC (`0x1000` chunks; per-segment IV)

Some Standard/CryptoAPI AES producers reuse an Agile-like “encrypt in 4096-byte segments” layout:

- Plaintext is processed in **0x1000 (4096) byte segments**; segment index `i` is **0-based**.
- Each segment is encrypted independently with AES-CBC (no chaining across segments).
- Each segment’s ciphertext is padded to a 16-byte boundary.
- The final ciphertext “segment” is often “whatever bytes remain” (can include an extra full padding
  block and/or trailing OLE stream slack).

The Standard/CryptoAPI AES-CBC implementation used by `crates/formula-io` derives IVs as:

```text
iv_full = SHA1(salt || LE32(i))
iv = iv_full[0..16]  // truncate to 16 bytes for AES
```

Where:

* `salt` is the 16-byte salt (`EncryptionVerifier.Salt`).
* `LE32(i)` is the segment index encoded as a little-endian `u32`.

Other CBC-style Standard variants exist (e.g. per-segment keys with IV=0). In this repo,
`crates/formula-office-crypto/src/standard.rs` tries a small set of these schemes (`StandardScheme`)
for compatibility.

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
