# MS-OFFCRYPTO Standard/CryptoAPI AES: `EncryptedPackage` decryption notes

This repo detects password-protected / encrypted OOXML workbooks as an OLE/CFB container with
`EncryptionInfo` + `EncryptedPackage` streams.

In `formula-io`, attempting to open these files without a password will surface `Error::PasswordRequired`.
The password-aware helpers (`open_workbook_with_password` / `open_workbook_model_with_password`) can
surface `Error::InvalidPassword`. (End-to-end decryption is still being wired; see
[`docs/21-encrypted-workbooks.md`](./21-encrypted-workbooks.md) for current behavior and entrypoints.)

For unknown `EncryptionInfo` versions, `formula-io` may surface
`Error::UnsupportedOoxmlEncryption { version_major, version_minor }`.

If/when we add **Standard Encryption (CryptoAPI AES)** decryption support, the most common
interoperability bugs are in the `EncryptedPackage` stream framing (the `u64` size prefix), **0x1000
segmenting**, **IV derivation**, and **padding/truncation**. This note is meant as a compact
developer reference.

## Normative spec references (MS-OFFCRYPTO)

* **`\\EncryptedPackage` stream layout**: MS-OFFCRYPTO **§2.3.4.4** “`\\EncryptedPackage` Stream”.
  * Defines `StreamSize` and notes the stream can be larger than `StreamSize` due to block padding.
* **Standard Encryption metadata (`\\EncryptionInfo`)**: MS-OFFCRYPTO **§2.3.4.5** “`\\EncryptionInfo`
  Stream (Standard Encryption)”.
* **Salt location/size**: MS-OFFCRYPTO **§2.3.4.7** “ECMA-376 Document Encryption Key Generation
  (Standard Encryption)”.
  * Salt is 16 bytes and stored in `EncryptionVerifier.Salt`.
* **Segmenting + per-segment IV**: MS-OFFCRYPTO **§2.3.4.12** “Initialization Vector Generation
  (Agile Encryption)” and **§2.3.4.15** “Data Encryption (Agile Encryption)”.
  * These sections normatively describe the 4096-byte segmenting and “segment number → IV” pattern.
  * In practice, Excel’s **CryptoAPI AES `EncryptedPackage`** decryption matches the same segment/IV
    scheme described there.

## `EncryptedPackage` stream layout

`EncryptedPackage` is an OLE stream with:

1. `orig_size: u64le` (`StreamSize` in the spec): **plaintext** (unencrypted) package size in bytes.
2. `ciphertext: [u8]` (`EncryptedData` in the spec): encrypted bytes of the underlying OPC package
   (the `.xlsx` ZIP bytes).

Spec note (MS-OFFCRYPTO §2.3.4.4): the *physical* stream length can be **larger** than `orig_size`
because the encrypted data is padded to a cipher block boundary.

## Segmenting (0x1000)

Plaintext is processed in **0x1000 (4096) byte segments**; segment index `i` is **0-based**.

For decryption this means:

* There is **no CBC chaining across segments**.
* Segments `0..n-2` correspond to 4096 bytes of plaintext each.
* The final segment corresponds to the remaining plaintext bytes (`orig_size % 4096`), plus whatever
  padding bytes the producer wrote.

MS-OFFCRYPTO reference: §2.3.4.15 (segmenting requirement; described for agile encryption).

## IV derivation (per segment)

For segment index `i` (0-based), derive the AES-CBC IV as:

```text
iv_full = SHA1(salt || LE32(i))
iv = iv_full[0..16]  // truncate to 16 bytes for AES
```

Where:

* `salt` is the 16-byte salt (`EncryptionVerifier.Salt`, MS-OFFCRYPTO §2.3.4.7).
* `LE32(i)` is the segment index encoded as a little-endian `u32`.

MS-OFFCRYPTO reference: §2.3.4.12 + §2.3.4.15 (IV generation using a blockKey derived from segment
number; hashing + truncation).

## AES decryption (per segment)

Decrypt each segment independently with **AES-CBC(key, iv(i))**:

* AES key: derived from the password and `EncryptionInfo` (out of scope for this note).
* Mode: CBC.
* Segment IV: derived as above.
* Ciphertext segment lengths must be a **multiple of 16** bytes.

## Padding + truncation (do not trust PKCS#7)

Do **not** treat decrypted bytes as PKCS#7 and “unpad”.

Instead:

1. Decrypt all segments.
2. Concatenate plaintext segments.
3. **Truncate the concatenated plaintext to `orig_size`**.

Rationale:

* MS-OFFCRYPTO §2.3.4.4 defines `StreamSize`/`orig_size` as authoritative for the unencrypted size,
  and explicitly allows the encrypted stream to be larger due to block-size padding.
* MS-OFFCRYPTO §2.3.4.15 notes the final data block is padded to the cipher block size (and padding
  bytes can be arbitrary).

### Edge case: “extra full padding block” from Excel

When `orig_size` is an exact multiple of both:

* 16 (AES block size), **and**
* 4096 (segment size),

Excel can emit an additional *full* 16-byte padding block at the end. In that case, the final
ciphertext “segment” is **0x1010 (4112) bytes**, not 0x1000.

Decryption should therefore treat the final segment as “the remainder of the stream”, not as exactly
4096 bytes.

## Validation rules (recommended)

These checks catch most corruption / truncation issues early:

* Stream length must be `>= 8` (need `orig_size` prefix).
* `orig_size` must fit in your address space / target type (e.g., `usize`).
* Let `ct = stream[8..]`. Require `ct.len() % 16 == 0` (AES block alignment).
* Let `n = ceil(orig_size / 4096)` (0 if `orig_size == 0`).
  * If `n > 0`, require `ct.len() >= (n - 1) * 4096` (need all non-final segments).
  * For segments `0..n-2`, require exactly 4096 bytes of ciphertext each.
  * For the final segment, require `(ct.len() - (n - 1) * 4096) % 16 == 0` and allow it to be `> 4096`
    (Excel padding edge case).
* After decryption, if produced plaintext is `< orig_size`, treat as truncated/corrupt; otherwise
  truncate to `orig_size` and continue.
