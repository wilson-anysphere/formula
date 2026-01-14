# MS-OFFCRYPTO Standard (CryptoAPI) encryption: key derivation + verifier validation

This document is a *from-scratch* implementation guide for decrypting **MS Office “Standard” encryption**
(`EncryptionInfo` `versionMinor == 2`, commonly **3.2**) used by password-protected OOXML files (e.g.
`.xlsx`, `.docx`, `.pptx`)
stored inside an OLE Compound File.

It focuses on:

1. Detecting **Standard (CryptoAPI)** encryption.
2. Parsing the **binary** `EncryptionInfo` stream layout.
3. Deriving keys using the fixed **spinCount = 50,000** password hashing loop.
4. Implementing **CryptoAPI `CryptDeriveKey`** (ipad/opad expansion with `0x36` / `0x5c`).
5. Validating the password by decrypting and checking the **verifier**.

The intent is that an engineer can implement this without reading any external references.

For repo-specific implementation notes (which parameter subsets we accept, crate entrypoints, writer
defaults), see `docs/office-encryption.md`.

Important nuance: this guide describes the **Excel/`msoffcrypto-tool`-style** Standard/CryptoAPI AES
variant implemented by `crates/formula-offcrypto`, where the verifier fields and `EncryptedPackage`
are decrypted with **AES-ECB** (no IV) after key derivation.

If you need broader real-world compatibility across non-Excel producers, prefer
`crates/formula-office-crypto`’s Standard decryptor.

For Agile (4.4) OOXML password decryption details (different scheme; XML descriptor + `dataIntegrity`
HMAC), see `docs/22-ooxml-encryption.md`.

---

## Implementation references in this repo

If you’re changing or debugging the code, start here:

* Standard/CryptoAPI parsing + key derivation + verifier check:
  * `crates/formula-offcrypto/src/lib.rs`:
    * `parse_encryption_info` (Standard branch)
    * `standard_derive_key` / `standard_verify_key` (Excel-default AES/SHA-1)
  * `crates/formula-io/src/offcrypto/standard.rs`:
    * `parse_encryption_info_standard`
    * `verify_password_standard`
* `EncryptedPackage` decryption helpers:
  * `crates/formula-offcrypto/src/lib.rs`: `decrypt_encrypted_package_ecb`
  * `crates/formula-io/src/offcrypto/encrypted_package.rs`:
    * `decrypt_encrypted_package_standard_aes_to_writer` (streaming AES-ECB)
    * `decrypt_standard_encrypted_package_stream` (buffered; includes a non-standard segmented fallback)
* More permissive Standard decryptor (supports more variants than `formula-offcrypto`):
  * `crates/formula-office-crypto/src/standard.rs`
* Related docs:
  * `docs/offcrypto-standard-encryptedpackage.md`
  * `docs/offcrypto-standard-cryptoapi-rc4.md`

---

## 1) Detecting “Standard” encryption (`versionMinor == 2`, commonly 3.2)

An encrypted OOXML file is an **OLE Compound File** (CFB) containing (at minimum) these streams:

* `EncryptionInfo` – metadata and verifier used to derive keys and validate the password.
* `EncryptedPackage` – the encrypted bytes of the real OOXML ZIP package.

To detect **Standard (CryptoAPI)** encryption:

1. Open the file as an OLE compound file.
2. Read the `EncryptionInfo` stream.
3. Parse its first 8 bytes as:
   * `major: u16le`
   * `minor: u16le`
   * `flags: u32le`
4. Standard encryption is identified by:

```text
minor = 2
major ∈ {2, 3, 4}   // Excel commonly uses 3.2; other producers may emit 4.2, etc.
```

Some producers also emit other `*.2` major versions (e.g. `4.2`) while still using the same
CryptoAPI/Standard container layout. A tolerant reader can treat `versionMinor == 2` as
“Standard/CryptoAPI” and then validate the rest of the structure defensively.

The `flags` field is part of the `EncryptionInfo` header (and must be consumed to keep offsets correct),
but it is not needed for the scheme dispatch.

If `versionMinor != 2`, the file is *not* Standard encryption (it may be Agile `4.4`, Extensible, etc.).

---

## 2) `EncryptionInfo` binary layout (`versionMinor == 2`)

All integer fields in this section are **little-endian**.

At a high level:

```text
EncryptionInfoStream =
  EncryptionVersionInfo (8 bytes)  // major + minor + flags
  HeaderSize            (4 bytes)  // u32le
  EncryptionHeader              // HeaderSize bytes
  EncryptionVerifier            // remainder of stream
```

### 2.1) `EncryptionVersionInfo` (8 bytes)

| Offset | Size | Type   | Name  | Meaning |
|-------:|-----:|--------|-------|---------|
| 0x00   | 2    | u16le  | Major | Commonly 3; `2` and `4` also occur with `Minor == 2` |
| 0x02   | 2    | u16le  | Minor | 2 for Standard |
| 0x04   | 4    | u32le  | Flags | Header flags (not required for decryption; consume for correct parsing) |

### 2.2) `HeaderSize` (4 bytes)

| Offset | Size | Type  | Name       | Meaning |
|-------:|-----:|-------|------------|---------|
| 0x08   | 4    | u32le | HeaderSize | Byte length of `EncryptionHeader` that follows |

`HeaderSize` **does not include** the 12 bytes of `EncryptionVersionInfo + HeaderSize` itself.

### 2.3) `EncryptionHeader` (variable length, `HeaderSize` bytes)

The header starts with eight DWORDs (**32 bytes**) and then contains a variable-length CSP name string.

Important: the `SizeExtra` field means the CSP name does **not** necessarily consume all remaining
`EncryptionHeader` bytes. Real-world producers exist where `SizeExtra != 0` and **opaque trailing
bytes** follow the CSP name inside the header blob. Tolerant parsers must not interpret these extra
bytes as UTF‑16LE text.

Byte layout:

```text
EncryptionHeaderBytes (HeaderSize bytes) =
  FixedFields[32] ||
  CSPNameBytes[HeaderSize - 32 - SizeExtra] ||   // UTF-16LE (often NUL-terminated)
  Extra[SizeExtra]                               // opaque / unused
```

```text
struct EncryptionHeader {
  u32 Flags;
  u32 SizeExtra;      // number of trailing bytes after CSPName (often 0, but not guaranteed)
  u32 AlgID;          // cipher ALG_ID (RC4 / AES_*)
  u32 AlgIDHash;      // hash ALG_ID (SHA1 / MD5)
  u32 KeySize;        // in *bits*
  u32 ProviderType;   // typically PROV_RSA_FULL (1) or PROV_RSA_AES (24)
  u32 Reserved1;      // should be 0
  u32 Reserved2;      // should be 0
  u8  CSPNameBytes[]; // UTF-16LE bytes (length = HeaderSize - 32 - SizeExtra; often includes a NUL terminator)
  u8  Extra[];        // trailing SizeExtra bytes (opaque; ignore)
}
```

Notes:

* `CSPName` is only metadata (e.g. `Microsoft Enhanced RSA and AES Cryptographic Provider`) and is not
  required to derive keys.
* Correct parsing rule (using the enclosing `HeaderSize`):
  * `csp_name_bytes_len = headerSize - 32 - sizeExtra` (validate `headerSize >= 32` and
    `sizeExtra <= headerSize - 32`)
  * Decode the CSP name from `header[32 : 32 + csp_name_bytes_len]` **only**.
  * Ignore the remaining `sizeExtra` bytes; do not treat them as UTF‑16.
* Do **not** scan for the NUL terminator to determine header length. Use `HeaderSize` (and
  `SizeExtra`) to compute the CSP-name slice, then (optionally) truncate at the first NUL *within that
  slice* when decoding.
* Interoperability motivation: some producers set `SizeExtra` non-zero and include non-text bytes
  after the CSP name. Treating “the rest of the header” as UTF‑16 can lead to decode failures or
  spurious provider names.

### 2.4) `EncryptionVerifier` (remainder of stream)

Immediately after `EncryptionHeader` ends, the `EncryptionVerifier` begins.

```text
struct EncryptionVerifier {
  u32 SaltSize;              // typically 16
  u8  Salt[SaltSize];
  u8  EncryptedVerifier[16]; // always 16 bytes
  u32 VerifierHashSize;      // 16 (MD5) or 20 (SHA1)
  u8  EncryptedVerifierHash[...]; // rest of stream
}
```

Important parsing detail:

* There is **no explicit length field** for `EncryptedVerifierHash`.
  Treat it as: “whatever bytes remain in the `EncryptionInfo` stream after the fields above”.
* For AES, the *plaintext* verifier hash is `VerifierHashSize` bytes, but the *ciphertext* is commonly
  **padded to a 16-byte AES block boundary** (e.g. SHA‑1 is 20 bytes plaintext but 32 bytes of ciphertext).
  Keep the whole ciphertext when decrypting, then compare only the first `VerifierHashSize` bytes.

---

## 3) Algorithm identifiers you must support (ALG_ID)

The `EncryptionHeader` `AlgID` and `AlgIDHash` are CryptoAPI `ALG_ID` values.

### 3.1) Cipher `AlgID` values

| Cipher | Name | ALG_ID (hex) | Key size bytes |
|--------|------|--------------|----------------|
| RC4 | `CALG_RC4` | `0x00006801` | `KeySize / 8` |
| AES-128 | `CALG_AES_128` | `0x0000660E` | 16 |
| AES-192 | `CALG_AES_192` | `0x0000660F` | 24 |
| AES-256 | `CALG_AES_256` | `0x00006610` | 32 |

**RC4 key size semantics:** `KeySize` is stored in *bits*. MS-OFFCRYPTO specifies that for **RC4**,
`KeySize == 0` MUST be interpreted as **40** (legacy 40-bit RC4).

For MS-OFFCRYPTO Standard RC4, the RC4 KSA key length is exactly `keyLen = KeySize / 8` bytes
(40→5, 56→7, 128→16). When deriving per-block keys, use the first `keyLen` bytes of the per-block
hash (`H_block[0..keyLen]`).

Note: some legacy CryptoAPI RC4 implementations “expand” 40-bit keys to 16 bytes by appending
11 zero bytes. Do **not** apply that padding for MS-OFFCRYPTO Standard RC4.

### 3.2) Hash `AlgIDHash` values

| Hash | Name | ALG_ID (hex) | Digest bytes |
|------|------|--------------|--------------|
| MD5 | `CALG_MD5` | `0x00008003` | 16 |
| SHA-1 | `CALG_SHA1` | `0x00008004` | 20 |
| SHA-256 | `CALG_SHA_256` | `0x0000800C` | 32 |
| SHA-384 | `CALG_SHA_384` | `0x0000800D` | 48 |
| SHA-512 | `CALG_SHA_512` | `0x0000800E` | 64 |

---

## 4) Password hashing (fixed spinCount = 50,000)

Standard encryption uses a custom iterative hash loop with a **fixed** iteration count:

```text
spinCount = 50000
```

### 4.1) Password bytes

Convert the password to **UTF‑16LE** bytes:

* No BOM
* No NUL terminator
* No Unicode normalization and no whitespace trimming: the KDF operates on the exact UTF-16LE byte
  sequence. NFC vs NFD (and trailing spaces) produce different derived keys.

Example:

```text
password = "password"
passwordUtf16le = 70 00 61 00 73 00 73 00 77 00 6f 00 72 00 64 00
```

### 4.2) Iterative hash loop

Let `Hash()` be MD5 or SHA1 depending on `AlgIDHash`.

Inputs:

* `salt` = `EncryptionVerifier.Salt` (`SaltSize` bytes, commonly 16)
* `passwordBytes` = UTF‑16LE bytes as above

Algorithm:

```text
H = Hash( salt || passwordBytes )

for i in 0 .. spinCount-1:
  H = Hash( LE32(i) || H )

H_final = H
```

Where:

* `LE32(i)` is the 4-byte little-endian encoding of the unsigned 32-bit integer `i`.
* The loop runs **exactly 50,000 times** (i = 0 through 49,999).

This `H_final` is the master password hash used to derive keys for:

* The verifier (password check)
* The package decryption keys

---

## 5) Key derivation overview

Standard encryption derives keys from the password hash and a per-block 32-bit index.

Terminology used below:

* `block` – a 32-bit unsigned “block key” (`u32`), encoded as `LE32(block)`.
  * `block = 0` is used to derive the **file key** (and to validate the password using the verifier).
  * For `EncryptedPackage`:
    * **AES** decrypts the full ciphertext with the derived file key (`block = 0`) using **AES-ECB**
      (no IV) (see §7.2.1).
    * **RC4** varies the **key** by `segmentIndex` (see §7.2.2).

### 5.1) Per-block hash input (key material)

Compute:

```text
H_block = Hash( H_final || LE32(block) )
```

This `H_block` is the per-block hash input from which the actual symmetric key bytes are derived.

* For **AES**, `H_block` is fed into the CryptoAPI `CryptDeriveKey` expansion (see §5.2).
* For **RC4**, the RC4 key is typically derived by truncating `H_block` (see §5.2).

### 5.2) Deriving the symmetric key bytes (AES vs RC4)

#### 5.2.1) RC4 (`CALG_RC4`): key material = truncate(`H_block`) (+ 40-bit padding quirk)

For Standard **RC4** encryption, the per-block RC4 key is derived by truncating the hash:

```text
keySizeBits = KeySize
if keySizeBits == 0:
  keySizeBits = 40                           // MS-OFFCRYPTO: RC4 KeySize=0 means 40-bit
keyLen = keySizeBits / 8

key_material(block) = H_block[0 : keyLen]

if keySizeBits == 40:
  rc4_key(block) = key_material(block) || 0x00 * 11   // 16 bytes total
else:
  rc4_key(block) = key_material(block)                // 7 bytes (56-bit) or 16 bytes (128-bit)
```

This matches the common Standard RC4 “re-key per 0x200-byte block” scheme (see §7.2.2).

Important: the **40-bit** padding behavior (`KeySize == 0` or `KeySize == 40`) is a
**CryptoAPI/Office interoperability quirk**.
RC4’s KSA depends on *both the key bytes and the key length*, so `key_material` (5 bytes) and
`key_material || 0x00*11` (16 bytes) yield different keystreams.

#### 5.2.2) AES (`CALG_AES_*`): CryptoAPI `CryptDeriveKey` (ipad/opad expansion)

CryptoAPI’s `CryptDeriveKey` is **not PBKDF2**. For MD5/SHA1 it behaves like:

* It expands the hash using an ipad/opad construction and then truncates to the required key length.

Constants:

```text
ipad byte = 0x36
opad byte = 0x5c
```

Pseudocode (MD5/SHA1 only; both use 64-byte blocks):

```text
// Inputs:
//   H_block: digest bytes (16 for MD5, 20 for SHA1)
//   keyLen:  desired key length in bytes (KeySize / 8)
// Output:
//   keyBytes[keyLen]

function CryptDeriveKey(Hash, H_block, keyLen):
  digestLen = len(H_block)         // 16 or 20

  blockLen = 64                    // MD5/SHA1 block size in bytes

  // 1) Pad digest to 64 bytes with zeros
  D = H_block || 0x00 * (blockLen - digestLen)

  // 2) Build 64-byte ipad/opad buffers
  I = 0x36 repeated blockLen times
  O = 0x5c repeated blockLen times

  // 3) HMAC-like expansion
  inner = Hash( D XOR I )
  outer = Hash( D XOR O )

  // 4) Concatenate and truncate
  derived = inner || outer         // 32 bytes for MD5, 40 bytes for SHA1
  if keyLen > len(derived):
    error("requested key length too long")
  return derived[0:keyLen]
```

This is sufficient for Standard encryption because Office only requests up to 32 bytes of key material
(AES-256), and `inner || outer` yields 40 bytes for SHA‑1.

### 5.3) IV derivation: none for Standard AES-ECB

RC4 is a stream cipher and has no IV.

For the most common **Standard AES** files (including the repo’s `msoffcrypto-tool`-generated
fixtures), both:

* the password verifier fields, and
* the `EncryptedPackage` ciphertext

are decrypted with **AES-ECB**, which uses **no IV**.

---

## 6) Password verifier validation (critical correctness details)

Once you can derive `key` for `block = 0`, you can check whether the password is correct.

Inputs from `EncryptionVerifier`:

* `Salt`
* `EncryptedVerifier` (16 bytes)
* `VerifierHashSize` (16 or 20)
* `EncryptedVerifierHash` (remaining bytes; AES ciphertext may include padding)

### 6.1) Derive block-0 key

```text
H_final  = hash_password(password, Salt, spinCount=50000)         // §4
H_block0 = Hash( H_final || LE32(0) )                             // §5.1
keySizeBits = KeySize
if AlgID == CALG_RC4 and keySizeBits == 0:
  keySizeBits = 40                                                // MS-OFFCRYPTO: RC4 KeySize=0 means 40-bit
keyLen   = keySizeBits / 8

if AlgID == CALG_RC4:
  key_material = H_block0[0:keyLen]
  if keySizeBits == 40:
    key = key_material || 0x00 * 11                               // §5.2.1
  else:
    key = key_material                                            // §5.2.1
else:
  key = CryptDeriveKey(Hash, H_block0, keyLen=keyLen)             // §5.2.2
```

### 6.2) Decrypt verifier fields (RC4 must be streamed; AES can be independent)

For **Standard RC4**, RC4 is a stream cipher and keystream position matters: the verifier and
verifier-hash bytes must be decrypted as **one continuous stream**. The simplest approach is to
concatenate the two ciphertext buffers and decrypt once.

For **Standard AES**, the verifier fields are **AES-ECB** (no IV), so there is no chaining and you
may decrypt the two ciphertext blobs independently. (The concatenate+decrypt approach below also
works for AES because ECB has no state.)

Compatibility note (non-Excel producers): some implementations encrypt the verifier fields using
**AES-CBC (no padding)** instead of ECB. In that case:

* You must treat `EncryptedVerifier || EncryptedVerifierHash` as a single CBC stream (CBC chaining
  crosses the boundary between the two fields).
* The IV is commonly derived from the verifier salt and block index 0:

  ```text
  iv0 = Hash(Salt || LE32(0))[0..16]   // Hash is AlgIDHash (typically SHA-1 for Standard AES)
  ```

  `formula-io`’s Standard verifier implementation tries **AES-ECB first**, then falls back to this
  derived-IV CBC mode for compatibility.

Steps:

1. Concatenate ciphertext:

   ```text
   C = EncryptedVerifier || EncryptedVerifierHash
   ```

2. Decrypt `C` with the cipher indicated by `EncryptionHeader.AlgID` using the derived `key`:

   * **AES (`CALG_AES_*`)**: AES-ECB over 16-byte blocks (no IV, no padding at the crypto layer).
     (If you implement the AES-CBC compatibility variant above, decrypt with AES-CBC using `iv0`.)
   * **RC4 (`CALG_RC4`)**: RC4 stream cipher with the derived key.

   This produces plaintext `P`.

3. Split plaintext:

   ```text
   Verifier      = P[0:16]
   VerifierHash  = P[16 : 16 + VerifierHashSize]
   // Ignore any remaining bytes after 16+VerifierHashSize (AES padding / unused bytes)
   ```

### 6.3) Hash the verifier and compare

Compute:

```text
ExpectedHash = Hash(Verifier)
```

The password is correct if:

```text
ExpectedHash[0:VerifierHashSize] == VerifierHash
```

Notes:

* For SHA1, `VerifierHashSize` is typically 20.
* For MD5, `VerifierHashSize` is 16.
* Always compare using a constant-time comparison if the code path is security-sensitive.

---

## 7) (Optional but recommended) Decrypting `EncryptedPackage`

While this document’s success criteria is key derivation + verifier validation, engineers usually need the
next step: decrypting the actual OOXML ZIP package.

### 7.1) Stream layout

The `EncryptedPackage` stream begins with:

```text
u32le OriginalPackageSizeLo
u32le OriginalPackageSizeHiOrReserved
u8    EncryptedBytes[...]
```

Compatibility note: while MS-OFFCRYPTO describes the size prefix as a `u64le`, some
producers/libraries treat it as `u32 totalSize` + `u32 reserved` (often 0). To avoid truncation or
“huge size” misreads, parse it as two little-endian DWORDs and recombine:

```text
lo = u32le(bytes[0..4])
hi = u32le(bytes[4..8])
OriginalPackageSize = lo as u64 | ((hi as u64) << 32)
```

After decryption, truncate the plaintext to exactly `OriginalPackageSize` bytes (the ciphertext is padded).

For additional `EncryptedPackage` framing/padding guidance and known non-Excel variants, see
`docs/offcrypto-standard-encryptedpackage.md`.

### 7.2) Encryption model (AES vs RC4)

#### 7.2.1) AES (`CALG_AES_*`): AES-ECB (no IV; optional 0x1000 segmenting)

For baseline Standard/CryptoAPI AES encryption, the `EncryptedPackage` ciphertext (after the 8-byte
size prefix) is **AES-ECB** encrypted with the derived `fileKey` (`block = 0`).

Algorithm:

1. Read `OriginalPackageSize` (8-byte plaintext size prefix; see note above).
2. Let `C = EncryptedBytes` (all remaining bytes).
3. Require `len(C) % 16 == 0`.
4. `P = AES-ECB-Decrypt(fileKey, C)`.
5. Return `P[0:OriginalPackageSize]` (truncate to the declared size).

Notes:

* The `EncryptedPackage` stream can be larger than `OriginalPackageSize` due to block padding and/or
  OLE sector slack. **Always truncate** to the declared size after decrypting.
* Some producers pad the ciphertext to a fixed size (e.g. to 4096 bytes for very small packages).
  Truncation handles this.
* AES-ECB has no IV. Excel-default Standard AES therefore has **no per-segment IV** (unlike Agile).
  Some non-Excel producers use a segmented AES-CBC variant with derived per-segment IVs (see below).

Even in the ECB baseline, many implementations decrypt `EncryptedPackage` in **0x1000-byte (4096)**
segments for streaming/bounded memory:

```text
segmentSize = 0x1000   // 4096
```

With ECB this is an implementation detail: you can decrypt the ciphertext in any chunking as long as
it stays AES-block-aligned (multiples of 16 bytes).

Compatibility note (non-Excel producers): some implementations encrypt `EncryptedPackage` in 0x1000
segments using **AES-CBC** with a per-segment IV derived from the verifier salt:

```text
iv_i = SHA1(Salt || LE32(i))[0..16]
plaintext_i = AES-CBC-Decrypt(fileKey, iv_i, ciphertext_i)
```

This is **not** the Excel-default Standard AES scheme, but `formula-io`’s `EncryptedPackage` helper
can try it as a fallback when a salt is available (see `docs/offcrypto-standard-encryptedpackage.md`).

#### 7.2.2) RC4 (`CALG_RC4`)

RC4-based Standard encryption uses **512-byte** blocks and resets the RC4 keystream per block:

```text
segmentSize = 0x200   // 512
```

For segment index `i = 0, 1, 2, ...`:

1. Derive `H_block = Hash(H_final || LE32(i))`.
2. Let `keySizeBits = KeySize` (if `keySizeBits == 0`, set `keySizeBits = 40` for RC4).
3. `key_material_i = H_block[0 : keySizeBits/8]` (truncate to the configured key size).
4. If `keySizeBits == 40`, set `key_i = key_material_i || 0x00*11` (16 bytes).
   Otherwise set `key_i = key_material_i`.
5. Initialize RC4 with `key_i` (fresh state for each segment) and decrypt exactly one segment of
   ciphertext.

Concatenate segments and truncate to `OriginalPackageSize`.

For a deeper Standard RC4 writeup (including test vectors and “0x200 vs 0x400 block size” gotchas),
see `docs/offcrypto-standard-cryptoapi-rc4.md`.

---

## 8) Worked example (test vector)

This example is intentionally small and deterministic. It is **not** a full Office file; it is just the key
derivation math that you can use as a unit test.

Parameters:

* Hash algorithm: SHA-1 (`CALG_SHA1`, `0x00008004`)
* Cipher: AES-256 (`CALG_AES_256`, `0x00006610`)
* KeySize: 256 bits → 32 bytes
* spinCount: 50,000
* `block = 0` (used to derive the file key and validate the verifier; RC4 uses other block indices
  for per-block package keys)

Inputs:

```text
password = "password"
passwordUtf16le =
  70 00 61 00 73 00 73 00 77 00 6f 00 72 00 64 00

salt =
  00 01 02 03 04 05 06 07 08 09 0a 0b 0c 0d 0e 0f
```

Derived values:

```text
H_final  = 1b5972284eab6481eb6565a0985b334b3e65e041
H_block0 = 6ad7dedf2da3514b1d85eabee069d47dd058967f

// If using CryptoAPI RC4 with KeySize=0/40 (40-bit), the per-block RC4 key for block=0 would be:
rc4_key_block0_40bit = 6ad7dedf2d0000000000000000000000

key (32 bytes, CryptDeriveKey expansion) =
  de5451b9dc3fcb383792cbeec80b6bc3
  0795c2705e075039407199f7d299b6e4

// AES-192 uses the same derivation output, truncated to 24 bytes:
key (24 bytes, AES-192) =
  de5451b9dc3fcb383792cbeec80b6bc3
  0795c2705e075039

`EncryptedPackage` for baseline Standard AES is decrypted with **AES-ECB** and uses **no IV**.

### 8.1) AES-128 + AES-ECB `EncryptedPackage` sanity check (real file bytes)

This vector matches a real Standard-encrypted workbook in this repo:
`fixtures/encrypted/ooxml/standard.xlsx` (password: `password`).
See also `fixtures/encrypted/ooxml/standard-unicode.xlsx` (password: `pässwörd🔒`, NFC, includes
non-BMP emoji) for Unicode password regression coverage.

Parameters:

* Hash algorithm: SHA‑1
* Cipher: AES‑128
* KeySize: 128 bits → 16 bytes
* spinCount: 50,000
* `block = 0`

Inputs:

```text
password = "password"
passwordUtf16le =
  70 00 61 00 73 00 73 00 77 00 6f 00 72 00 64 00

salt =
  00 11 22 33 44 55 66 77 88 99 aa bb cc dd ee ff
```

Derived values:

```text
H_final  = 5ac2afc2a117ec25f449be1993cbe67c068458d8
H_block0 = 11958e53fc62fdebc0107e2fae8650147de123e8

key (AES-128, 16 bytes) =
  5e8727d6c94408a903aececf1382b380
```

`EncryptedPackage` bytes:

```text
origSize (8-byte LE) = 6b 0c 00 00 00 00 00 00   // 3179 bytes

ciphertext[0:32] =
  dcb1367fc380378657fc11e7b968b3a6
  28bff7d8aca261ceca0591bcd81b0075
```

Decrypting `ciphertext[0:32]` with AES-ECB using the derived key yields:

```text
plaintext[0:32] =
  50 4b 03 04 14 00 00 00 08 00 00 00 21 00 ae 40
  8d 88 30 01 00 00 ad 03 00 00 13 00 00 00 5b 43
```

The first 4 bytes are `50 4b 03 04` (`PK\x03\x04`), confirming the Standard AES `EncryptedPackage`
payload is **AES-ECB** encrypted ZIP bytes (then truncated to `origSize`).

### 8.2) AES-128 key derivation sanity check (shows `CryptDeriveKey` is not truncation)

A common bug is to treat the AES-128 key as a simple truncation:
`key = H_block0[0:16]`. Standard AES does **not** do that: the `CryptDeriveKey` ipad/opad expansion
is still applied.

Using the §8.1 values:

```text
H_block0[0:16] =
  11958e53fc62fdebc0107e2fae865014

key (AES-128, 16 bytes; CryptDeriveKey result) =
  5e8727d6c94408a903aececf1382b380
```

### 8.3) RC4 per-block key example (128-bit)

If the file uses **RC4** (`AlgID = CALG_RC4`) with a 128-bit key (`KeySize = 128` bits → 16 bytes),
then (for SHA‑1) the per-block RC4 key is simply the first 16 bytes of `H_block`:

```text
rc4_key(block=0) = H_block0[0:16] =
  6ad7dedf2da3514b1d85eabee069d47d

H_block1 = SHA1(H_final || LE32(1)) =
  2ed4e8825cd48aa4a47994cda7415b4a9687377d

rc4_key(block=1) = H_block1[0:16] =
  2ed4e8825cd48aa4a47994cda7415b4a
```

### 8.4) Verifier check example (AES-ECB)

The following values are a *synthetic* verifier example that is consistent with the derivation above
(AES-256 + SHA-1). It is useful as an end-to-end unit test for the verifier logic.

Inputs:

```text
VerifierPlain (16 bytes) =
  00 11 22 33 44 55 66 77 88 99 aa bb cc dd ee ff

SHA1(VerifierPlain) =
  73 9e 0e 84 90 ea cb cb 2e a1 1d 4a 5d be fb ae 88 8b 09 2e

VerifierHashPlainPadded (32 bytes) =
  73 9e 0e 84 90 ea cb cb 2e a1 1d 4a 5d be fb ae 88 8b 09 2e
  00 00 00 00 00 00 00 00 00 00 00 00
```

Ciphertext (AES-ECB encrypted with `key` from the main example):

```text
EncryptedVerifier (16 bytes) =
  25 89 ae bb 86 0b 8a 41 fa 69 0e 76 ed 56 b0 be

EncryptedVerifierHash (32 bytes) =
  de dd b6 b2 9f 9c d3 82 a0 a2 04 a5 1b 7f df 7e
  b1 23 3f 14 fe 5c 99 d9 6f 5d e4 db 08 3d 92 5e
```

Verification procedure:

1. Decrypt `EncryptedVerifier || EncryptedVerifierHash` with AES-ECB using `key`.
2. Split plaintext into `VerifierPlain` (first 16 bytes) and `VerifierHashPlain` (next 20 bytes).
3. Compute `SHA1(VerifierPlain)` and compare to `VerifierHashPlain`.

If your implementation produces different bytes for this example, the most likely causes are:

* Off-by-one in the 50,000-iteration loop.
* Incorrect UTF‑16LE password encoding (BOM or NUL terminator accidentally included).
* Reversed concatenation order (`Hash(block || H)` vs `Hash(H || block)`).
* Incorrect `CryptDeriveKey` expansion (ipad/opad must use bytes `0x36` and `0x5c`).

---

## 9) End-to-end decryption pseudocode (standard OOXML wrapper)

This appendix ties together parsing, key derivation, verifier validation, and `EncryptedPackage`
decryption.

```text
function decrypt_standard_ooxml(ole, password):
  encInfoBytes = ole.read_stream("EncryptionInfo")
  encPkgBytes  = ole.read_stream("EncryptedPackage")

  // ---------- parse EncryptionInfo ----------
  major = U16LE(encInfoBytes[0:2])
  minor = U16LE(encInfoBytes[2:4])
  flags = U32LE(encInfoBytes[4:8])

  if minor != 2:
    error("not Standard/CryptoAPI")

  headerSize = U32LE(encInfoBytes[8:12])
  header     = encInfoBytes[12 : 12+headerSize]
  verifier   = encInfoBytes[12+headerSize : ]

  // parse EncryptionHeader fixed 8 DWORDs
  algId      = U32LE(header[ 8:12])       // cipher
  algIdHash  = U32LE(header[12:16])       // hash
  keySizeBits = U32LE(header[16:20])      // stored in bits
  if algId == CALG_RC4 and keySizeBits == 0:
    keySizeBits = 40                      // MS-OFFCRYPTO: RC4 KeySize=0 means 40-bit
  keyLen     = keySizeBits / 8

  // parse EncryptionVerifier
  saltSize = U32LE(verifier[0:4])
  salt     = verifier[4 : 4+saltSize]
  encVer   = verifier[4+saltSize : 4+saltSize+16]
  hashSize = U32LE(verifier[4+saltSize+16 : 4+saltSize+20])
  encVerHash = verifier[4+saltSize+20 : ]    // rest of stream (often padded)

  // ---------- derive password hash (H_final) ----------
  pw = UTF16LE(password)    // no BOM, no terminator
  H = Hash( salt || pw )
  for i in 0..49999:
    H = Hash( LE32(i) || H )
  H_final = H

  // ---------- derive file/verifier key (block = 0) ----------
  H_block0 = Hash( H_final || LE32(0) )
  if algId == CALG_RC4:
    keyMaterial = H_block0[0:keyLen]
    if keySizeBits == 40:
      fileKey = keyMaterial || 0x00 * 11  // 16 bytes total
    else:
      fileKey = keyMaterial
  else:
    fileKey = CryptDeriveKey(Hash, H_block0, keyLen)

  // ---------- verify password ----------
  C = encVer || encVerHash
  if algId is AES:
    P = AES_ECB_Decrypt(fileKey, C)
  else if algId == CALG_RC4:
    P = RC4_Decrypt_Stream(fileKey, C)
  else:
    error("unsupported cipher")

  verifierPlain = P[0:16]
  verifierHashPlain = P[16 : 16+hashSize]
  if Hash(verifierPlain)[0:hashSize] != verifierHashPlain:
    error("invalid password")

  // ---------- decrypt EncryptedPackage ----------
  origSizeLo = U32LE(encPkgBytes[0:4])
  origSizeHi = U32LE(encPkgBytes[4:8])
  origSize = origSizeLo | (origSizeHi << 32)
  ciphertext = encPkgBytes[8:]

  if algId is AES:
    if origSize == 0:
      return []
    if len(ciphertext) % 16 != 0:
      error("invalid ciphertext length")

    out = AES_ECB_Decrypt(fileKey, ciphertext)
    if len(out) < origSize:
      error("truncated ciphertext")
    return out[0:origSize]

  else if algId == CALG_RC4:
    // RC4 has no padding; decrypt exactly origSize bytes.
    out = ciphertext[0:origSize]
    for blockIndex, block in enumerate(chunks(out, 0x200)):
      H_block = Hash(H_final || LE32(blockIndex))
      keyMaterial = H_block[0:keyLen]
      if keySizeBits == 40:
        rc4Key = keyMaterial || 0x00 * 11  // 16 bytes total
      else:
        rc4Key = keyMaterial
      block   = RC4_Decrypt_Block(rc4Key, block)
    return out
```
