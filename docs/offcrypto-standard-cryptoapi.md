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

For Agile (4.4) OOXML password decryption details (different scheme; XML descriptor + `dataIntegrity`
HMAC), see `docs/22-ooxml-encryption.md`.

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
4. Standard encryption (as produced by Excel/Office 2007-era “Standard” encryption) is identified by:

```text
major = 3
minor = 2
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

The header starts with eight DWORDs, followed by a UTF‑16LE NUL‑terminated CSP name string.

```text
struct EncryptionHeader {
  u32 Flags;
  u32 SizeExtra;      // should be 0
  u32 AlgID;          // cipher ALG_ID (RC4 / AES_*)
  u32 AlgIDHash;      // hash ALG_ID (SHA1 / MD5)
  u32 KeySize;        // in *bits*
  u32 ProviderType;   // typically PROV_RSA_FULL (1) or PROV_RSA_AES (24)
  u32 Reserved1;      // should be 0
  u32 Reserved2;      // should be 0
  u16 CSPName[];      // UTF-16LE string including NUL terminator; consumes the remaining bytes
}
```

Notes:

* `CSPName` is only metadata (e.g. `Microsoft Enhanced RSA and AES Cryptographic Provider`) and is not
  required to derive keys.
* Prefer parsing the header as: fixed 32 bytes + “the rest is CSPName”.
  Do **not** search for the NUL terminator to determine header length; use `HeaderSize`.

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

**RC4 40-bit interoperability note:** CryptoAPI/Office represent a “40-bit” RC4 key as a 128-bit
RC4 key with the low 40 bits set and the remaining 88 bits zero. Concretely, when `KeySize = 40`,
the RC4 key bytes passed into the RC4 KSA are:

```text
rc4_key = H_block[0..5] || 0x00 * 11   // 16 bytes total
```

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
    * **AES** uses the same derived file key (`block = 0`) for the entire stream and decrypts the
      ciphertext with **AES-ECB** (no IV) (see §7.2.1).
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

#### 5.2.1) RC4 (`CALG_RC4`): key = truncate(`H_block`)

For Standard **RC4** encryption, the per-block RC4 key is derived by truncating the hash:

```text
rc4_key(block) = H_block[0 : keyLen]    // keyLen = KeySize/8
```

This matches the common Standard RC4 “re-key per 0x200-byte block” scheme (see §7.2.2).

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
(AES-256), and `SHA1(inner||outer)` yields 40 bytes.

### 5.3) IV derivation: none for Standard AES-ECB (compatibility notes)

RC4 is a stream cipher and has no IV.

For the most common **Standard AES** files (including the repo’s `msoffcrypto-tool`-generated
fixtures), both:

* the password verifier fields, and
* the `EncryptedPackage` ciphertext

are decrypted with **AES-ECB**, which uses **no IV**.

Some third-party producers and internal tooling use **non-standard AES-CBC** layouts for
`EncryptedPackage` (segmented or stream-CBC). If you need maximum compatibility, one commonly
encountered IV derivation for CBC-segmented variants is:

```text
IV_full = Hash( Salt || LE32(segmentIndex) )
IV = IV_full[0:16]   // AES block size
```

This IV derivation is **not used** by the baseline Standard AES-ECB algorithm described in §7.2.1.

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
keyLen   = KeySize / 8

if AlgID == CALG_RC4:
  key = H_block0[0:keyLen]                                        // §5.2.1
else:
  key = CryptDeriveKey(Hash, H_block0, keyLen=keyLen)             // §5.2.2
```

### 6.2) Decrypt verifier + verifier-hash as a *single* stream

This is the most common implementation bug:

> **Decrypt `EncryptedVerifier` and `EncryptedVerifierHash` together as one ciphertext stream.**

This matters most for **RC4**, where decrypting the two buffers separately would reset the keystream.
For **AES-ECB**, decrypting separately happens to be equivalent (ECB has no chaining), but treating it
as one stream keeps the implementation uniform across algorithms.

Steps:

1. Concatenate ciphertext:

   ```text
   C = EncryptedVerifier || EncryptedVerifierHash
   ```

2. Decrypt `C` with the cipher indicated by `EncryptionHeader.AlgID` using the derived `key`:

   * **AES (`CALG_AES_*`)**: AES-ECB over 16-byte blocks (no IV, no padding at the crypto layer).
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
u64le OriginalPackageSize
u8    EncryptedBytes[...]
```

After decryption, truncate the plaintext to exactly `OriginalPackageSize` bytes (the ciphertext is padded).

### 7.2) Encryption model (AES vs RC4)

#### 7.2.1) AES (`CALG_AES_*`): AES-ECB (no IV, no segmenting)

For baseline Standard/CryptoAPI AES encryption, the `EncryptedPackage` ciphertext (after the 8-byte
size prefix) is **AES-ECB** encrypted with the derived `fileKey` (`block = 0`).

Algorithm:

1. Read `OriginalPackageSize` (`u64le`).
2. Let `C = EncryptedBytes` (all remaining bytes).
3. Require `len(C) % 16 == 0`.
4. `P = AES-ECB-Decrypt(fileKey, C)`.
5. Return `P[0:OriginalPackageSize]` (truncate to the declared size).

Notes:

* The `EncryptedPackage` stream can be larger than `OriginalPackageSize` due to block padding and/or
  OLE sector slack. **Always truncate** to the declared size after decrypting.
* Some producers pad the ciphertext to a fixed size (e.g. to 4096 bytes for very small packages).
  Truncation handles this.
* AES-ECB has no IV. If you see per-segment IV derivation in other code, that is for an alternative
  (non-ECB) scheme (see §5.3 compatibility note).

#### 7.2.2) RC4 (`CALG_RC4`)

RC4-based Standard encryption uses **512-byte** blocks and resets the RC4 keystream per block:

```text
segmentSize = 0x200   // 512
```

For segment index `i = 0, 1, 2, ...`:

1. Derive `H_block = Hash(H_final || LE32(i))`.
2. `key_i = H_block[0 : KeySize/8]` (truncate to the configured key size).
3. Initialize RC4 with `key_i` (fresh state for each segment) and decrypt exactly one segment of
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

// If using CryptoAPI RC4 with KeySize=40, the per-block RC4 key for block=0 would be:
rc4_key_block0_40bit = 6ad7dedf2d0000000000000000000000

key (32 bytes, CryptDeriveKey expansion) =
  de5451b9dc3fcb383792cbeec80b6bc3
  0795c2705e075039407199f7d299b6e4

`EncryptedPackage` for baseline Standard AES is decrypted with **AES-ECB** and uses **no IV**.

Optional (CBC-segmented variant only): if you encounter an `EncryptedPackage` encrypted with a
per-segment IV derived as `IV = SHA1(salt || LE32(segmentIndex))[0:16]`, then:

```text
iv0 (segmentIndex=0) =
  719ea750a65a93d80e1e0ba33a2ba0e7
```

### 8.1) AES-128 key derivation sanity check (shows `CryptDeriveKey` is not truncation)

This second vector is useful because it demonstrates a common bug:
**for AES-128, you still must run the `CryptDeriveKey` ipad/opad step**, even though the desired key
length (16 bytes) is smaller than the SHA‑1 digest length (20 bytes).

Parameters:

* Hash algorithm: SHA‑1
* Cipher: AES‑128
* KeySize: 128 bits → 16 bytes
* spinCount: 50,000
* `block = 0`

Inputs:

```text
password = "Password1234_"
passwordUtf16le =
  50 00 61 00 73 00 73 00 77 00 6f 00 72 00 64 00 31 00 32 00 33 00 34 00 5f 00

salt =
  e8 82 66 49 0c 5b d1 ee bd 2b 43 94 e3 f8 30 ef
```

Derived values:

```text
H_final  = a00d5360ec463ee782df8c267525ae9ac66cd605
H_block0 = e2f8cde457e5d449eb205057c88d201d14531ff3

key (AES-128, 16 bytes; CryptDeriveKey result) =
  40b13a71f90b966e375408f2d181a1aa

Optional (CBC-segmented variant only): with `IV = SHA1(salt || LE32(segmentIndex))[0:16]`:

```text
iv0 (segmentIndex=0) =
  a1cdc25336964d314dd968da998d05b8
```

Sanity check:

```text
H_block0[0:16] =
  e2f8cde457e5d449eb205057c88d201d
```

This is **not** the AES-128 key; if your code uses `H_block0[0:16]` directly, you will derive the
wrong key.

### 8.2) RC4 per-block key example (128-bit)

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

### 8.3) Verifier check example (AES-ECB)

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
  keySizeBit = U32LE(header[16:20])
  keyLen     = keySizeBit / 8

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
    fileKey = H_block0[0:keyLen]
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
  origSize = U64LE(encPkgBytes[0:8])
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
      rc4Key  = H_block[0:keyLen]
      block   = RC4_Decrypt_Block(rc4Key, block)
    return out
```
