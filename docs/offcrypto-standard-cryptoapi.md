# MS-OFFCRYPTO Standard (CryptoAPI) encryption: key derivation + verifier validation

This document is a *from-scratch* implementation guide for decrypting **MS Office “Standard” encryption**
(`EncryptionInfo` version **3.2**) used by password-protected OOXML files (e.g. `.xlsx`, `.docx`, `.pptx`)
stored inside an OLE Compound File.

It focuses on:

1. Detecting **Standard (CryptoAPI)** encryption.
2. Parsing the **binary** `EncryptionInfo` stream layout.
3. Deriving keys using the fixed **spinCount = 50,000** password hashing loop.
4. Implementing **CryptoAPI `CryptDeriveKey`** (ipad/opad expansion with `0x36` / `0x5c`).
5. Validating the password by decrypting and checking the **verifier**.

The intent is that an engineer can implement this without reading any external references.

---

## 1) Detecting “Standard” encryption (version 3.2)

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
major = 3
minor = 2
```

The `flags` field is part of the `EncryptionInfo` header (and must be consumed to keep offsets correct),
but it is not needed for the scheme dispatch.

If the version is not `3.2`, the file is *not* Standard encryption (it may be Agile, Extensible, etc.).

---

## 2) `EncryptionInfo` binary layout (version 3.2)

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
| 0x00   | 2    | u16le  | Major | 3 for Standard |
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

### 3.2) Hash `AlgIDHash` values

| Hash | Name | ALG_ID (hex) | Digest bytes |
|------|------|--------------|--------------|
| MD5 | `CALG_MD5` | `0x00008003` | 16 |
| SHA-1 | `CALG_SHA1` | `0x00008004` | 20 |

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
  * `block = 0` is used for the **password verifier** key.
  * Some producers also derive per-segment keys for `EncryptedPackage` using `block = segmentIndex`
    (e.g. `block = 0, 1, 2, ...` for successive segments).

### 5.1) Per-block hash input (key material)

Compute:

```text
H_block = Hash( H_final || LE32(block) )
```

This `H_block` is not yet the symmetric key; it is the hash input to `CryptDeriveKey`.

### 5.2) Implementing CryptoAPI `CryptDeriveKey` (ipad/opad expansion)

CryptoAPI’s `CryptDeriveKey` is **not PBKDF2**. For MD5/SHA1 it behaves like:

* If the needed key length is **≤ digest length**: use the **prefix** of the hash bytes.
* If the needed key length is **> digest length**: expand with an ipad/opad construction.

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

  if keyLen <= digestLen:
    return H_block[0:keyLen]

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
  return derived[0:keyLen]
```

This is sufficient for Standard encryption because Office only requests up to 32 bytes of key material
(AES-256), and `SHA1(inner||outer)` yields 40 bytes.

### 5.3) IV derivation (AES-CBC)

RC4 is a stream cipher and has no IV.

For AES-based Standard encryption, data is commonly encrypted in **independent AES-CBC segments**.
Each segment `block` uses an IV derived from the verifier `Salt`:

```text
IV_full = Hash( Salt || LE32(block) )
IV = IV_full[0:16]   // AES block size
```

This IV formula is used both for:

* the verifier (`block = 0`), and
* `EncryptedPackage` segments (`block = 0, 1, 2, ...`) when using per-segment IVs.

---

## 6) Password verifier validation (critical correctness details)

Once you can derive `key` for `block = 0`, you can check whether the password is correct.

Inputs from `EncryptionVerifier`:

* `Salt`
* `EncryptedVerifier` (16 bytes)
* `VerifierHashSize` (16 or 20)
* `EncryptedVerifierHash` (remaining bytes; AES ciphertext may include padding)

### 6.1) Derive block-0 key (+ IV for AES)

```text
H_final  = hash_password(password, Salt, spinCount=50000)         // §4
H_block0 = Hash( H_final || LE32(0) )                             // §5.1
key      = CryptDeriveKey(Hash, H_block0, keyLen=KeySize/8)       // §5.2
iv       = Hash(Salt || LE32(0))[0:16]    // AES only              // §5.3
```

### 6.2) Decrypt verifier + verifier-hash as a *single* stream

This is the most common implementation bug:

> **Decrypt `EncryptedVerifier` and `EncryptedVerifierHash` together as one ciphertext stream.**

Steps:

1. Concatenate ciphertext:

   ```text
   C = EncryptedVerifier || EncryptedVerifierHash
   ```

2. Decrypt `C` with the cipher indicated by `EncryptionHeader.AlgID` using the derived `key`:

   * **AES**: AES-CBC with `iv` from §6.1 (and `NoPadding` at the crypto layer; padding/truncation
     is handled by sizes in the container).
   * **RC4**: RC4 stream cipher with the derived key.

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

### 7.2) Segment encryption model

In practice, Office decryptors process `EncryptedPackage` in **4096-byte plaintext segments**:

```text
segmentSize = 0x1000   // 4096
```

For segment index `i = 0, 1, 2, ...`:

1. Use the **same derived key** (`block = 0`) from §6.1 for all segments.
2. Derive `iv_i = Hash(Salt || LE32(i))[0:16]` (same derivation as §5.3 / §6.1, but with `block=i`).
3. Decrypt the segment ciphertext with AES-CBC(key, iv_i).
4. Append the decrypted bytes and continue until you have at least `OriginalPackageSize` bytes.

Finally, **truncate** the concatenated plaintext to `OriginalPackageSize` bytes.

---

## 8) Worked example (test vector)

This example is intentionally small and deterministic. It is **not** a full Office file; it is just the key
derivation math that you can use as a unit test.

Parameters:

* Hash algorithm: SHA-1 (`CALG_SHA1`, `0x00008004`)
* Cipher: AES-256 (`CALG_AES_256`, `0x00006610`)
* KeySize: 256 bits → 32 bytes
* spinCount: 50,000
* `block = 0` (Standard uses a fixed `LE32(0)` block key)

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

key (32 bytes, CryptDeriveKey expansion) =
  de5451b9dc3fcb383792cbeec80b6bc3
  0795c2705e075039407199f7d299b6e4

iv0 (AES-CBC, 16 bytes; block=0) =
  719ea750a65a93d80e1e0ba33a2ba0e7
```

If your implementation produces different bytes for this example, the most likely causes are:

* Off-by-one in the 50,000-iteration loop.
* Incorrect UTF‑16LE password encoding (BOM or NUL terminator accidentally included).
* Reversed concatenation order (`Hash(block || H)` vs `Hash(H || block)`).
* Incorrect `CryptDeriveKey` expansion (ipad/opad must use bytes `0x36` and `0x5c`).
