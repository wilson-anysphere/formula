# MS-OFFCRYPTO Standard Encryption (CryptoAPI RC4)

This document is the developer-facing reference for the **Standard / CryptoAPI / RC4** encryption
scheme described in **[MS-OFFCRYPTO]**.

It is intentionally written so a future contributor can re-implement (or audit) the algorithm
without needing to hunt through the spec PDF, blog posts, or other libraries.

Scope:

- **OOXML password-to-open** files stored as an OLE/CFB container with `EncryptionInfo` +
  `EncryptedPackage` streams.
- Specifically the **RC4** variant of *Standard Encryption* (not Agile encryption, not legacy BIFF
  `FILEPASS`).

References:

- **MS-OFFCRYPTO** (Office Document Cryptography Structure):  
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
- **CryptoAPI algorithm IDs / `CALG_*` constants** (WinCrypt):  
  https://learn.microsoft.com/en-us/windows/win32/seccrypto/cryptographic-algorithm-identifiers

Normative spec sections (MS-OFFCRYPTO) referenced by this doc:

- **§2.3.4.4** — `\EncryptedPackage` stream layout (8-byte plaintext size prefix)
- **§2.3.4.5** — `\EncryptionInfo` stream layout (Standard / CryptoAPI; `versionMinor == 2`)
- **§2.3.4.7** — password hashing + key generation (fixed **50,000** iterations)

## Container and stream layout (what’s on disk)

Password-encrypted OOXML workbooks are **not ZIP files on disk**, even if the extension is
`.xlsx`/`.xlsm`/`.xlsb`.

Instead, they are an **OLE Compound File Binary** (CFB) container with two relevant streams:

- `EncryptionInfo` — parameters and verifier (Standard header structures for `versionMinor == 2`;
  commonly `3.2`)
- `EncryptedPackage` — the encrypted bytes of the real OPC/ZIP package

### `EncryptedPackage` stream layout

The `EncryptedPackage` stream layout is:

```text
8 bytes   original_size header (plaintext; see note below)
N bytes   ciphertext (encrypted OPC/ZIP bytes)
...       optional trailing bytes / padding (OLE sector slack, producer quirks)
```

Important details:

- `original_size` is the size **of the decrypted package bytes** (the ZIP payload).
- In the wild, different implementations interpret the 8-byte prefix either as:
  - `u32 totalSize` + `u32 reserved` (where `reserved` is often `0`), or
  - a single `u64` little-endian size.

  To be compatible with both, parse the prefix as two little-endian DWORDs and combine into a
  64-bit value:

  ```text
  lo = u32le(bytes[0..4])
  hi = u32le(bytes[4..8])
  original_size = lo as u64 | ((hi as u64) << 32)
  ```
- The ciphertext may contain extra trailing bytes beyond `original_size` (e.g. OLE sector padding).
  Callers should decrypt and then **truncate to `original_size`**.
- `original_size` is *not* a crypto padding indicator (RC4 is a stream cipher; there is no PKCS#7).

## RC4 block size and re-keying

Standard RC4 encryption re-keys the RC4 stream cipher frequently. The ciphertext is processed in
**blocks of 0x200 bytes (512 bytes)**:

- Block index `0` decrypts bytes `[0x000..0x1FF]` of the ciphertext (i.e. the first 512 ciphertext
  bytes **after** the 8-byte `original_size` prefix).
- Block index `1` decrypts the next 512 bytes, and so on.

### Why 0x200 and not 0x400?

Be careful not to copy/paste the legacy BIFF8 RC4 implementation:

- **MS-OFFCRYPTO Standard RC4 (`EncryptedPackage`)** uses **0x200-byte** blocks.
- **MS-XLS BIFF8 RC4 (`FILEPASS`)** uses **0x400-byte** blocks.

If you use the wrong block size, you will derive the wrong per-block keys and the decrypted bytes
will diverge after the first re-key boundary.

## Key derivation (password → per-block RC4 key)

Notation:

- `Hash()` is the hash algorithm specified by `EncryptionHeader.algIdHash`
  (commonly `CALG_SHA1`).
- `LE32(x)` is the 4-byte little-endian encoding of `x` (an unsigned 32-bit integer).
- `||` is byte concatenation.

Inputs (from `EncryptionInfo`):

- `password`: user-supplied password string
- `salt`: `EncryptionVerifier.salt` bytes (`EncryptionVerifier.saltSize` is typically 16)
- `spin_count`: **50000** for Standard CryptoAPI RC4 (`0x0000C350`)
- `key_size_bits`: `EncryptionHeader.keySize` (stored in *bits*)
  - Example: 0x80 for 128-bit RC4 keys.
  - **KeySize=0 note (RC4 only):** MS-OFFCRYPTO specifies `keySize == 0` MUST be interpreted as
    **40** (legacy 40-bit RC4).
- `key_size_bytes`: `key_size_bits / 8` (after applying the `keySize == 0 → 40` rule above)

### Step 1: encode the password

Encode the password as UTF-16LE bytes:

```text
pw_bytes = UTF16LE(password)    // no BOM, no terminator
```

This is the same encoding used by `formula-xlsx`’s Agile encryption helpers.

### Step 2: “spin” the password hash (50000 iterations)

```text
h = Hash(salt || pw_bytes)
for i in 0..spin_count-1:
  h = Hash(LE32(i) || h)
```

`h` is the “spun” password hash (for SHA-1, 20 bytes).

### Step 3: derive a per-block RC4 key

For each ciphertext block index `block` (0-based):

```text
h_block = Hash(h || LE32(block))
rc4_key = h_block[0..key_size_bytes]   // key_size_bytes = key_size_bits/8 (for RC4, key_size_bits==0 means 40)
```

Then decrypt exactly 0x200 ciphertext bytes using RC4 with `rc4_key` (reset RC4 state per block).

**40-bit note:** MS-OFFCRYPTO specifies `keySize == 0` MUST be interpreted as 40-bit RC4. In both
cases (“0” or “40”), `key_size_bytes == 5`; use the 5-byte key directly
(`rc4_key = h_block[0..5]`). Do **not** pad it to 16 bytes.

## Password verification (EncryptionVerifier)

Standard CryptoAPI stores a verifier to check whether the derived key is correct before attempting
to parse the decrypted ZIP bytes.

High-level flow:

1. Derive `rc4_key` for `block = 0` using the steps above.
2. Initialize an RC4 stream with that key.
3. Decrypt:
   - `encryptedVerifier` (16 bytes)
   - then `encryptedVerifierHash` (use the **same RC4 stream**, continuing the keystream)
4. Compute `Hash(verifier_plaintext)` and compare it to the decrypted verifier hash.

Notes:

- `EncryptionVerifier.verifierHashSize` indicates how many bytes of the hash are meaningful
  (commonly 20 for SHA-1). Some producers may store/pad the encrypted hash field beyond that.
- Treat a mismatch as **wrong password**, not “corrupt file”.

## Test vectors (used in unit tests)

The unit test `crates/formula-io/tests/offcrypto_standard_rc4_vectors.rs` uses the following
deterministic vectors to lock down the derivation details (UTF-16LE encoding, spin loop ordering,
and LE32 block index encoding) for multiple `EncryptionHeader.algIdHash` values.

### SHA-1 (`CALG_SHA1`)

Parameters:

- password: `"password"`
- salt (hex): `00 01 02 03 04 05 06 07 08 09 0a 0b 0c 0d 0e 0f`
- spinCount: `50000` (`0x0000C350`)
- keySize: `16` bytes (128-bit RC4)

Expected values:

```text
spun password hash (H) =
  1b5972284eab6481eb6565a0985b334b3e65e041

block 0 rc4_key =
  6ad7dedf2da3514b1d85eabee069d47d

block 1 rc4_key =
  2ed4e8825cd48aa4a47994cda7415b4a

RC4(key=block0, plaintext=\"Hello, RC4 CryptoAPI!\") ciphertext =
  e7c9974140e69857dbdec656c7ccb4f9283d723236
```

### 40-bit RC4 key length example (5-byte key; do not pad to 16 bytes)

Using the same parameters as above but with `KeySize = 0` (interpreted as 40) or `KeySize = 40`
bits (`key_size_bytes = 5`):

```text
block 0 H_block (SHA1(H || LE32(0))) =
  6ad7dedf2da3514b1d85eabee069d47dd058967f

key_material (5 bytes) =
  6ad7dedf2d

zero-padded 16-byte key (a common legacy behavior; **incorrect for MS-OFFCRYPTO Standard RC4**) =
  6ad7dedf2d0000000000000000000000

RC4(key=raw 5-byte key_material, plaintext=\"Hello, RC4 CryptoAPI!\") ciphertext =
  d1fa444913b4839b06eb4851750a07761005f025bf

RC4(key=zero-padded 16-byte key, plaintext=\"Hello, RC4 CryptoAPI!\") ciphertext =
  7a8bd000713a6e30ba9916476d27b01d36707a6ef8
```

These ciphertexts differ, demonstrating why implementations must treat `keySize==0/40` as a **5-byte
RC4 key** (`keyLen = keySize/8`) and must **not** zero-pad it to 16 bytes for Standard RC4.

### 56-bit RC4 key truncation example

Using the same parameters as above but with `KeySize = 56` bits (`key_size_bytes = 7`):

```text
block 0 rc4_key =
  6ad7dedf2da351

block 1 rc4_key =
  2ed4e8825cd48a

RC4(key=block0, plaintext="Hello, RC4 CryptoAPI!") ciphertext =
  883dbf39789abb12c0245ad562f13dd69da9b44660
```

This is a straight truncation of the 128-bit keys: it uses the **first 7 bytes** of the
`SHA1(H || LE32(block))` digest.
### MD5 (`CALG_MD5`)

Some Standard/CryptoAPI files use `CALG_MD5` instead of `CALG_SHA1` for the fixed-spin password
hashing loop and block-key derivation.

Parameters:

- password: `"password"`
- salt (hex): `00 01 02 03 04 05 06 07 08 09 0a 0b 0c 0d 0e 0f`
- spinCount: `50000` (`0x0000C350`)
- algIdHash: `CALG_MD5` (`0x00008003`)
- keySize: `16` bytes (128-bit RC4)

Expected values:

```text
spun password hash (H) =
  2079476089fda784c3a3cfeb98102c7e

block 0 rc4_key =
  69badcae244868e209d4e053ccd2a3bc

block 1 rc4_key =
  6f4d502ab37700ffdab5704160455b47

RC4(key=block0, plaintext="Hello, RC4 CryptoAPI!") ciphertext =
  425dd9c8165e1216065e53eb586e897b5e85a07a6d
```

For the same MD5 parameters but `KeySize = 56` bits (`key_size_bytes = 7`), the per-block keys are:

```text
block 0 rc4_key =
  69badcae244868

block 1 rc4_key =
  6f4d502ab37700

RC4(key=block0, plaintext="Hello, RC4 CryptoAPI!") ciphertext =
  acdabc88ff665d0454d32d952b18e05e8331dfb44e
```

#### MD5 + 40-bit key length example (5-byte key; do not pad to 16 bytes)

This vector covers the `CALG_MD5` variant and guards against the common mistake of treating a
40-bit key as a 16-byte (zero-padded) RC4 key. For MD5, note that
`EncryptionVerifier.verifierHashSize` is **16** bytes.

Parameters:

- password: `"password"`
- salt (hex): `00 01 02 03 04 05 06 07 08 09 0a 0b 0c 0d 0e 0f`
- spinCount: `50000` (`0x0000C350`)
- algIdHash: `CALG_MD5` (`0x00008003`)
- keySize: `40` bits (5 bytes used for RC4; `keyLen = keySize/8`)

Expected values:

```text
block 0 hash (Hb0 = MD5(H || LE32(0))) =
  69badcae244868e209d4e053ccd2a3bc

key_material (40-bit) =
  69badcae24

RC4(key=key_material, plaintext=\"Hello, RC4 CryptoAPI!\") ciphertext =
  db037cd60d38c882019b5f5d8c43382373f476da28

RC4(key=zero-padded 16-byte key (incorrect), plaintext=\"Hello, RC4 CryptoAPI!\") ciphertext =
  565016a3d8158632bb36ce1d76996628512061bfa3
```

## CryptoAPI constants (for parsing `EncryptionHeader`)

The Standard `EncryptionHeader` uses Windows CryptoAPI IDs, not OIDs.

Common values for Standard RC4-encrypted OOXML:

| Field | Meaning | Typical value |
|------:|---------|---------------|
| `EncryptionHeader.algId` | cipher | `CALG_RC4` = `0x00006801` |
| `EncryptionHeader.algIdHash` | hash | `CALG_SHA1` = `0x00008004` |
| `EncryptionHeader.keySize` | key size (bits) | `0x00000080` (128-bit) |
| `EncryptionHeader.providerType` | CryptoAPI provider | often `PROV_RSA_FULL` (= 1) |

See Microsoft’s WinCrypt reference for the full `CALG_*` table:
https://learn.microsoft.com/en-us/windows/win32/seccrypto/cryptographic-algorithm-identifiers
