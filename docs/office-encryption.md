# Office workbook encryption (MS-OFFCRYPTO / MS-XLS) — implementation reference
This note is **maintainer-facing**. It captures:

- which Excel *workbook encryption* schemes we implement in this repo (OOXML + legacy BIFF),
- the exact parameters / algorithm choices we expect (and would emit if we implement encryption on
  save), and
- the key-derivation / verifier nuances that commonly cause interoperability bugs.

This document is intentionally **not** user-facing (“how do I open a password-protected file?”).
For a higher-level overview (detection, UX error semantics), see `docs/21-encrypted-workbooks.md`.

## Terminology / file shapes

### OOXML “password to open” encryption (`.xlsx`/`.xlsm`/`.xlsb`)
Excel **Encrypt with Password** (“password to open”) does **not** encrypt the ZIP container
directly. Instead it writes an **OLE/CFB** (Structured Storage) container with (at minimum):

- `EncryptionInfo` — encryption metadata (binary Standard/CryptoAPI or XML Agile).
- `EncryptedPackage` — encrypted bytes of the *original* OPC/ZIP package.

The decrypted payload of `EncryptedPackage` is the raw `.xlsx`/`.xlsm`/`.xlsb` ZIP bytes.

### Legacy `.xls` BIFF encryption
Legacy `.xls` encryption is signaled via a `FILEPASS` record in the workbook globals stream
([MS-XLS]). This is a different format than the OOXML `EncryptedPackage` wrapper.

## What is implemented in this repo (today)

### Summary table
| Format | Scheme | Marker | Implemented crypto | Notes / entry points |
|---|---|---|---|---|
| OOXML (`.xlsx`/`.xlsm`/`.xlsb`) | **Agile** | `EncryptionInfo` **4.4** | ✅ decrypt (library) + ✅ encrypt (writer), ❌ not yet plumbed through `formula-io` open APIs | `crates/formula-office-crypto` (end-to-end decrypt + Agile writer), `crates/formula-xlsx/src/offcrypto/*` (Agile primitives), `crates/formula-offcrypto` (Agile XML parsing subset) |
| OOXML (`.xlsx`/`.xlsm`/`.xlsb`) | **Standard / CryptoAPI (AES)** | `EncryptionInfo` **3.2** (minor=2; major ∈ {2,3,4} in the wild) | ✅ decrypt (library), ❌ not yet plumbed through `formula-io` open APIs | `crates/formula-office-crypto` (end-to-end decrypt), `crates/formula-offcrypto` (parse + standard key derivation + verifier; stricter alg gating), `crates/formula-io/src/offcrypto/encrypted_package.rs` (decrypt `EncryptedPackage` given key+salt), `docs/offcrypto-standard-encryptedpackage.md` |
| Legacy `.xls` (BIFF8) | **FILEPASS RC4 CryptoAPI** | BIFF `FILEPASS` record | ✅ decrypt when password provided (import API) | `formula_xls::import_xls_path_with_password`, `crates/formula-xls/src/decrypt.rs` |

Important: `formula-io`’s public open APIs currently **detect** encryption and surface dedicated
errors (OOXML: `PasswordRequired` / `InvalidPassword` / `UnsupportedOoxmlEncryption`; legacy `.xls`:
`EncryptedWorkbook`), but do not yet decrypt encrypted OOXML workbooks end-to-end. The crypto code
above exists to enable that future integration.

## Supported schemes / parameter subsets

### OOXML: Agile encryption (4.4)
We implement the password key encryptor subset of Agile (`keyEncryptor uri=".../password"`).

Supported parameters (as enforced by decryptors today):

- `cipherAlgorithm`: `AES`
- `cipherChaining`: `ChainingModeCBC`
- `hashAlgorithm`: `SHA1` / `SHA256` / `SHA384` / `SHA512`
- `keyBits`: `128` / `192` / `256` (must match AES key size)
- `blockSize`: `16` (AES block size; parsed from file and used for IV truncation)
- `saltSize`: typically `16` (parsed; we do not require a fixed value but many writers use 16)
- `spinCount`: file-provided `u32` (guard with a DoS max; see [Spin count limits](#spin-count-dos-limits))

Not implemented:

- certificate-based key encryptors
- IRM / “DataSpaces” transforms

### OOXML: Standard encryption (CryptoAPI; `versionMinor == 2`)
Supported subset (CryptoAPI AES):

- Cipher: AES-128/192/256 (`CALG_AES_128/192/256`).
- Hash (`AlgIDHash`) for password hashing / key derivation:
  - `crates/formula-office-crypto` supports `CALG_SHA1`/`CALG_SHA_256`/`CALG_SHA_384`/`CALG_SHA_512`.
  - `crates/formula-offcrypto` intentionally gates to **SHA-1** only (Excel default) and rejects
    other hashes as “unsupported algorithm”.
- Salt size: file-provided (`EncryptionVerifier.saltSize`, typically 16).

Other combinations (RC4, mismatched key sizes, etc.) are treated as unsupported by the current code.

Note on version gating in helper APIs:

- Most parsers in this repo treat Standard as **`versionMinor == 2`** with `versionMajor ∈ {2,3,4}`.
- Some convenience APIs are intentionally stricter:
  - `formula_offcrypto::decrypt_standard_ooxml_from_bytes` currently requires **exactly `3.2`**
    (it is a thin wrapper over the upstream `office_crypto` crate).
  - For best compatibility across Standard variants, prefer `crates/formula-office-crypto`’s
    decryptor.

### Legacy `.xls`: BIFF8 `FILEPASS` RC4 CryptoAPI
Currently supported in `formula-xls`:

- BIFF8 `FILEPASS` with `wEncryptionType=0x0001` (RC4) and `wEncryptionSubType=0x0002` (CryptoAPI)
- RC4 with SHA-1 and 50,000 password-hash iterations (see [Legacy `.xls` key derivation](#legacy-xls-biff8-filepass-rc4-cryptoapi))

Not implemented:

- XOR obfuscation
- other BIFF encryption variants (including AES CryptoAPI for BIFF)

## Container format details (what’s inside the OLE file)

### `EncryptionInfo` (common header)
Both Standard and Agile `EncryptionInfo` start with:

```text
u16 versionMajor
u16 versionMinor
u32 flags
u32 headerSize   // Standard: byte length of `EncryptionHeader`; Agile: byte length of the XML descriptor
... scheme-specific payload
```

The version pair is used to dispatch:
* **minor == 2** and major ∈ {2,3,4} ⇒ Standard (CryptoAPI)
* **4.4** ⇒ Agile

### `EncryptedPackage` stream framing
`EncryptedPackage` begins with an **8-byte little-endian** unsigned integer:

```text
u64 original_size
u8  encrypted_bytes[...]
```

After decryption, the plaintext stream is truncated to `original_size`.

## Key derivation + verification (nuances that matter)

### Password encoding (both schemes)
The password is converted to UTF-16LE bytes **without** a trailing NUL. (I.e. the same byte sequence
as `password.encode_utf16().flat_map(|u| u.to_le_bytes())` in Rust.)

Empty passwords are valid.

### Agile: password KDF (spin count loop)
Agile uses a custom iterative hash KDF (not PBKDF2):

```text
pw = UTF16LE(password)
H0 = Hash(salt + pw)
H  = H0
for i in 0..spinCount:
  H = Hash(LE32(i) + H)

derived = Hash(H + blockKey)
key = derived[0..keyBits/8]
```

Notes:
* `Hash` is the `hashAlgorithm` declared by the relevant XML node (`SHA1`/`SHA256`/`SHA384`/`SHA512`).
* The iteration counter `i` is a **little-endian u32**.

#### Agile blockKey constants
Agile defines several 8-byte `blockKey` constants. We use the canonical values from MS-OFFCRYPTO:

```text
VERIFIER_HASH_INPUT_BLOCK = FE A7 D2 76 3B 4B 9E 79
VERIFIER_HASH_VALUE_BLOCK = D7 AA 0F 6D 30 61 34 4E
KEY_VALUE_BLOCK           = 14 6E 0B E7 AB AC D0 D6
HMAC_KEY_BLOCK            = 5F B2 AD 01 0C B9 E1 F6
HMAC_VALUE_BLOCK          = A0 67 7F 02 B2 2C 84 33
```

These `blockKey` constants are the most common source of “almost works” interop bugs: they are not
derivable from the file and must be hard-coded exactly.

#### Agile password verification
The password key encryptor stores:
* `encryptedVerifierHashInput` — AES-CBC-encrypted random verifier bytes
* `encryptedVerifierHashValue` — AES-CBC-encrypted `Hash(verifierBytes)`

To validate a password:
1. Derive `K_input` with `VERIFIER_HASH_INPUT_BLOCK`, decrypt `encryptedVerifierHashInput` ⇒ `V`
2. Derive `K_value` with `VERIFIER_HASH_VALUE_BLOCK`, decrypt `encryptedVerifierHashValue` ⇒ `HV_enc`
3. Compute `HV = Hash(V)` and constant-time-compare `HV[..hashSize]` with `HV_enc[..hashSize]`

#### Agile package key extraction
The actual package encryption key is stored as `encryptedKeyValue`:
1. Derive `K_key` with `KEY_VALUE_BLOCK`
2. AES-CBC-decrypt `encryptedKeyValue` ⇒ `package_key` (truncate to `keyBits/8`)

The package key is cached for streaming decryption (see [Caching](#caching--streaming-decrypt)).

#### Agile package decryption (segment IV derivation)
The package is encrypted in 4096-byte segments. For segment index `i`:

```text
iv = Hash(keyData.saltValue + LE32(i))[0..blockSize]
plaintext_i = AES-CBC-Decrypt(key=package_key, iv=iv, ciphertext=segment_i)
```

Each segment is independently AES-CBC encrypted; segment ciphertext is padded to the AES block size.

#### Agile integrity (HMAC)
Agile includes a `dataIntegrity` block that authenticates the package bytes:

* `encryptedHmacKey` is AES-CBC-decrypted using **`package_key` as the AES key**, and an IV derived
  from `keyData.saltValue` and `HMAC_KEY_BLOCK` (truncate/pad to `blockSize`).
* `encryptedHmacValue` is AES-CBC-decrypted using **`package_key` as the AES key**, and an IV derived
  from `keyData.saltValue` and `HMAC_VALUE_BLOCK` (truncate/pad to `blockSize`).
* Compute `HMAC-(keyData.hashAlgorithm)(hmacKey, EncryptedPackageStreamBytes)` and compare it to the
  decrypted `hmacValue` (constant-time compare).
  - Note: in our implementations, the HMAC input is the **entire `EncryptedPackage` stream bytes**
    (8-byte length prefix + ciphertext), not the decrypted ZIP bytes.

Implementation status:

- `crates/formula-xlsx::offcrypto` validates `dataIntegrity` and returns `IntegrityMismatch` on
  failure.
- `crates/formula-offcrypto` parses `encryptedHmacKey`/`encryptedHmacValue` for completeness but does
  not currently validate them.
- `crates/formula-office-crypto` also parses these fields but does not currently validate them.
- If/when we wire Agile decryption into `formula-io`, we should strongly consider enabling HMAC
  verification by default to distinguish wrong passwords from “ZIP happened to parse”.

### Standard (CryptoAPI): password KDF + verifier (ECMA-376)

Standard encryption key derivation is **not PBKDF2**. It is CryptoAPI/ECMA-376’s iterative hash loop
(`spinCount = 50,000`) plus CryptoAPI `CryptDeriveKey`-style key material expansion.

Hash algorithm nuance:

- The password hashing loop uses the hash identified by `EncryptionHeader.algIdHash`
  (`CALG_SHA1`, `CALG_SHA_256`, …).
- Excel’s Standard/CryptoAPI AES commonly uses **SHA-1**; some non-Excel tooling can emit other
  hashes.

In this repo, the reference implementation is `crates/formula-offcrypto/src/lib.rs`:
`standard_derive_key` + `standard_verify_key`.

`crates/formula-office-crypto` also implements a compatible Standard key deriver
(`StandardKeyDeriver`) but supports additional `AlgIDHash` values (SHA-256/384/512) for
compatibility.

High-level shape (Excel-default SHA-1):

```text
pw = UTF16LE(password)                      // no BOM, no NUL
H  = SHA1(salt || pw)
for i in 0..50000:
  H = SHA1(LE32(i) || H)

Hfinal = SHA1(H || LE32(0))
keyMaterial = SHA1((0x36*64) XOR Hfinal) || SHA1((0x5c*64) XOR Hfinal)
key = keyMaterial[0..keySizeBytes]          // keySizeBytes = keySizeBits/8
```

Verifier nuances (very common bug source):

- `EncryptionVerifier.encryptedVerifier` and `EncryptionVerifier.encryptedVerifierHash` are
  encrypted with **AES-ECB** (no IV) using the derived key.
- The verifier hash is `EncryptionHeader.algIdHash` (commonly SHA-1, 20 bytes) and the encrypted
  blob is padded to an AES block boundary (for SHA-1, typically **32 bytes** on disk).

### Standard (CryptoAPI): `EncryptedPackage` decryption (AES-CBC, 0x1000 segments)

Standard `EncryptedPackage` decryption uses **AES-CBC** over the package payload, segmented into
4096-byte plaintext segments with per-segment IV derivation:

```text
iv_i = SHA1(salt || LE32(i))[0..16]
plaintext_i = AES-CBC-Decrypt(key, iv_i, ciphertext_segment_i)
```

Where:

- `salt` is `EncryptionVerifier.salt` (16 bytes).
- Segment index `i` is 0-based.
- After concatenation, truncate plaintext to the 8-byte `orig_size` prefix.

See `docs/offcrypto-standard-encryptedpackage.md` for edge cases (notably the “extra full padding
block” case where the final ciphertext segment can be `0x1010` bytes).

Implementation nuance:

- `crates/formula-io/src/offcrypto/encrypted_package.rs` implements the CBC + per-segment-IV scheme
  above (given a derived key and salt).
- `crates/formula-office-crypto` attempts a small set of Standard/CryptoAPI AES variants observed in
  the wild (including the scheme above) to maximize compatibility.

## Interop notes / fixture generation

### Which Excel versions produce which scheme?
In practice:

* **Excel 2007** typically produces **Standard (3.2)** encryption.
* **Excel 2010+** typically produces **Agile (4.4)** encryption.

Excel can also vary algorithms via “Advanced” encryption options; our supported subset assumes the
defaults above.

### Generating fixtures (recommended)
To generate fixtures for regression tests:

1. Create a minimal workbook (one sheet, a few cells).
2. In Excel: **File → Info → Protect Workbook → Encrypt with Password**.
3. Save as `.xlsx` (Agile by default on modern Excel).

For Standard fixtures, prefer using a known Excel 2007 install (or a controlled fixture generator
tooling in CI) so the output is stable.

Repo fixtures:

- OOXML encrypted fixtures live in `fixtures/encrypted/ooxml/` (see that directory’s README).
- BIFF8 `.xls` encrypted fixtures for `FILEPASS` live under `crates/formula-xls/tests/fixtures/`.

Note: the Apache POI-based generator under `tools/encrypted-ooxml-fixtures/` can emit Standard
encryption with an `EncryptionInfo` version of `4.2` (still Standard/CryptoAPI; `versionMinor == 2`).

### Inspecting the scheme quickly
To identify the scheme without full parsing:

* Open the file as OLE/CFB and read the first 8 bytes of `EncryptionInfo`.
* Interpret as `u16 major, u16 minor, u32 flags` (little-endian).
  * 4.4 ⇒ Agile
  * `minor == 2` and `major ∈ {2,3,4}` ⇒ Standard (CryptoAPI; commonly 3.2)

### Defaults for *writing* encrypted OOXML (`formula-office-crypto`)
This repo does not yet re-encrypt workbooks on save in the `formula-io` export path (round-tripping
an encrypted workbook will eventually require emitting a new `EncryptionInfo` + `EncryptedPackage`
wrapper).

However, `crates/formula-office-crypto` implements an **Agile encryption writer**
(`formula_office_crypto::encrypt_package_to_ole`) that is used for round-trip tests. Its defaults
are the repo’s current “recommended writer settings” unless a compatibility requirement dictates
otherwise:

- **Agile (default; `EncryptOptions::default()`):**
  - `cipherAlgorithm=AES`, `cipherChaining=ChainingModeCBC`
  - `hashAlgorithm=SHA512`
  - `keyBits=256`, `blockSize=16`, `saltSize=16`
  - `spinCount=100000` (matches common modern Excel output)
  - `EncryptedPackage` segmentation: 4096-byte plaintext segments (`0x1000`)
  - Generate independent CSPRNG salts for:
    - `keyData/@saltValue`
    - password key encryptor `saltValue`
  - Generate a random 16-byte verifier input (`verifierHashInput`).
  - Emit `dataIntegrity` (HMAC) in the XML descriptor (note: not all decrypt paths validate it yet).

- **Standard writer:** not implemented yet in `formula-office-crypto`. If we add it, the intended
  defaults are:
  - AES-128 (`CALG_AES_128`) + SHA-1 (`CALG_SHA1`)
  - Iteration count is effectively fixed at 50,000 in the key derivation.
  - `EncryptedPackage` segmentation: 4096-byte plaintext segments (`0x1000`) and IV derivation
    `SHA1(salt || LE32(i))[0..16]` (see `docs/offcrypto-standard-encryptedpackage.md`).

These defaults are intended for **interoperability** with Excel, not for novel cryptographic
design.

## Legacy `.xls` (BIFF8) `FILEPASS` RC4 CryptoAPI

This is distinct from OOXML `EncryptedPackage` encryption.

Implementation: `crates/formula-xls/src/decrypt.rs` (used by
`formula_xls::import_xls_path_with_password`).

### Key derivation (RC4 CryptoAPI)

Terminology:

- `salt` is `EncryptionVerifier.salt` from the CryptoAPI `EncryptionInfo` embedded inside the BIFF
  `FILEPASS` record.
- `pw` is UTF-16LE(password) (no BOM, no terminator).

Algorithm (as implemented):

```text
H0 = SHA1(pw)
H  = SHA1(salt || H0)
for i in 0..50000:
  H = SHA1(LE32(i) || H)

// per-block RC4 key (keyLen is 5, 7, or 16 bytes depending on KeySizeBits)
K_block = SHA1(H || LE32(blockIndex))[0..keyLen]
```

### Payload decryption model (record-payload-only RC4 stream)

- The BIFF workbook stream is treated as a stream of **record payload bytes only** (record headers
  are not decrypted).
- RC4 keystream is applied across payload bytes, rekeying every **1024 bytes** of payload.
- After successful decryption, the `FILEPASS` record id is masked to `0xFFFF` (leaving record sizes
  intact) so downstream parsers that do not implement encryption can parse the stream without
  shifting offsets.

## Security notes

### Password handling
* Treat passwords as secrets: **never** log them and avoid storing them in long-lived structs.
* Convert to UTF-16LE only for KDF input; immediately zeroize intermediate buffers where feasible.

### Zeroization
We aim to `zeroize()`:
* password UTF-16LE buffers
* intermediate hash state output buffers (`H`, derived keys)
* decrypted package key material when dropped

Be aware that Rust `Vec` reallocation, copies, and allocator behavior can still duplicate sensitive
bytes. Zeroization is best-effort, not a formal guarantee.

### Caching / streaming decrypt
To support streaming readers (don’t materialize the entire decrypted ZIP in memory):

* Cache the **derived package key** (Agile) or **derived AES key** (Standard) after password
  verification.
* Do **not** cache the user password or its UTF-16LE encoding.
* Compute IVs per segment deterministically; caching IVs is unnecessary and increases secret
  surface area.

### Spin count DoS limits
Spin counts are attacker-controlled and can be set extremely high to cause CPU denial of service.
The reader should enforce a reasonable maximum spin count (and surface a specific error).

Note: Standard encryption uses a fixed 50,000 iteration count; the DoS concern is primarily for
Agile’s file-provided `spinCount`.

## Spec references (sections we implement)
Primary:
* **MS-OFFCRYPTO** (Office encryption container, Standard + Agile):
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
* **MS-CFB** (OLE/CFB container format used by the encrypted OOXML wrapper):
  https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
* **MS-XLS** (legacy BIFF / `FILEPASS` encryption marker):
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/

### MS-OFFCRYPTO sections (most relevant)
The MS-OFFCRYPTO spec is long; these are the sections we repeatedly refer to when implementing or
debugging encryption:

* **§2.3.4.4** — “`\\EncryptedPackage` Stream” (8-byte plaintext size prefix; ciphertext may be larger
  due to block padding; truncate to the declared size after decrypting).
* **§2.3.4.5** — “`\\EncryptionInfo` Stream (Standard Encryption)” (binary header layout:
  version+flags+headerSize, `EncryptionHeader`, `EncryptionVerifier`).
* **§2.3.4.7** — “ECMA-376 Document Encryption Key Generation (Standard Encryption)” (fixed 50,000
  password-hash iterations; salt is `EncryptionVerifier.Salt`).
* **§2.3.4.12** — “Initialization Vector Generation (Agile Encryption)” (segment index → IV hashing +
  truncation).
* **§2.3.4.15** — “Data Encryption (Agile Encryption)” (4096-byte segmenting and padding/truncation
  behavior; in practice the same segment framing pattern is used by CryptoAPI AES `EncryptedPackage`).

Other useful keywords inside MS-OFFCRYPTO:
* `spinCount` (Agile password hashing loop)
* `encryptedHmacKey` / `encryptedHmacValue` (“Data Integrity”; HMAC computed over the full
  `EncryptedPackage` stream bytes, including the size prefix)

Additional repo-specific references:
- `docs/offcrypto-standard-encryptedpackage.md` (Standard `EncryptedPackage` segmenting/IV/padding)
- `docs/21-encrypted-workbooks.md` (detection + UX/API semantics; fixture locations)
