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
| OOXML (`.xlsx`/`.xlsm`/`.xlsb`) | **Agile** | `EncryptionInfo` **4.4** | ✅ decrypt (library) + ✅ encrypt (writer); ✅ open in `formula-io` behind `encrypted-workbooks` (Agile `.xlsx`/`.xlsm`/`.xlsb`) | `crates/formula-office-crypto` (end-to-end decrypt + OOXML encryption writer), `crates/formula-xlsx/src/offcrypto/*` (Agile decryptor + primitives), `crates/formula-offcrypto` (Agile parse + decrypt helpers; optional integrity verification) |
| OOXML (`.xlsx`/`.xlsm`/`.xlsb`) | **Standard / CryptoAPI (AES + RC4)** | `EncryptionInfo` `minor=2` (major ∈ {2,3,4} in the wild) | ✅ decrypt (library) + ✅ encrypt (writer; AES-only); ✅ open in `formula-io` behind `encrypted-workbooks` (`.xlsx`/`.xlsm`/`.xlsb`) | `crates/formula-office-crypto` (end-to-end decrypt + Standard AES writer), `crates/formula-offcrypto` (parse + Standard key derivation + verifier + `EncryptedPackage` decrypt for AES-ECB and RC4; stricter alg gating for Standard AES), `docs/offcrypto-standard-encryptedpackage.md` |
| Legacy `.xls` (BIFF5/BIFF8) | **FILEPASS** (XOR / RC4 Standard / RC4 CryptoAPI) | BIFF `FILEPASS` record | ✅ decrypt when password provided (import API) | `formula_xls::import_xls_path_with_password` / `import_xls_bytes_with_password`, `crates/formula-xls/src/decrypt.rs` |

Important: `formula-io`’s public open APIs **detect** encryption and surface dedicated errors so
callers can decide whether to prompt for a password vs report “unsupported encryption”.

- For Office-encrypted OOXML (`EncryptionInfo` + `EncryptedPackage`):
  - Without `formula-io/encrypted-workbooks`, encrypted OOXML containers surface
    `Error::UnsupportedEncryption` (and `Error::UnsupportedOoxmlEncryption` for unknown/invalid
    `EncryptionInfo` versions).
  - With `formula-io/encrypted-workbooks`, the password-aware open APIs can decrypt and open Agile
    (4.4) and Standard/CryptoAPI (minor=2; commonly `3.2`/`4.2`) encrypted `.xlsx`/`.xlsm`/`.xlsb`
    workbooks in memory.
    - For Agile, `dataIntegrity` (HMAC) is validated when present; some real-world producers omit it.
    - Decrypted packages containing `xl/workbook.bin` are routed to the XLSB open path.
  - `open_workbook_with_options` can also decrypt and open encrypted OOXML wrappers when a password
    is provided (returns `Workbook::Xlsx` / `Workbook::Xlsb` depending on the decrypted payload).
  - `open_workbook_model_with_options` can also decrypt encrypted OOXML wrappers when
    `formula-io/encrypted-workbooks` is enabled (and surfaces `PasswordRequired` when
    `OpenOptions.password` is `None`). Without that feature, encrypted OOXML containers surface
    `UnsupportedEncryption`. `open_workbook_model_with_password` is a convenience wrapper.
- For legacy `.xls` BIFF `FILEPASS`:
  - `open_workbook(..)` / `open_workbook_model(..)` surface `Error::EncryptedWorkbook`
  - `open_workbook_with_password(..)` / `open_workbook_model_with_password(..)` surface
    `PasswordRequired` / `InvalidPassword` (and attempt legacy `.xls` `FILEPASS` decryption when a
    password is provided: XOR / RC4 “standard” / RC4 CryptoAPI).

## Supported schemes / parameter subsets

### OOXML: Agile encryption (4.4)
We implement the password key encryptor subset of Agile (`keyEncryptor uri=".../password"`).

For a focused developer reference on Agile OOXML decryption (especially `dataIntegrity` HMAC target
bytes and IV/salt pitfalls), see [`docs/22-ooxml-encryption.md`](./22-ooxml-encryption.md).

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
Supported subset (CryptoAPI AES + RC4):

- Cipher:
  - AES-128/192/256 (`CALG_AES_128/192/256`).
  - RC4 (`CALG_RC4`). (Note: Standard/CryptoAPI RC4 uses 0x200-byte blocks; see
    `docs/offcrypto-standard-cryptoapi-rc4.md`.)
- Hash (`AlgIDHash`) for password hashing / key derivation:
  - Standard RC4 (`CALG_RC4`): `CALG_SHA1` and `CALG_MD5` (both occur in the wild).
  - Standard AES (`CALG_AES_*`): Excel-default is `CALG_SHA1`. In this repo:
    - `crates/formula-offcrypto` gates to **AES + SHA-1** only.
    - `crates/formula-office-crypto` also supports `CALG_SHA_256`/`CALG_SHA_384`/`CALG_SHA_512` for compatibility.
- Salt size: file-provided (`EncryptionVerifier.saltSize`, typically 16).

Other Standard combinations (mismatched key sizes, non-CryptoAPI flags, etc.) are treated as
unsupported by the current code.

Note on version gating in helper APIs:

- Most parsers in this repo treat Standard as **`versionMinor == 2`** with `versionMajor ∈ {2,3,4}`.
- Some convenience APIs are intentionally scoped:
  - `formula_offcrypto::decrypt_standard_ooxml_from_bytes` performs Standard-only decryption and will
    reject non-Standard inputs (Agile / Extensible) with
    `OffcryptoError::UnsupportedEncryption { encryption_type: ... }` before attempting any password
    verification.
- For best compatibility across Standard variants (non-default hashes/algorithms), prefer
  `crates/formula-office-crypto`’s decryptor.

### Legacy `.xls`: BIFF `FILEPASS` (BIFF5/BIFF8)
Currently supported in `formula-xls` (see also the fixture inventory in
[`crates/formula-xls/tests/fixtures/encrypted/README.md`](../crates/formula-xls/tests/fixtures/encrypted/README.md)):

- **BIFF8 XOR obfuscation** (`wEncryptionType=0x0000`)
- **BIFF8 RC4 “standard”** (`wEncryptionType=0x0001`, `wEncryptionSubType=0x0001`)
  - Password is truncated to the first **15 UTF-16 code units** (Excel legacy behavior).
- **BIFF8 RC4 CryptoAPI** (`wEncryptionType=0x0001`, `wEncryptionSubType=0x0002`)
  - Supports both CryptoAPI FILEPASS payload layouts we see in the wild (see
    [Legacy `.xls` `FILEPASS` details](#legacy-xls-filepass-encryption-biff5biff8)):
    - `wEncryptionSubType==0x0002` with a length-prefixed embedded `EncryptionInfo`
    - legacy `wEncryptionInfo==0x0004` with an embedded `EncryptionHeader` + `EncryptionVerifier`
  - `AlgIDHash` is typically `CALG_SHA1`, but some producers use `CALG_MD5`.
  - Fixed 50,000 password-hash iterations (CryptoAPI loop; **not PBKDF2**) for the `subType=0x0002`
    layout.
  - `KeySizeBits` values: `0/40`, `56`, `128`
    - `KeySizeBits==0` is treated as 40-bit RC4.
    - Key bytes are the first `KeySizeBits/8` bytes of the per-block digest (5/7/16 bytes; no 40-bit
      padding-to-16 quirk for BIFF CryptoAPI).
- (best-effort) **BIFF5-era XOR obfuscation** (Excel 5/95)

Not implemented:

- BIFF CryptoAPI with `AlgID != CALG_RC4` (including AES CryptoAPI-for-BIFF variants).
- BIFF CryptoAPI with `AlgIDHash` not in {`CALG_SHA1`, `CALG_MD5`} or other uncommon verifier/hash variants.
- Other uncommon BIFF encryption variants.

## Container format details (what’s inside the OLE file)

### `EncryptionInfo` (common header)
Both Standard and Agile `EncryptionInfo` start with:

```text
u16 versionMajor
u16 versionMinor
u32 flags
[Standard only] u32 headerSize
... scheme-specific payload (Standard: binary header/verifier; Agile: XML descriptor)
```

The version pair is used to dispatch:
* **minor == 2** and major ∈ {2,3,4} ⇒ Standard (CryptoAPI)
* **4.4** ⇒ Agile

### `EncryptedPackage` stream framing
`EncryptedPackage` begins with an **8-byte plaintext size prefix** (little-endian), followed by the
ciphertext bytes:

```text
u32le original_size_lo
u32le original_size_hi_or_reserved
u8    encrypted_bytes[...]
```

Compatibility note: while MS-OFFCRYPTO describes the prefix as a `u64le`, some producers/libraries
treat it as `u32 totalSize` + `u32 reserved` (often 0). To avoid truncation or “huge size” misreads,
parse as `lo=u32le(bytes[0..4])`, `hi=u32le(bytes[4..8])`, then
`original_size = lo as u64 | ((hi as u64) << 32)`.

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
key = TruncateHash(derived, keyBits/8)
```

Notes:
* `Hash` is the `hashAlgorithm` declared by the relevant XML node (`SHA1`/`SHA256`/`SHA384`/`SHA512`).
* The iteration counter `i` is a **little-endian u32**.
* `TruncateHash` truncates when `keyBits/8 <= hashLen`, otherwise it expands by padding with `0x36`
  bytes to reach the requested length (matches Excel + `msoffcrypto-tool`).

#### Spin count DoS limits

`spinCount` is **attacker-controlled input**. Extremely large values can cause CPU DoS because the
iterated hash loop runs `spinCount` digest operations.

Current state in this repo:

- `crates/formula-offcrypto` enforces a configurable maximum via `DecryptOptions` /
  `DecryptLimits.max_spin_count` (default: `DEFAULT_MAX_SPIN_COUNT = 1_000_000`) and returns
  `OffcryptoError::SpinCountTooLarge { spin_count, max }` when exceeded.
- `crates/formula-xlsx::offcrypto` enforces a configurable maximum via
  `offcrypto::DecryptOptions.max_spin_count` (default: `DEFAULT_MAX_SPIN_COUNT = 1_000_000`) and
  returns `OffCryptoError::SpinCountTooLarge { spin_count, max }` when exceeded.
- `crates/formula-office-crypto` enforces a configurable maximum via
  `DecryptOptions.max_spin_count` (default: `DEFAULT_MAX_SPIN_COUNT = 1_000_000`) and returns
  `OfficeCryptoError::SpinCountTooLarge { spin_count, max }` when exceeded.

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

IV nuance (common interop footgun):

- For these password-key-encryptor blobs (and `encryptedKeyValue`), the AES-CBC IV is simply the
  password key encryptor `saltValue` truncated to `blockSize` (typically 16).
- The fixed `blockKey` constants above are used for **key derivation**, not IV derivation.

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
  - Note: MS-OFFCRYPTO/Excel authenticates the **entire `EncryptedPackage` stream bytes** (8-byte
    length prefix + ciphertext). For compatibility with some non-Excel producers, our decryptors may
    also accept alternative HMAC targets (e.g. plaintext ZIP bytes, ciphertext-only, or header +
    plaintext). For the exact per-crate acceptance behavior, see
    [`docs/22-ooxml-encryption.md`](./22-ooxml-encryption.md).

Implementation status:

- `crates/formula-xlsx::offcrypto` validates `dataIntegrity` and returns `IntegrityMismatch` on
  failure.
- `crates/formula-office-crypto` validates `dataIntegrity` and returns:
  - `IntegrityCheckFailed` on HMAC mismatch
  - When the `<dataIntegrity>` element is missing, decryption can still succeed but **no integrity
    verification** is performed (decrypted bytes are unauthenticated).
- `formula-io` (behind the `encrypted-workbooks` feature):
  - The legacy `_with_password` APIs use the `formula-xlsx` Agile decryptor and validate
    `dataIntegrity` when it is present.
  - `open_workbook_with_options` also uses the `formula-xlsx` Agile decryptor and validates
    `dataIntegrity` when it is present (when `<dataIntegrity>` is missing, decryption proceeds but no
    integrity verification is performed).
  - The streaming decrypt reader (`crates/formula-io/src/encrypted_ooxml.rs`) does not validate
    `dataIntegrity`.
    - It is used for some compatibility fallbacks (for example Agile files that omit
      `<dataIntegrity>`) and to open Standard/CryptoAPI AES-encrypted `.xlsx`/`.xlsm` into a model
      without materializing the decrypted ZIP bytes (`open_workbook_model_with_password` /
      `open_workbook_model_with_options`).
    - Other encrypted-open paths (including `open_workbook_with_password` / `open_workbook_with_options`)
      decrypt `EncryptedPackage` into an in-memory buffer first.
- `crates/formula-offcrypto` can validate `dataIntegrity` when decrypting Agile packages via
  `decrypt_encrypted_package` with `DecryptOptions.verify_integrity = true` (default: `false`).
  - It verifies only the spec/Excel target (the full `EncryptedPackage` stream bytes) and returns
    `OffcryptoError::IntegrityCheckFailed` on mismatch.
  - Other helper APIs (e.g. `decrypt_agile_ooxml_from_bytes`) currently do not perform integrity
    verification.
- HMAC verification is strongly recommended when possible to distinguish wrong passwords from “ZIP
  happened to parse”.

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

Implementation note: in this repo (matching `crates/formula-offcrypto`), Standard/CryptoAPI **AES**
uses **AES-ECB** (no IV) for both the verifier fields **and** `EncryptedPackage` for
Excel-default/ECMA-376 Standard AES. Some third-party producers use a non-standard segmented
AES-CBC variant with a per-segment IV derived from the verifier salt; `formula-io` includes
compatibility fallbacks for some of these cases (see `docs/offcrypto-standard-encryptedpackage.md`).

### Standard (CryptoAPI): `EncryptedPackage` decryption (AES-ECB, no IV)

Standard/CryptoAPI AES `EncryptedPackage` decryption is **AES-ECB** (no IV) over the ciphertext
bytes (after the 8-byte `orig_size` prefix):

```text
plaintext = AES-ECB-Decrypt(key, ciphertext)
plaintext = plaintext[0..orig_size]   // truncate to the declared size
```

Where:

- `orig_size` is the 8-byte little-endian size prefix at the start of the `EncryptedPackage` stream.
- `ciphertext` is the remaining bytes after the prefix and must be a multiple of 16 bytes.
- Do **not** “unpad” decrypted bytes; always truncate to `orig_size` (padding/trailing bytes are
  not reliable PKCS#7).

See `docs/offcrypto-standard-encryptedpackage.md` for framing/padding/truncation edge cases.

Implementation nuance:

- `crates/formula-offcrypto/src/encrypted_package.rs` implements Standard AES `EncryptedPackage`
  decryption via AES-ECB and truncation to `orig_size`.
- `crates/formula-office-crypto` implements end-to-end Standard decryption and is intentionally more
  permissive about Standard parameter variants for compatibility.

### Standard (CryptoAPI): RC4 `EncryptedPackage` decryption (0x200 blocks, CryptoAPI 40-bit padding)

MS-OFFCRYPTO Standard encryption can also use `CALG_RC4` (instead of AES).

In this repo, `formula-io` (behind the `encrypted-workbooks` feature) can decrypt and open
Standard/CryptoAPI RC4-encrypted `.xlsx`/`.xlsm`/`.xlsb` workbooks end-to-end. (Encrypted `.xlsb`
payloads decrypt to an OOXML ZIP containing `xl/workbook.bin` and are routed through the `.xlsb`
reader (`formula-xlsb`) in both native `formula-io` and the WASM loader (`formula-wasm`). The streaming RC4
decryptor lives in `crates/formula-io/src/rc4_cryptoapi.rs` and is exercised by
`crates/formula-io/tests/offcrypto_standard_rc4_vectors.rs`.

Critical nuance: for **40-bit** RC4 (`KeySize == 0` or `KeySize == 40`), CryptoAPI represents the
key as a **16-byte**
RC4 key where the low 40 bits come from the derived key material and the remaining 88 bits are zero.

High-level shape (SHA-1; per MS-OFFCRYPTO Standard RC4):

```text
pw = UTF16LE(password)                      // no BOM, no NUL
H  = SHA1(salt || pw)
for i in 0..50000:
  H = SHA1(LE32(i) || H)

for blockIndex = 0, 1, 2, ...:              // 0x200-byte blocks
  Hb = SHA1(H || LE32(blockIndex))
  keySizeBits = KeySize
  if keySizeBits == 0:
    keySizeBits = 40                        // MS-OFFCRYPTO: RC4 KeySize=0 means 40-bit
  key_material = Hb[0..keySizeBits/8]
  if keySizeBits == 40:
    rc4_key = key_material || 0x00 * 11     // 16 bytes total (CryptoAPI quirk)
  else:
    rc4_key = key_material                  // 7 bytes (56-bit) or 16 bytes (128-bit)
  plaintext_block = RC4(rc4_key, ciphertext_block)
```

See `docs/offcrypto-standard-cryptoapi-rc4.md` for a full writeup and an example that shows raw
5-byte vs padded-16-byte keys produce different ciphertext.

## Interop notes / fixture generation

### Which Excel versions produce which scheme?
In practice:

* **Excel 2007** typically produces **Standard (3.2)** encryption.
* **Excel 2010+** typically produces **Agile (4.4)** encryption.

Excel can also vary algorithms via “Advanced” encryption options; our supported subset assumes the
defaults above.

### Generating fixtures (recommended)
Repo note: most committed encrypted OOXML fixtures under `fixtures/encrypted/ooxml/` are generated
using Python + [`msoffcrypto-tool`](https://github.com/nolze/msoffcrypto-tool), but some fixtures are
generated via Apache POI (for example `standard-4.2.xlsx`, `standard-unicode.xlsx`). See
`fixtures/encrypted/ooxml/README.md` for the canonical per-fixture provenance + exact tool
versions/passwords.

To generate fixtures for regression tests:

1. Create a minimal workbook (one sheet, a few cells).
2. In Excel: **File → Info → Protect Workbook → Encrypt with Password**.
3. Save as `.xlsx` (Agile by default on modern Excel).

For Standard fixtures, prefer deterministic tooling (like the repo’s `msoffcrypto-tool`-generated
`fixtures/encrypted/ooxml/standard.xlsx`) or a known Excel 2007 install (if you specifically need
Excel-ground-truth output).

Repo fixtures:

- OOXML encrypted fixtures live in `fixtures/encrypted/ooxml/` (see that directory’s README).
- BIFF `.xls` encrypted fixtures for `FILEPASS` live under `crates/formula-xls/tests/fixtures/encrypted/`
  (see [`crates/formula-xls/tests/fixtures/encrypted/README.md`](../crates/formula-xls/tests/fixtures/encrypted/README.md)).

Note: the Apache POI-based generator under `tools/encrypted-ooxml-fixtures/` can emit Standard
encryption with an `EncryptionInfo` version of `4.2` (still Standard/CryptoAPI; `versionMinor == 2`).

### Inspecting the scheme quickly
To identify the scheme without full parsing:

* Open the file as OLE/CFB and read the first 8 bytes of `EncryptionInfo`.
* Interpret as `u16 major, u16 minor, u32 flags` (little-endian).
  * 4.4 ⇒ Agile
  * `minor == 2` and `major ∈ {2,3,4}` ⇒ Standard (CryptoAPI; commonly 3.2)

### Defaults for *writing* encrypted OOXML (`formula-office-crypto`)
`formula-io` does not automatically preserve Office encryption when you call `save_workbook`: it will
write a plaintext `.xlsx`/`.xlsm` unless you explicitly opt into re-encryption.

With the `formula-io/encrypted-workbooks` feature enabled, callers that want to preserve Office
encryption for an encrypted OOXML OLE/CFB container can use:

- `formula_io::open_workbook_with_password_and_preserved_ole(..)` to capture non-encryption OLE
  streams/storages (e.g. `\u{0005}SummaryInformation`), and
- `OpenedWorkbookWithPreservedOle::save_preserving_encryption(..)` to re-encrypt the workbook **in
  memory** and write a new OLE wrapper (with fresh `EncryptionInfo` + `EncryptedPackage` streams).

The underlying writer lives in `crates/formula-office-crypto`
(`formula_office_crypto::encrypt_package_to_ole`) and is used for round-trip tests. Its defaults are
the repo’s current “recommended writer settings” unless a compatibility requirement dictates
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

- **Standard (supported; `scheme=Standard`):**
  - CryptoAPI **AES** only (key sizes: 128/192/256 via `key_bits`)
  - `hash_algorithm=SHA1` and `spin_count=50_000` (Standard/CryptoAPI constraints; enforced by the writer)
  - `EncryptionInfo` header version: `3.2` (commonly accepted by Office tooling)
  - `EncryptedPackage` encryption: **AES-ECB** (no IV). Ciphertext is block-aligned (pad plaintext to
    16-byte blocks when encrypting), and plaintext is truncated to the 8-byte `orig_size` prefix
    after decrypting (see `docs/offcrypto-standard-encryptedpackage.md`).
  - Note: Standard/CryptoAPI **RC4** writing is not currently supported by this writer.

These defaults are intended for **interoperability** with Excel, not for novel cryptographic
design.

## Legacy `.xls` `FILEPASS` encryption (BIFF5/BIFF8)

This is distinct from OOXML `EncryptedPackage` encryption.

Implementation entry points:

- `formula_xls::import_xls_path_with_password` / `formula_xls::import_xls_bytes_with_password`
- Under the hood:
  - XOR + RC4 “standard”: `crates/formula-xls/src/biff/encryption.rs`
  - RC4 CryptoAPI: `crates/formula-xls/src/decrypt.rs`

Fixtures / test corpus:
[`crates/formula-xls/tests/fixtures/encrypted/README.md`](../crates/formula-xls/tests/fixtures/encrypted/README.md).

### Password semantics (Excel legacy)

- **XOR obfuscation:** passwords are effectively limited to **15 bytes**; extra characters are
  ignored.
  - Excel uses a legacy password-to-bytes mapping for XOR. Our implementation is best-effort for
    Unicode passwords by trying both MS-OFFCRYPTO “method 1” and “method 2” byte derivations.
- **RC4 “standard” truncation:** only the first **15 UTF-16 code units** are significant; extra
  characters are ignored.
- **RC4 CryptoAPI:** uses the full password string (UTF-16LE, no 15-character truncation).
- **Empty passwords:** supported when the workbook was encrypted that way (pass `""`).

### Decryption model (BIFF record stream)

The BIFF workbook stream is a sequence of records:

```text
u16 record_id
u16 record_len
u8  record_payload[record_len]
```

Encryption begins immediately after the `FILEPASS` record and continues until the end of the
workbook stream:

- Record headers (`record_id` + `record_len`) are always plaintext.
- Record payload bytes after `FILEPASS` are encrypted/obfuscated (with a few scheme-specific
  exceptions for legacy XOR/CryptoAPI variants).
- After successful decryption, we **mask the `FILEPASS` record id** to `0xFFFF` (leaving record sizes
  and payload bytes intact). This allows downstream BIFF parsers (and `calamine`) that do not
  implement encryption to treat the stream as plaintext without shifting offsets (notably
  `BoundSheet8.lbPlyPos`).
- For XOR obfuscation (and the legacy CryptoAPI layout below), some records/fields remain plaintext
  even after `FILEPASS` (notably `BOF`/`FILEPASS`/`INTERFACEHDR`, and `BoundSheet8.lbPlyPos`).

### CryptoAPI (`wEncryptionType=0x0001`) FILEPASS payload layouts

RC4 CryptoAPI (`wEncryptionSubType=0x0002`) appears in two payload layouts in the wild; we support
both.

#### Layout A: `wEncryptionSubType == 0x0002` (length-prefixed `EncryptionInfo`)

FILEPASS payload layout:

```text
u16 wEncryptionType    = 0x0001
u16 wEncryptionSubType = 0x0002
u32 dwEncryptionInfoLen
u8  encryptionInfo[dwEncryptionInfoLen]  // MS-OFFCRYPTO EncryptionInfo (CryptoAPI)
```

Key derivation (as implemented; fixed 50,000 iterations; `Hash` is `SHA1` or `MD5` based on
`EncryptionHeader.AlgIDHash`):

```text
pw = UTF16LE(password)
H  = Hash(salt || pw)
for i in 0..50000:
  H = Hash(LE32(i) || H)

// KeySizeBits can be 0 (meaning 40-bit), 40, 56, or 128.
H_block = Hash(H || LE32(blockIndex))
K_block = H_block[0..keyLen]             // keyLen = (KeySizeBits==0 ? 40 : KeySizeBits)/8
```

Decrypt model:

- Treat the encrypted byte stream as **record payload bytes only** (record headers are not
  decrypted and do not advance the RC4 position).
- Rekey every **1024 bytes** of payload.

#### Layout B: legacy `wEncryptionInfo == 0x0004` (embedded header/verifier)

Some RC4 CryptoAPI workbooks use an older FILEPASS layout where the CryptoAPI header/verifier are
embedded directly in the FILEPASS payload (no length-prefixed `EncryptionInfo` blob):

```text
u16 wEncryptionType = 0x0001
u16 wEncryptionInfo = 0x0004
u16 vMajor
u16 vMinor
u16 reserved        = 0
u32 headerSize
u8  encryptionHeader[headerSize]
u8  encryptionVerifier[...]
```

Nuances (as implemented; mirrors Excel/Apache POI behavior):

- **Different key-material derivation:** the legacy layout derives key material as
  `Hash(salt || UTF16LE(password))` (no 50,000-iteration spin loop).
  - Some legacy headers omit/zero `AlgIDHash`; we treat `AlgIDHash==0` as SHA-1 for compatibility.
- **Different stream-position semantics:** the RC4 block index + in-block position are derived from
  the **absolute workbook-stream offset**. Record headers are not encrypted, but they still advance
  the RC4 “encryption stream position”.
- A small set of records are **never encrypted** (notably `BOF`, `FILEPASS`, `INTERFACEHDR`), but
  their plaintext bytes still advance the RC4 position.
- `BoundSheet8.lbPlyPos` (the first 4 bytes of the `BOUNDSHEET` payload) remains plaintext so sheet
  offsets remain valid after masking `FILEPASS`.

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
* For **Agile**, compute per-segment IVs deterministically; caching IVs is unnecessary and increases
  secret surface area. (Standard AES-ECB uses no IV.)

### Spin count DoS limits
Spin counts are attacker-controlled and can be set extremely high to cause CPU denial of service.
The reader should enforce a reasonable maximum spin count (and surface a specific error).

Note: Standard encryption uses a fixed 50,000 iteration count; the DoS concern is primarily for
Agile’s file-provided `spinCount`.

### `EncryptionInfo` size limits (XML + base64 fields)
Agile `EncryptionInfo` embeds an XML descriptor with multiple base64-encoded fields
(`saltValue`, `encryptedKeyValue`, `encryptedVerifierHash*`, `encryptedHmac*`, …).
These fields are attacker-controlled and can be made extremely large to cause **memory** DoS.

In this repo:

- `crates/formula-xlsx::offcrypto` enforces bounded parsing via `ParseOptions` (defaults: **1 MiB**
  max XML length / base64 field length / decoded length), and returns structured errors:
  `OffCryptoError::EncryptionInfoTooLarge` / `OffCryptoError::FieldTooLarge`.
  - See: `crates/formula-xlsx/src/offcrypto/encryption_info.rs`

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
  behavior).

Other useful keywords inside MS-OFFCRYPTO:
* `spinCount` (Agile password hashing loop)
* `encryptedHmacKey` / `encryptedHmacValue` (“Data Integrity”; HMAC computed over the full
  `EncryptedPackage` stream bytes, including the size prefix)

Additional repo-specific references:
- `docs/offcrypto-standard-encryptedpackage.md` (Standard `EncryptedPackage` framing + AES-ECB decrypt + truncation)
- `docs/offcrypto-standard-cryptoapi.md` (Standard key derivation + verifier validation, from-scratch)
- `docs/offcrypto-standard-cryptoapi-rc4.md` (Standard CryptoAPI RC4 reference + test vectors)
- `docs/21-encrypted-workbooks.md` (detection + UX/API semantics; fixture locations)
