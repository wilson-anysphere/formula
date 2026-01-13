# Office workbook encryption (MS-OFFCRYPTO / MS-XLS) — implementation reference
This note is **maintainer-facing**. It captures the subset of Office encryption we support for Excel
workbooks, the *exact* encryption parameters we emit when writing, and the non-obvious key
derivation / verification details that tend to cause interop bugs.

This document is intentionally **not** user-facing (“how do I open a password-protected file?”).

## Terminology / file shapes

### OOXML open-password encryption (`.xlsx`/`.xlsm`/`.xlsb`)
Excel “Encrypt with Password” (open password) does **not** encrypt the ZIP container directly.
Instead, Excel writes an **OLE/CFB compound file** containing (at minimum):

* `EncryptionInfo` — encryption metadata (either binary “Standard” or XML “Agile”).
* `EncryptedPackage` — encrypted bytes of the *original* OPC/ZIP package.

The decrypted payload of `EncryptedPackage` is the raw `.xlsx`/`.xlsm`/`.xlsb` ZIP bytes.

### Legacy `.xls` BIFF encryption
Legacy BIFF `.xls` workbooks are signaled via a `FILEPASS` record in the workbook globals stream
([MS-XLS]). This is a *different* encryption stack than OOXML open-password encryption.

## Supported schemes (reader + writer)

### Summary table
| Scheme | EncryptionInfo version | Reader | Writer | Notes |
|---|---:|:---:|:---:|---|
| **Agile encryption** | 4.4 | ✅ | ✅ (default) | Modern Office (2010+). `EncryptionInfo` contains an XML descriptor. |
| **Standard encryption (CryptoAPI)** | 3.2 | ✅ | ✅ (opt-in) | Office 2007-era. `EncryptionInfo` contains `EncryptionHeader` + `EncryptionVerifier`. |
| Certificate-based key encryptors | n/a | ❌ | ❌ | We only implement **password** key encryption (`keyEncryptor uri=".../password"`). |
| IRM / “DataSpaces” transforms | n/a | ❌ | ❌ | We ignore `DataSpaces/*` and rely on `EncryptionInfo` + `EncryptedPackage`. |
| Legacy `.xls` FILEPASS decryption | n/a | ❌ | ❌ | We detect and error (do not attempt to decrypt BIFF streams). |

### Agile encryption: supported parameter subset
We accept (and when writing, emit) the following Agile subset:

* `cipherAlgorithm`: `AES`
* `cipherChaining`: `ChainingModeCBC`
* `hashAlgorithm`: `SHA512`
* `keyBits`: `256`
* `blockSize`: `16`
* `saltSize`: `16`
* `spinCount`: any `u32` **up to** our configured maximum (see [Security notes](#security-notes))

If a document uses other algorithms (e.g. SHA-1, AES-128, CFB, etc) we currently treat it as
unsupported rather than attempting partial decryption.

### Standard encryption (CryptoAPI): supported parameter subset
We accept (and when writing, emit) the following Standard subset:

* `AlgID`: AES-128 (`CALG_AES_128`)
* `AlgIDHash`: SHA-1 (`CALG_SHA1`)
* `KeySize`: `128`
* `ProviderType`: `PROV_RSA_AES`

Other Standard combinations (RC4, AES-256, non-SHA1 hashes) are treated as unsupported.

## Writer defaults (encryption we emit)

Unless explicitly overridden, the writer produces **Agile encryption** compatible with modern Excel
(tested against Excel 2016/2019/365).

### Defaults (Agile writer)
* Scheme: **Agile**
* Cipher: **AES-256-CBC**
* Hash: **SHA-512**
* `spinCount`: **100_000**
* `saltSize`: **16 bytes**
* `blockSize`: **16 bytes**
* Package segment size: **4096 bytes** (`0x1000`)

Randomness:
* `keyData.saltValue`: 16 bytes, generated from a CSPRNG
* `passwordKeyEncryptor.saltValue`: 16 bytes, generated independently from a CSPRNG
* `passwordKeyEncryptor.verifierHashInput`: 16 bytes, generated from a CSPRNG
* `dataIntegrity.hmacKey`: 64 bytes, generated from a CSPRNG (HMAC-SHA512 key)

### Defaults (Standard writer)
Standard encryption is available only for compatibility testing / fixture generation. Defaults:

* Scheme: **Standard (CryptoAPI)**
* Cipher: **AES-128-CBC**
* Hash: **SHA-1**
* `spinCount`: **50_000**
* `saltSize`: **16 bytes**
* Segment size: **512 bytes** (`0x200`)

## Container format details (what’s inside the OLE file)

### `EncryptionInfo` (common header)
Both Standard and Agile `EncryptionInfo` start with:

```text
u16 versionMinor
u16 versionMajor
u32 flags
... scheme-specific payload
```

The version pair is used to dispatch:
* **3.2** ⇒ Standard (CryptoAPI)
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
* `Hash` is the `hashAlgorithm` declared by the relevant XML node (SHA-512 in our supported subset).
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
Agile includes a `dataIntegrity` block that authenticates the plaintext package:

* `encryptedHmacKey` is decrypted using a key derived from `package_key` and `HMAC_KEY_BLOCK`
* `encryptedHmacValue` is decrypted using a key derived from `package_key` and `HMAC_VALUE_BLOCK`
* `HMAC-SHA512(hmacKey, plaintextPackageBytes)` must match `hmacValue` (constant-time compare)

### Standard (CryptoAPI): password KDF
Standard uses the same “spinCount loop” shape but always SHA-1 in our supported subset:

```text
pw = UTF16LE(password)
H0 = SHA1(salt + pw)
H  = H0
for i in 0..spinCount:
  H = SHA1(LE32(i) + H)

key = SHA1(H + LE32(0))[0..keySize/8]
```

The `LE32(0)` “block key” constant is fixed for Standard encryption key derivation.

#### Standard password verification
`EncryptionVerifier` contains:
* `salt` (16 bytes)
* `encryptedVerifier` (16 bytes)
* `encryptedVerifierHash` (20 bytes for SHA-1)

The verifier bytes are decrypted with AES-CBC using IV=`salt` (per spec). The decrypted verifier is
then hashed with SHA-1 and compared to the decrypted verifier hash.

#### Standard package decryption (segment IV derivation)
Standard encrypts `EncryptedPackage` in 512-byte segments. For segment index `i`:

```text
iv = SHA1(salt + LE32(i))[0..16]
plaintext_i = AES-CBC-Decrypt(key=key, iv=iv, ciphertext=segment_i)
```

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

### Inspecting the scheme quickly
To identify the scheme without full parsing:

* Open the file as OLE/CFB and read the first 8 bytes of `EncryptionInfo`.
* Interpret as `u16 minor, u16 major, u32 flags` (little-endian).
  * 3.2 ⇒ Standard
  * 4.4 ⇒ Agile

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

## Spec references (sections we implement)
Primary:
* **MS-OFFCRYPTO** (Office encryption container, Standard + Agile):
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
* **MS-XLS** (legacy BIFF / `FILEPASS` encryption marker):
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/

Useful entry points / keywords inside MS-OFFCRYPTO:
* `EncryptionInfo` stream format (version dispatch, flags)
* “Standard Encryption” (`EncryptionHeader`, `EncryptionVerifier`)
* “Agile Encryption” (XML descriptor, password key encryptor, `EncryptedPackage` segmenting)
* “Data Integrity” (`encryptedHmacKey`, `encryptedHmacValue`)

