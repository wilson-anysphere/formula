# Encrypted OOXML test fixtures

## `offcrypto_standard_cryptoapi_password.xlsx`

An OOXML workbook encrypted using **MS-OFFCRYPTO Standard / CryptoAPI** encryption (Office 2007-style).

- Container format: **OLE Compound File** (`EncryptionInfo` + `EncryptedPackage` streams)
- Password: `password`
- EncryptionInfo version: **major=3 minor=2**
- Cipher: **AES-128** (`algId = 0x660E`, `algIdHash = 0x8004` / SHA-1)

### How it was produced

This fixture is committed as a binary blob so CI does not need to generate it.

1. Create a minimal XLSX with Apache POI and encrypt it with `EncryptionMode.standard`.
2. Patch the first 2 bytes of the `EncryptionInfo` stream from `0x0004` → `0x0003`
   (POI writes `4.2`, but MS-OFFCRYPTO also defines the Standard encryption header layout for `3.2`).

The resulting file is still a valid encrypted OOXML container and exercises our parser/password verifier.

