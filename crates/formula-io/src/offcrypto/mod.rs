//! Helpers for parsing/decrypting Office binary encryption formats (MS-OFFCRYPTO).
//!
//! This module currently contains helpers for decrypting the `EncryptedPackage` stream used by
//! "Standard" (CryptoAPI) encryption inside an OLE compound file.
//!
//! Note: MS-OFFCRYPTO Standard encryption also has an **RC4** variant ("CryptoAPI RC4") whose
//! `EncryptedPackage` payload is decrypted in **0x200-byte** blocks with per-block keys derived from
//! the password + salt + a 50000-iteration SHA-1 spin loop. That variant is documented in
//! `docs/offcrypto-standard-cryptoapi-rc4.md`.

pub mod cryptoapi;
pub mod encrypted_package;
pub mod standard;

pub use encrypted_package::{
    decrypt_standard_cryptoapi_rc4_encrypted_package_stream,
    decrypt_standard_encrypted_package_stream, EncryptedPackageError,
};
pub use standard::{
    parse_encryption_info_standard, verify_password_standard, EncryptionHeader, EncryptionVerifier,
    OffcryptoError, StandardEncryptionInfo, CALG_AES_128, CALG_AES_192, CALG_AES_256, CALG_MD5,
    CALG_RC4, CALG_SHA1,
};
