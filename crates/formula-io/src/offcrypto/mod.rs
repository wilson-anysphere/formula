//! Helpers for parsing/decrypting Office binary encryption formats (MS-OFFCRYPTO).
//!
//! This module currently contains helpers for decrypting the `EncryptedPackage` stream used by
//! "Standard" (CryptoAPI) encryption inside an OLE compound file.
//!
//! Note: MS-OFFCRYPTO Standard encryption also has an **RC4** variant ("CryptoAPI RC4") whose
//! `EncryptedPackage` payload is decrypted in **0x200-byte** blocks with per-block keys derived from
//! the password + salt + a 50000-iteration SHA-1 spin loop. That variant is documented in
//! `docs/offcrypto-standard-cryptoapi-rc4.md`.

pub mod encrypted_package;

pub use encrypted_package::{decrypt_standard_encrypted_package_stream, EncryptedPackageError};
