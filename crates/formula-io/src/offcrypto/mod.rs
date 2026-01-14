//! Helpers for parsing/decrypting Office binary encryption formats (MS-OFFCRYPTO).
//!
//! This module currently contains helpers for decrypting the `EncryptedPackage` stream for
//! MS-OFFCRYPTO "Standard" (CryptoAPI) encryption inside an OLE compound file.
//!
//! Standard/CryptoAPI AES `EncryptedPackage` decryption (baseline MS-OFFCRYPTO/ECMA-376) uses
//! **AES-ECB** (no IV). The stream framing + truncation rules are documented in
//! `docs/offcrypto-standard-encryptedpackage.md`.
//!
//! Note: MS-OFFCRYPTO Standard encryption also has an **RC4** variant ("CryptoAPI RC4") whose
//! `EncryptedPackage` payload is decrypted in **0x200-byte** blocks with per-block keys derived from
//! the password + salt + a 50000-iteration `Hash()` spin loop (commonly SHA-1, sometimes MD5). That
//! variant is documented in
//! `docs/offcrypto-standard-cryptoapi-rc4.md`.
//!
//! Separately, we keep small self-contained parsers/verifiers here that are intended to be safe on
//! untrusted inputs (no panics and bounded allocations). Property-based tests for these helpers are
//! gated behind the `formula-io/offcrypto` feature so CI can run hardening without enabling all
//! encrypted workbook fixtures.

pub mod cryptoapi;
pub mod encrypted_package;
pub mod standard;

pub use encrypted_package::{
    decrypt_encrypted_package_standard_aes_to_writer,
    decrypt_standard_cryptoapi_rc4_encrypted_package_stream,
    decrypt_standard_cryptoapi_rc4_encrypted_package_stream_with_hash,
    decrypt_standard_encrypted_package_stream,
    EncryptedPackageDecryptError,
    EncryptedPackageError,
    EncryptedPackageToWriterError,
    InvalidCiphertextLenReason,
};
pub use standard::{
    parse_encryption_info_standard, verify_password_standard, EncryptionHeader, EncryptionVerifier,
    OffcryptoError, StandardEncryptionInfo, CALG_AES_128, CALG_AES_192, CALG_AES_256, CALG_MD5,
    CALG_RC4, CALG_SHA1,
};
