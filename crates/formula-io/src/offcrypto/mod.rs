//! Helpers for parsing/decrypting Office binary encryption formats (MS-OFFCRYPTO).
//!
//! This module currently contains helpers for decrypting the `EncryptedPackage` stream used by
//! "Standard" (CryptoAPI) encryption inside an OLE compound file.

pub mod encrypted_package;

pub use encrypted_package::{decrypt_standard_encrypted_package_stream, EncryptedPackageError};

