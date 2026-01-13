//! OOXML encrypted package support (`[MS-OFFCRYPTO]`).
//!
//! Excel's **Encrypt with Password** feature produces files that *look like* `.xlsx`, but they are
//! no longer ZIP-based Open Packaging Convention (OPC) archives. Instead, Office wraps the real
//! workbook package in an **OLE Compound File Binary** (CFB) container with (at least) two streams:
//!
//! - `EncryptionInfo`: encryption parameters describing the algorithms, salts, spin count, key
//!   encryptor, and integrity metadata (for modern files this is the *Agile Encryption* XML
//!   descriptor).
//! - `EncryptedPackage`: the encrypted bytes of the original ZIP/OPC package.
//!
//! This module focuses on the `[MS-OFFCRYPTO]` **Agile Encryption** scheme (the default in modern
//! Office).
//!
//! In addition to higher-level parsing/decryption routines, this module also exposes small,
//! reusable crypto primitives (password hashing, key derivation, IV derivation) used by Agile
//! encryption.
//!
//! ## Spec references
//! Microsoft Open Specifications:
//! - `[MS-OFFCRYPTO] Office Document Cryptography Structure`
//! - *Agile Encryption*
//! - *EncryptionInfo Stream*
//! - *EncryptedPackage Stream*
//!
//! ## EncryptedPackage layout
//! After decryption, the `EncryptedPackage` stream yields:
//!
//! ```text
//! 8B   original_package_size (u64 little-endian)
//! ...  ZIP bytes (length = original_package_size)
//! ...  padding (to the cipher block size)
//! ```
//!
//! The ciphertext is processed in **4096-byte segments** (plaintext segment size). Each segment is
//! encrypted independently using the package key and a per-segment IV.
//!
//! ### IV derivation (Agile)
//! For segment `i` (0-based), the IV is derived from the `keyData/@saltValue` and the segment index:
//!
//! ```text
//! iv_i = Truncate(keyData/@blockSize, Hash(keyData/@saltValue || u32le(i)))
//! ```
//!
//! where `Hash` is the algorithm indicated by `keyData/@hashAlgorithm` and `Truncate` keeps the
//! first `blockSize` bytes of the hash output.
//!
//! ## Errors
//! Decryption failures are surfaced via [`OffCryptoError`], which aims to make failures actionable
//! (wrong password vs unsupported algorithms vs file corruption) without leaking sensitive inputs
//! such as passwords or derived keys.

mod agile;
mod aes_cbc;
mod agile_decrypt;
mod crypto;
mod encryption_info;
mod error;

pub use agile::{
    decrypt_agile_encrypted_package_stream, decrypt_agile_encrypted_package_stream_with_key,
    decrypt_agile_keys, parse_agile_encryption_info_stream, AgileDataIntegrity, AgileDecryptedKeys,
    AgileEncryptionInfo, AgileKeyData, AgilePasswordKeyEncryptor,
};
pub use aes_cbc::{
    decrypt_aes_cbc_no_padding, decrypt_aes_cbc_no_padding_in_place, AesCbcDecryptError,
    AES_BLOCK_SIZE,
};
pub use agile_decrypt::decrypt_agile_encrypted_package;
pub use crypto::{
    derive_iv, derive_key, hash_password, CryptoError, HashAlgorithm, HMAC_KEY_BLOCK,
    HMAC_VALUE_BLOCK, KEY_VALUE_BLOCK, VERIFIER_HASH_INPUT_BLOCK, VERIFIER_HASH_VALUE_BLOCK,
    derive_segment_iv, segment_block_key,
};
pub use encryption_info::{
    parse_agile_encryption_info_xml, AgileEncryptionInfo, EncryptionInfoWarning, PasswordKeyEncryptor,
    KEY_ENCRYPTOR_URI_CERTIFICATE, KEY_ENCRYPTOR_URI_PASSWORD,
};
pub use error::{OffCryptoError, Result};

