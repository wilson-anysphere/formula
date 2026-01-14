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

mod aes_cbc;
mod agile;
mod agile_decrypt;
mod crypto;
mod encryption_info;
mod error;
mod ooxml;
mod rc4;

#[allow(unused_imports)]
pub use aes_cbc::{
    decrypt_aes_cbc_no_padding, decrypt_aes_cbc_no_padding_in_place, AesCbcDecryptError,
    AES_BLOCK_SIZE,
};
pub use agile::{
    decrypt_agile_encrypted_package_stream, decrypt_agile_encrypted_package_stream_with_key,
    decrypt_agile_keys, parse_agile_encryption_info_stream,
    parse_agile_encryption_info_stream_with_options, AgileDataIntegrity, AgileDecryptedKeys,
    AgileEncryptionInfo, AgileKeyData, AgilePasswordKeyEncryptor,
    AgileEncryptionInfoWarning,
};
pub use agile_decrypt::decrypt_agile_encrypted_package;
#[allow(unused_imports)]
pub use crypto::{
    derive_iv, derive_key, derive_segment_iv, hash_password, segment_block_key, CryptoError,
    HashAlgorithm, HMAC_KEY_BLOCK, HMAC_VALUE_BLOCK, KEY_VALUE_BLOCK, VERIFIER_HASH_INPUT_BLOCK,
    VERIFIER_HASH_VALUE_BLOCK,
};
#[allow(unused_imports)]
pub use encryption_info::{
    decode_base64_field_limited, extract_encryption_info_xml, parse_agile_encryption_info_xml,
    AgileEncryptionInfoXml, EncryptionInfoWarning, ParseOptions, PasswordKeyEncryptor,
    KEY_ENCRYPTOR_URI_CERTIFICATE, KEY_ENCRYPTOR_URI_PASSWORD,
};
pub use error::{OffCryptoError, Result};
pub use ooxml::decrypt_ooxml_encrypted_package;

use std::io::{Cursor, Read, Seek};

/// Decrypt an Agile-encrypted OOXML package directly from an OLE CFB container.
///
/// This is a convenience wrapper around [`decrypt_agile_encrypted_package`] that handles stream
/// extraction from a `cfb::CompoundFile`.
///
/// Stream names are matched at the root level and support both common forms:
/// - `EncryptionInfo` / `EncryptedPackage`
/// - `/EncryptionInfo` / `/EncryptedPackage`
pub fn decrypt_agile_ooxml_from_cfb<R: Read + Seek>(
    cfb: &mut cfb::CompoundFile<R>,
    password: &str,
) -> Result<Vec<u8>> {
    let encryption_info = read_cfb_stream_bytes(cfb, "EncryptionInfo")?;
    let encrypted_package = read_cfb_stream_bytes(cfb, "EncryptedPackage")?;
    decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password)
}

/// Decrypt an Agile-encrypted OOXML package from an in-memory OLE/CFB container.
///
/// This helper opens the compound file container and delegates to
/// [`decrypt_agile_ooxml_from_cfb`].
pub fn decrypt_agile_ooxml_from_ole_bytes(ole_bytes: &[u8], password: &str) -> Result<Vec<u8>> {
    let mut cfb =
        cfb::CompoundFile::open(Cursor::new(ole_bytes)).map_err(|source| OffCryptoError::Io {
            source,
        })?;
    decrypt_agile_ooxml_from_cfb(&mut cfb, password)
}

/// Decrypt an Agile-encrypted OOXML package from an OLE/CFB reader.
///
/// This helper opens the compound file container and delegates to
/// [`decrypt_agile_ooxml_from_cfb`].
pub fn decrypt_agile_ooxml_from_ole_reader<R: Read + Seek>(
    reader: R,
    password: &str,
) -> Result<Vec<u8>> {
    let mut cfb = cfb::CompoundFile::open(reader).map_err(|source| OffCryptoError::Io { source })?;
    decrypt_agile_ooxml_from_cfb(&mut cfb, password)
}

fn read_cfb_stream_bytes<R: Read + Seek>(
    cfb: &mut cfb::CompoundFile<R>,
    name: &'static str,
) -> Result<Vec<u8>> {
    let mut stream = match cfb.open_stream(name) {
        Ok(s) => s,
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => {
            let with_slash = format!("/{name}");
            match cfb.open_stream(&with_slash) {
                Ok(s) => s,
                Err(err) if err.kind() == std::io::ErrorKind::NotFound => {
                    return Err(OffCryptoError::MissingRequiredStream {
                        stream: name.to_string(),
                    });
                }
                Err(err) => return Err(OffCryptoError::Io { source: err }),
            }
        }
        Err(err) => return Err(OffCryptoError::Io { source: err }),
    };

    let mut buf = Vec::new();
    stream
        .read_to_end(&mut buf)
        .map_err(|source| OffCryptoError::Io { source })?;
    Ok(buf)
}
