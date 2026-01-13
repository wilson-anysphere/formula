//! MS-OFFCRYPTO support.
//!
//! This crate currently implements a small subset of the MS-OFFCRYPTO spec:
//! - Parsing of the `EncryptionInfo` stream for **Standard** (CryptoAPI) encryption.
//! - Decryption of the `EncryptedPackage` stream for Standard encryption, given a
//!   derived AES key.
//!
//! Key derivation and Agile encryption are not implemented yet.

#![forbid(unsafe_code)]

use aes::cipher::{generic_array::GenericArray, BlockDecrypt, KeyInit};
use thiserror::Error;

/// Errors returned by this crate.
#[derive(Debug, Error, PartialEq, Eq)]
pub enum OffcryptoError {
    #[error("EncryptionInfo stream is truncated")]
    Truncated,

    #[error("invalid EncryptionInfo: {0}")]
    InvalidEncryptionInfo(&'static str),

    #[error(
        "unsupported EncryptionInfo version (major={version_major}, minor={version_minor})"
    )]
    UnsupportedEncryptionInfoVersion { version_major: u16, version_minor: u16 },

    #[error("unsupported encryption algorithm id 0x{0:08X}")]
    UnsupportedAlgorithm(u32),

    #[error("EncryptedPackage stream is too short: expected at least 8 bytes, got {len}")]
    EncryptedPackageTooShort { len: usize },

    #[error(
        "EncryptedPackage ciphertext length must be a multiple of 16 bytes for AES-ECB, got {len}"
    )]
    InvalidCiphertextLength { len: usize },

    #[error("invalid AES key length {len}; expected 16, 24, or 32 bytes")]
    InvalidKeyLength { len: usize },

    #[error(
        "EncryptedPackage declared plaintext size {total_size} exceeds decrypted length {decrypted_len}"
    )]
    TotalSizeOutOfBounds { total_size: u64, decrypted_len: usize },

    #[error("EncryptedPackage declared plaintext size {total_size} does not fit into usize")]
    TotalSizeTooLarge { total_size: u64 },
}

/// Decrypts the `EncryptedPackage` stream for ECMA-376 Standard encryption.
///
/// The stream layout is:
/// - `total_size` (u64, little-endian) at bytes 0..8
/// - AES-ECB ciphertext at bytes 8..
///
/// The ciphertext is decrypted in full, and the returned plaintext is truncated to `total_size`.
pub fn standard_decrypt_package(
    key: &[u8],
    encrypted_package: &[u8],
) -> Result<Vec<u8>, OffcryptoError> {
    if encrypted_package.len() < 8 {
        return Err(OffcryptoError::EncryptedPackageTooShort {
            len: encrypted_package.len(),
        });
    }

    let total_size = u64::from_le_bytes(
        encrypted_package[0..8]
            .try_into()
            .expect("slice length checked"),
    );

    let ciphertext = &encrypted_package[8..];
    if ciphertext.len() % 16 != 0 {
        return Err(OffcryptoError::InvalidCiphertextLength {
            len: ciphertext.len(),
        });
    }

    let mut decrypted = ciphertext.to_vec();
    match key.len() {
        16 => {
            let cipher = aes::Aes128::new_from_slice(key)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            for block in decrypted.chunks_mut(16) {
                cipher.decrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        24 => {
            let cipher = aes::Aes192::new_from_slice(key)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            for block in decrypted.chunks_mut(16) {
                cipher.decrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        32 => {
            let cipher = aes::Aes256::new_from_slice(key)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            for block in decrypted.chunks_mut(16) {
                cipher.decrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        _ => return Err(OffcryptoError::InvalidKeyLength { len: key.len() }),
    }

    let total_size_usize: usize = total_size
        .try_into()
        .map_err(|_| OffcryptoError::TotalSizeTooLarge { total_size })?;

    if total_size_usize > decrypted.len() {
        return Err(OffcryptoError::TotalSizeOutOfBounds {
            total_size,
            decrypted_len: decrypted.len(),
        });
    }

    decrypted.truncate(total_size_usize);
    Ok(decrypted)
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EncryptionInfo {
    Standard(StandardEncryptionInfo),
}

impl EncryptionInfo {
    /// Parse an `EncryptionInfo` stream.
    pub fn parse(bytes: &[u8]) -> Result<Self, OffcryptoError> {
        let mut cur = ByteCursor::new(bytes);
        let version_major = cur.read_u16_le()?;
        let version_minor = cur.read_u16_le()?;

        // Standard encryption is identified by versionMinor == 2 and versionMajor in {2,3,4}.
        if version_minor == 2 && matches!(version_major, 2 | 3 | 4) {
            return Ok(Self::Standard(parse_standard_encryption_info(&mut cur)?));
        }

        Err(OffcryptoError::UnsupportedEncryptionInfoVersion {
            version_major,
            version_minor,
        })
    }
}

/// Parsed contents of a Standard (CryptoAPI) `EncryptionInfo` stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StandardEncryptionInfo {
    pub header: StandardHeader,
    pub verifier: StandardVerifier,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StandardHeader {
    pub flags: u32,
    pub size_extra: u32,
    pub alg_id: u32,
    pub alg_id_hash: u32,
    pub key_size: u32,
    pub provider_type: u32,
    pub reserved1: u32,
    pub reserved2: u32,
    pub csp_name: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StandardVerifier {
    pub salt_size: u32,
    pub salt: [u8; 16],
    pub encrypted_verifier: [u8; 16],
    pub verifier_hash_size: u32,
    pub encrypted_verifier_hash: [u8; 32],
}

const CALG_AES_128: u32 = 0x0000_660E;
const CALG_AES_192: u32 = 0x0000_660F;
const CALG_AES_256: u32 = 0x0000_6610;

fn parse_standard_encryption_info(
    cur: &mut ByteCursor<'_>,
) -> Result<StandardEncryptionInfo, OffcryptoError> {
    // header_flags: u32 (ignored for decryption)
    let _header_flags = cur.read_u32_le()?;

    let encryption_header_size = cur.read_u32_le()? as usize;
    let encryption_header_bytes = cur.take(encryption_header_size)?;

    let header = parse_standard_header(encryption_header_bytes)?;
    let verifier = parse_standard_verifier(cur)?;

    Ok(StandardEncryptionInfo { header, verifier })
}

fn parse_standard_header(bytes: &[u8]) -> Result<StandardHeader, OffcryptoError> {
    // The fixed portion of the EncryptionHeader is 8 u32 fields = 32 bytes.
    const FIXED_LEN: usize = 8 * 4;
    if bytes.len() < FIXED_LEN {
        return Err(OffcryptoError::InvalidEncryptionInfo(
            "encryption header is too short",
        ));
    }

    let mut cur = ByteCursor::new(bytes);
    let flags = cur.read_u32_le()?;
    let size_extra = cur.read_u32_le()?;
    let alg_id = cur.read_u32_le()?;
    let alg_id_hash = cur.read_u32_le()?;
    let key_size = cur.read_u32_le()?;
    let provider_type = cur.read_u32_le()?;
    let reserved1 = cur.read_u32_le()?;
    let reserved2 = cur.read_u32_le()?;

    // We currently only support AES (Office 2007+ Standard encryption).
    if !matches!(alg_id, CALG_AES_128 | CALG_AES_192 | CALG_AES_256) {
        return Err(OffcryptoError::UnsupportedAlgorithm(alg_id));
    }

    let csp_bytes = cur.remaining_bytes();
    let csp_name = decode_utf16le_null_terminated(csp_bytes)?;

    Ok(StandardHeader {
        flags,
        size_extra,
        alg_id,
        alg_id_hash,
        key_size,
        provider_type,
        reserved1,
        reserved2,
        csp_name,
    })
}

fn parse_standard_verifier(cur: &mut ByteCursor<'_>) -> Result<StandardVerifier, OffcryptoError> {
    let salt_size = cur.read_u32_le()?;
    if salt_size != 16 {
        return Err(OffcryptoError::InvalidEncryptionInfo(
            "standard encryption verifier salt_size must be 16",
        ));
    }

    let salt = cur.read_array::<16>()?;
    let encrypted_verifier = cur.read_array::<16>()?;
    let verifier_hash_size = cur.read_u32_le()?;
    let encrypted_verifier_hash = cur.read_array::<32>()?;

    Ok(StandardVerifier {
        salt_size,
        salt,
        encrypted_verifier,
        verifier_hash_size,
        encrypted_verifier_hash,
    })
}

fn decode_utf16le_null_terminated(bytes: &[u8]) -> Result<String, OffcryptoError> {
    if bytes.is_empty() {
        return Ok(String::new());
    }
    if bytes.len() % 2 != 0 {
        return Err(OffcryptoError::InvalidEncryptionInfo(
            "UTF-16LE string has odd byte length",
        ));
    }

    let mut units = Vec::with_capacity(bytes.len() / 2);
    for chunk in bytes.chunks_exact(2) {
        units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }

    let end = units.iter().position(|&u| u == 0).unwrap_or(units.len());
    Ok(String::from_utf16_lossy(&units[..end]))
}

struct ByteCursor<'a> {
    bytes: &'a [u8],
    pos: usize,
}

impl<'a> ByteCursor<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, pos: 0 }
    }

    fn remaining_bytes(&self) -> &'a [u8] {
        &self.bytes[self.pos..]
    }

    fn take(&mut self, n: usize) -> Result<&'a [u8], OffcryptoError> {
        if self.pos.checked_add(n).is_none() {
            return Err(OffcryptoError::Truncated);
        }
        if self.pos + n > self.bytes.len() {
            return Err(OffcryptoError::Truncated);
        }
        let out = &self.bytes[self.pos..self.pos + n];
        self.pos += n;
        Ok(out)
    }

    fn read_u16_le(&mut self) -> Result<u16, OffcryptoError> {
        let b = self.take(2)?;
        Ok(u16::from_le_bytes([b[0], b[1]]))
    }

    fn read_u32_le(&mut self) -> Result<u32, OffcryptoError> {
        let b = self.take(4)?;
        Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
    }

    fn read_array<const N: usize>(&mut self) -> Result<[u8; N], OffcryptoError> {
        let b = self.take(N)?;
        Ok(b.try_into().expect("slice length checked"))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use aes::cipher::{BlockEncrypt, KeyInit};

    fn push_u16(buf: &mut Vec<u8>, v: u16) {
        buf.extend_from_slice(&v.to_le_bytes());
    }

    fn push_u32(buf: &mut Vec<u8>, v: u32) {
        buf.extend_from_slice(&v.to_le_bytes());
    }

    #[test]
    fn standard_decrypt_package_roundtrip() {
        let key = [0x42u8; 16];

        let plaintext: Vec<u8> = (0u8..=42).collect(); // 43 bytes; exercises truncation
        let total_size = plaintext.len() as u64;

        let mut padded = plaintext.clone();
        padded.resize((padded.len() + 15) / 16 * 16, 0u8);

        let cipher = aes::Aes128::new_from_slice(&key).expect("valid AES-128 key");
        let mut ciphertext = padded.clone();
        for block in ciphertext.chunks_mut(16) {
            cipher.encrypt_block(GenericArray::from_mut_slice(block));
        }

        let mut encrypted_package = total_size.to_le_bytes().to_vec();
        encrypted_package.extend_from_slice(&ciphertext);

        let decrypted = standard_decrypt_package(&key, &encrypted_package).expect("decrypt");
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn parses_synthetic_standard_encryption_info() {
        let mut buf = Vec::new();

        let salt: [u8; 16] = std::array::from_fn(|i| i as u8);
        let encrypted_verifier: [u8; 16] = std::array::from_fn(|i| (i + 16) as u8);
        let encrypted_verifier_hash: [u8; 32] = std::array::from_fn(|i| (i + 32) as u8);

        // versionMajor/versionMinor
        push_u16(&mut buf, 4);
        push_u16(&mut buf, 2);

        // header_flags (ignored)
        push_u32(&mut buf, 0);

        // Build EncryptionHeader bytes.
        let csp_name = "Test CSP";
        let mut csp_bytes = Vec::new();
        for u in csp_name.encode_utf16() {
            csp_bytes.extend_from_slice(&u.to_le_bytes());
        }
        // Null terminator.
        csp_bytes.extend_from_slice(&0u16.to_le_bytes());

        let mut header_bytes = Vec::new();
        push_u32(&mut header_bytes, 0x0000_0024); // flags
        push_u32(&mut header_bytes, 0); // size_extra
        push_u32(&mut header_bytes, CALG_AES_128); // alg_id
        push_u32(&mut header_bytes, 0x0000_8004); // alg_id_hash (SHA1)
        push_u32(&mut header_bytes, 128); // key_size (bits)
        push_u32(&mut header_bytes, 0x0000_0018); // provider_type (PROV_RSA_AES)
        push_u32(&mut header_bytes, 0); // reserved1
        push_u32(&mut header_bytes, 0); // reserved2
        header_bytes.extend_from_slice(&csp_bytes);

        push_u32(&mut buf, header_bytes.len() as u32);
        buf.extend_from_slice(&header_bytes);

        // EncryptionVerifier.
        push_u32(&mut buf, 16); // salt_size
        buf.extend_from_slice(&salt);
        buf.extend_from_slice(&encrypted_verifier);
        push_u32(&mut buf, 20); // verifier_hash_size (SHA1)
        buf.extend_from_slice(&encrypted_verifier_hash);

        let parsed = EncryptionInfo::parse(&buf).expect("parse");
        let EncryptionInfo::Standard(info) = parsed;

        assert_eq!(info.header.alg_id, CALG_AES_128);
        assert_eq!(info.header.key_size, 128);
        assert_eq!(info.header.csp_name, csp_name);

        assert_eq!(info.verifier.salt_size, 16);
        assert_eq!(info.verifier.salt, salt);
        assert_eq!(info.verifier.encrypted_verifier, encrypted_verifier);
        assert_eq!(info.verifier.verifier_hash_size, 20);
        assert_eq!(info.verifier.encrypted_verifier_hash, encrypted_verifier_hash);
    }

    #[test]
    fn rejects_unsupported_alg_id() {
        let mut buf = Vec::new();
        push_u16(&mut buf, 4);
        push_u16(&mut buf, 2);
        push_u32(&mut buf, 0);

        let csp_bytes = 0u16.to_le_bytes();
        let mut header_bytes = Vec::new();
        push_u32(&mut header_bytes, 0);
        push_u32(&mut header_bytes, 0);
        push_u32(&mut header_bytes, 0xDEAD_BEEF); // alg_id
        push_u32(&mut header_bytes, 0);
        push_u32(&mut header_bytes, 128);
        push_u32(&mut header_bytes, 0);
        push_u32(&mut header_bytes, 0);
        push_u32(&mut header_bytes, 0);
        header_bytes.extend_from_slice(&csp_bytes);

        push_u32(&mut buf, header_bytes.len() as u32);
        buf.extend_from_slice(&header_bytes);

        // Minimal verifier.
        push_u32(&mut buf, 16);
        buf.extend([0u8; 16]);
        buf.extend([0u8; 16]);
        push_u32(&mut buf, 20);
        buf.extend([0u8; 32]);

        let err = EncryptionInfo::parse(&buf).expect_err("expected unsupported alg");
        assert_eq!(err, OffcryptoError::UnsupportedAlgorithm(0xDEAD_BEEF));
    }
}
