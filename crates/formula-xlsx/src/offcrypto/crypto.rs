//! Crypto primitives for MS-OFFCRYPTO "Agile Encryption".
//!
//! This module implements the password hashing + key/IV derivation helpers described in
//! MS-OFFCRYPTO for Agile encryption.
//!
//! References:
//! - MS-OFFCRYPTO: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/

use digest::Digest as _;

const MAX_DIGEST_LEN: usize = 64; // SHA-512

/// Hash algorithm identifiers used by MS-OFFCRYPTO Agile encryption.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HashAlgorithm {
    Sha1,
    Sha256,
    Sha384,
    Sha512,
}

impl HashAlgorithm {
    /// Parse a hash algorithm name as used in MS-OFFCRYPTO XML.
    ///
    /// Names are case-insensitive (e.g. `SHA1`, `sha256`).
    pub fn parse_offcrypto_name(name: &str) -> Result<Self, CryptoError> {
        // MS-OFFCRYPTO XML typically uses `SHA1`/`SHA256` etc, but tolerate minor variations
        // (e.g. `SHA-256` / `sha_256`) seen in other tooling.
        let normalized = name
            .trim()
            .to_ascii_lowercase()
            .replace(['-', '_'], "");
        match normalized.as_str() {
            "sha1" => Ok(Self::Sha1),
            "sha256" => Ok(Self::Sha256),
            "sha384" => Ok(Self::Sha384),
            "sha512" => Ok(Self::Sha512),
            other => Err(CryptoError::UnsupportedHashAlgorithm(other.to_string())),
        }
    }

    fn digest_len(self) -> usize {
        match self {
            HashAlgorithm::Sha1 => 20,
            HashAlgorithm::Sha256 => 32,
            HashAlgorithm::Sha384 => 48,
            HashAlgorithm::Sha512 => 64,
        }
    }

    fn hash_two_into(self, a: &[u8], b: &[u8], out: &mut [u8]) {
        debug_assert!(
            out.len() >= self.digest_len(),
            "hash output buffer too small"
        );
        match self {
            HashAlgorithm::Sha1 => {
                let mut h = sha1::Sha1::new();
                h.update(a);
                h.update(b);
                out[..20].copy_from_slice(&h.finalize());
            }
            HashAlgorithm::Sha256 => {
                let mut h = sha2::Sha256::new();
                h.update(a);
                h.update(b);
                out[..32].copy_from_slice(&h.finalize());
            }
            HashAlgorithm::Sha384 => {
                let mut h = sha2::Sha384::new();
                h.update(a);
                h.update(b);
                out[..48].copy_from_slice(&h.finalize());
            }
            HashAlgorithm::Sha512 => {
                let mut h = sha2::Sha512::new();
                h.update(a);
                h.update(b);
                out[..64].copy_from_slice(&h.finalize());
            }
        }
    }
}

#[derive(Debug, thiserror::Error)]
pub enum CryptoError {
    #[error("unsupported hash algorithm: {0}")]
    UnsupportedHashAlgorithm(String),
    #[error("invalid parameter: {0}")]
    InvalidParameter(&'static str),
}

/// MS-OFFCRYPTO Agile: block key used for deriving the "verifierHashInput" key.
pub const VERIFIER_HASH_INPUT_BLOCK: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
/// MS-OFFCRYPTO Agile: block key used for deriving the "verifierHashValue" key.
pub const VERIFIER_HASH_VALUE_BLOCK: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
/// MS-OFFCRYPTO Agile: block key used for deriving the "keyValue" key.
pub const KEY_VALUE_BLOCK: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];
/// MS-OFFCRYPTO Agile: block key used for deriving the HMAC key.
pub const HMAC_KEY_BLOCK: [u8; 8] = [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];
/// MS-OFFCRYPTO Agile: block key used for deriving the HMAC value.
pub const HMAC_VALUE_BLOCK: [u8; 8] = [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];

fn password_utf16le_bytes(password: &str) -> Vec<u8> {
    // UTF-16LE with no BOM and no terminator.
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for ch in password.encode_utf16() {
        out.extend_from_slice(&ch.to_le_bytes());
    }
    out
}

/// Hash an Agile encryption password per MS-OFFCRYPTO.
///
/// Algorithm:
/// 1. `pw = UTF-16LE(password)` (no BOM, no terminator)
/// 2. `H = Hash(salt || pw)`
/// 3. For `i in 0..spinCount`: `H = Hash(LE32(i) || H)`
pub fn hash_password(
    password: &str,
    salt: &[u8],
    spin: u32,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>, CryptoError> {
    if salt.is_empty() {
        return Err(CryptoError::InvalidParameter("salt must be non-empty"));
    }

    let pw = password_utf16le_bytes(password);
    let digest_len = hash_alg.digest_len();
    debug_assert!(digest_len <= MAX_DIGEST_LEN);

    // Avoid per-iteration allocations: keep the current hash in a fixed-size buffer and overwrite it
    // after each digest round.
    let mut h_buf = [0u8; MAX_DIGEST_LEN];
    hash_alg.hash_two_into(salt, &pw, &mut h_buf[..digest_len]);

    match hash_alg {
        HashAlgorithm::Sha1 => {
            for i in 0..spin {
                let mut d = sha1::Sha1::new();
                d.update(i.to_le_bytes());
                d.update(&h_buf[..digest_len]);
                h_buf[..digest_len].copy_from_slice(&d.finalize());
            }
        }
        HashAlgorithm::Sha256 => {
            for i in 0..spin {
                let mut d = sha2::Sha256::new();
                d.update(i.to_le_bytes());
                d.update(&h_buf[..digest_len]);
                h_buf[..digest_len].copy_from_slice(&d.finalize());
            }
        }
        HashAlgorithm::Sha384 => {
            for i in 0..spin {
                let mut d = sha2::Sha384::new();
                d.update(i.to_le_bytes());
                d.update(&h_buf[..digest_len]);
                h_buf[..digest_len].copy_from_slice(&d.finalize());
            }
        }
        HashAlgorithm::Sha512 => {
            for i in 0..spin {
                let mut d = sha2::Sha512::new();
                d.update(i.to_le_bytes());
                d.update(&h_buf[..digest_len]);
                h_buf[..digest_len].copy_from_slice(&d.finalize());
            }
        }
    }

    Ok(h_buf[..digest_len].to_vec())
}

/// Derive a key of `key_len` bytes per MS-OFFCRYPTO Agile.
///
/// Algorithm:
/// 1. `K = Hash(H || blockKey)`
/// 2. If `key_len > digest_len`, append `0x00` bytes; else truncate.
pub fn derive_key(
    h: &[u8],
    block_key: &[u8],
    key_len: usize,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>, CryptoError> {
    if h.is_empty() {
        return Err(CryptoError::InvalidParameter("password hash must be non-empty"));
    }
    if block_key.is_empty() {
        return Err(CryptoError::InvalidParameter("block_key must be non-empty"));
    }

    let digest_len = hash_alg.digest_len();

    // Avoid allocating a temporary `H || blockKey` buffer: update the digest twice.
    let mut digest = [0u8; MAX_DIGEST_LEN];
    hash_alg.hash_two_into(h, block_key, &mut digest);

    if key_len <= digest_len {
        Ok(digest[..key_len].to_vec())
    } else {
        let mut out = vec![0u8; key_len];
        out[..digest_len].copy_from_slice(&digest[..digest_len]);
        Ok(out)
    }
}

/// Derive an IV of `iv_len` bytes per MS-OFFCRYPTO Agile.
///
/// Algorithm:
/// 1. `IV = Hash(salt || blockKey)`
/// 2. If `iv_len > digest_len`, append `0x00` bytes; else truncate.
pub fn derive_iv(
    salt: &[u8],
    block_key: &[u8],
    iv_len: usize,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>, CryptoError> {
    if salt.is_empty() {
        return Err(CryptoError::InvalidParameter("salt must be non-empty"));
    }
    if block_key.is_empty() {
        return Err(CryptoError::InvalidParameter("block_key must be non-empty"));
    }

    let digest_len = hash_alg.digest_len();

    // Avoid allocating a temporary `salt || blockKey` buffer: update the digest twice.
    let mut digest = [0u8; MAX_DIGEST_LEN];
    hash_alg.hash_two_into(salt, block_key, &mut digest);

    if iv_len <= digest_len {
        Ok(digest[..iv_len].to_vec())
    } else {
        let mut out = vec![0u8; iv_len];
        out[..digest_len].copy_from_slice(&digest[..digest_len]);
        Ok(out)
    }
}

/// Block key for `EncryptedPackage` segment IV derivation.
///
/// MS-OFFCRYPTO Agile uses a per-segment block key equal to `LE32(segment_index)`.
#[inline]
pub fn segment_block_key(segment_index: u32) -> [u8; 4] {
    segment_index.to_le_bytes()
}

/// Derive an IV for a specific `EncryptedPackage` segment.
///
/// This is a convenience wrapper for:
/// `derive_iv(salt, &LE32(segment_index), iv_len, hash_alg)`.
pub fn derive_segment_iv(
    salt: &[u8],
    segment_index: u32,
    iv_len: usize,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>, CryptoError> {
    let block_key = segment_block_key(segment_index);
    derive_iv(salt, &block_key, iv_len, hash_alg)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::time::{Duration, Instant};

    #[test]
    fn password_utf16le_encoding_no_bom_no_terminator() {
        assert_eq!(password_utf16le_bytes("A"), vec![0x41, 0x00]);
        assert_eq!(password_utf16le_bytes("AB"), vec![0x41, 0x00, 0x42, 0x00]);

        // Non-BMP char (😀 = U+1F600) encodes as surrogate pair: D83D DE00 (LE bytes).
        assert_eq!(
            password_utf16le_bytes("😀"),
            vec![0x3D, 0xD8, 0x00, 0xDE]
        );
    }

    #[test]
    fn spin_count_changes_password_hash() {
        let salt = [0x11u8; 16];
        let h0 = hash_password("password", &salt, 0, HashAlgorithm::Sha256).unwrap();
        let h1 = hash_password("password", &salt, 1, HashAlgorithm::Sha256).unwrap();
        assert_ne!(h0, h1, "spinCount=0 and spinCount=1 must differ");
        assert_eq!(h0.len(), HashAlgorithm::Sha256.digest_len());
    }

    #[test]
    fn derive_key_truncates_and_pads_with_zeros() {
        let h = vec![0x22u8; 32];
        let full = derive_key(&h, &VERIFIER_HASH_INPUT_BLOCK, 20, HashAlgorithm::Sha1).unwrap();
        assert_eq!(full.len(), 20);

        let trunc = derive_key(&h, &VERIFIER_HASH_INPUT_BLOCK, 16, HashAlgorithm::Sha1).unwrap();
        assert_eq!(trunc.len(), 16);
        assert_eq!(&full[..16], &trunc[..]);

        let padded = derive_key(&h, &VERIFIER_HASH_INPUT_BLOCK, 24, HashAlgorithm::Sha1).unwrap();
        assert_eq!(padded.len(), 24);
        assert_eq!(&padded[..20], &full[..]);
        assert_eq!(&padded[20..], &[0u8; 4]);
    }

    #[test]
    fn derive_iv_truncates_and_pads_with_zeros() {
        let salt = [0x33u8; 16];
        let full = derive_iv(&salt, &KEY_VALUE_BLOCK, 20, HashAlgorithm::Sha1).unwrap();
        assert_eq!(full.len(), 20);

        let trunc = derive_iv(&salt, &KEY_VALUE_BLOCK, 12, HashAlgorithm::Sha1).unwrap();
        assert_eq!(trunc.len(), 12);
        assert_eq!(&full[..12], &trunc[..]);

        let padded = derive_iv(&salt, &KEY_VALUE_BLOCK, 28, HashAlgorithm::Sha1).unwrap();
        assert_eq!(padded.len(), 28);
        assert_eq!(&padded[..20], &full[..]);
        assert_eq!(&padded[20..], &[0u8; 8]);
    }

    #[test]
    fn rejects_empty_salt_or_block_key() {
        let err = hash_password("pw", &[], 0, HashAlgorithm::Sha1).unwrap_err();
        assert!(matches!(err, CryptoError::InvalidParameter(_)));

        let err = derive_key(&[1, 2, 3], &[], 16, HashAlgorithm::Sha1).unwrap_err();
        assert!(matches!(err, CryptoError::InvalidParameter(_)));

        let err = derive_iv(&[], &[1], 16, HashAlgorithm::Sha1).unwrap_err();
        assert!(matches!(err, CryptoError::InvalidParameter(_)));
    }

    #[test]
    fn parse_hash_algorithm_name() {
        assert_eq!(
            HashAlgorithm::parse_offcrypto_name("SHA1").unwrap(),
            HashAlgorithm::Sha1
        );
        assert_eq!(
            HashAlgorithm::parse_offcrypto_name("sha512").unwrap(),
            HashAlgorithm::Sha512
        );
        assert_eq!(
            HashAlgorithm::parse_offcrypto_name("SHA-256").unwrap(),
            HashAlgorithm::Sha256
        );
        let err = HashAlgorithm::parse_offcrypto_name("md5").unwrap_err();
        assert!(matches!(err, CryptoError::UnsupportedHashAlgorithm(_)));
    }

    #[test]
    fn hash_password_perf_guard_spin_10k() {
        // A simple regression guard: the Agile password KDF loop is often 100k iterations in the
        // real world, so even modest per-iteration overhead (e.g. heap allocations) can become
        // noticeable. This test uses a smaller spinCount so it stays fast in debug CI while still
        // exercising the hot loop.
        let salt = [0x11u8; 16];
        let start = Instant::now();
        let _ = hash_password("password", &salt, 10_000, HashAlgorithm::Sha256).unwrap();
        assert!(
            start.elapsed() < Duration::from_secs(2),
            "hash_password(spinCount=10_000) took too long: {:?}",
            start.elapsed()
        );
    }

    #[test]
    fn segment_block_key_is_le32() {
        assert_eq!(segment_block_key(0), [0, 0, 0, 0]);
        assert_eq!(segment_block_key(0x1122_3344), [0x44, 0x33, 0x22, 0x11]);
    }

    #[test]
    fn derive_segment_iv_matches_generic_derive_iv() {
        let salt = [0x44u8; 16];
        let idx = 7u32;
        let a = derive_segment_iv(&salt, idx, 16, HashAlgorithm::Sha1).unwrap();
        let b = derive_iv(&salt, &idx.to_le_bytes(), 16, HashAlgorithm::Sha1).unwrap();
        assert_eq!(a, b);
    }
}
