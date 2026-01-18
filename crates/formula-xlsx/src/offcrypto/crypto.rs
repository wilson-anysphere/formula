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
    fn eq_ignore_ascii_case_stripping_separators(s: &str, expected: &str) -> bool {
        fn next_non_sep(bytes: &[u8], mut i: usize) -> Option<(u8, usize)> {
            while i < bytes.len() {
                let b = bytes[i];
                if b != b'-' && b != b'_' {
                    return Some((b, i + 1));
                }
                i += 1;
            }
            None
        }

        let expected_bytes = expected.as_bytes();
        let bytes = s.as_bytes();
        let mut pos = 0usize;

        for &exp in expected_bytes {
            let Some((b, next)) = next_non_sep(bytes, pos) else {
                return false;
            };
            pos = next;
            if !b.is_ascii() {
                return false;
            }
            // `expected` is provided in lowercase ASCII.
            if b.to_ascii_lowercase() != exp {
                return false;
            }
        }

        // Ensure no remaining non-separator bytes.
        next_non_sep(bytes, pos).is_none()
    }

    /// Parse a hash algorithm name as used in MS-OFFCRYPTO XML.
    ///
    /// Names are case-insensitive (e.g. `SHA1`, `sha256`).
    pub fn parse_offcrypto_name(name: &str) -> Result<Self, CryptoError> {
        // MS-OFFCRYPTO XML typically uses `SHA1`/`SHA256` etc, but tolerate minor variations
        // (e.g. `SHA-256` / `sha_256`) seen in other tooling.
        let raw = name.trim();
        if Self::eq_ignore_ascii_case_stripping_separators(raw, "sha1") {
            return Ok(Self::Sha1);
        }
        if Self::eq_ignore_ascii_case_stripping_separators(raw, "sha256") {
            return Ok(Self::Sha256);
        }
        if Self::eq_ignore_ascii_case_stripping_separators(raw, "sha384") {
            return Ok(Self::Sha384);
        }
        if Self::eq_ignore_ascii_case_stripping_separators(raw, "sha512") {
            return Ok(Self::Sha512);
        }
        Err(CryptoError::UnsupportedHashAlgorithm(raw.to_string()))
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

    /// Hash a single buffer using this algorithm.
    pub fn hash(self, bytes: &[u8]) -> Vec<u8> {
        match self {
            HashAlgorithm::Sha1 => {
                let mut h = sha1::Sha1::new();
                h.update(bytes);
                h.finalize().to_vec()
            }
            HashAlgorithm::Sha256 => {
                let mut h = sha2::Sha256::new();
                h.update(bytes);
                h.finalize().to_vec()
            }
            HashAlgorithm::Sha384 => {
                let mut h = sha2::Sha384::new();
                h.update(bytes);
                h.finalize().to_vec()
            }
            HashAlgorithm::Sha512 => {
                let mut h = sha2::Sha512::new();
                h.update(bytes);
                h.finalize().to_vec()
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
    #[error("allocation failure: {0}")]
    AllocationFailure(&'static str),
}

fn normalize_key_material(bytes: &[u8], out_len: usize) -> Result<Vec<u8>, CryptoError> {
    if out_len == 0 {
        return Err(CryptoError::InvalidParameter("output length must be non-zero"));
    }
    if out_len > 1024 {
        // Excel's Agile encryption key/IV sizes are small (bytes, not KB). Refuse pathological sizes
        // to avoid attacker-controlled allocations.
        return Err(CryptoError::InvalidParameter("output length too large"));
    }

    let prefix_len = bytes.len().min(out_len);

    let mut out = Vec::new();
    if out.try_reserve_exact(out_len).is_err() {
        return Err(CryptoError::AllocationFailure("normalize_key_material output"));
    }

    // MS-OFFCRYPTO `TruncateHash` behavior:
    // - If the digest is longer than needed: truncate.
    // - If the digest is shorter: pad with 0x36 bytes (matches `msoffcrypto-tool`).
    out.extend_from_slice(&bytes[..prefix_len]);
    if prefix_len < out_len {
        out.resize(out_len, 0x36u8);
    }
    Ok(out)
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

fn update_password_utf16le<D: digest::Digest>(hasher: &mut D, password: &str) {
    // UTF-16LE with no BOM and no terminator.
    for unit in password.encode_utf16() {
        hasher.update(unit.to_le_bytes());
    }
}

#[cfg(test)]
fn password_utf16le_bytes(password: &str) -> Vec<u8> {
    // UTF-16LE with no BOM and no terminator.
    let mut out = Vec::new();
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

    let digest_len = hash_alg.digest_len();
    debug_assert!(digest_len <= MAX_DIGEST_LEN);

    // Avoid per-iteration allocations: keep the current hash in a fixed-size buffer and overwrite it
    // after each digest round.
    let mut h_buf = [0u8; MAX_DIGEST_LEN];

    match hash_alg {
        HashAlgorithm::Sha1 => {
            let mut d = sha1::Sha1::new();
            d.update(salt);
            update_password_utf16le(&mut d, password);
            h_buf[..20].copy_from_slice(&d.finalize());
            for i in 0..spin {
                let mut d = sha1::Sha1::new();
                d.update(i.to_le_bytes());
                d.update(&h_buf[..digest_len]);
                h_buf[..digest_len].copy_from_slice(&d.finalize());
            }
        }
        HashAlgorithm::Sha256 => {
            let mut d = sha2::Sha256::new();
            d.update(salt);
            update_password_utf16le(&mut d, password);
            h_buf[..32].copy_from_slice(&d.finalize());
            for i in 0..spin {
                let mut d = sha2::Sha256::new();
                d.update(i.to_le_bytes());
                d.update(&h_buf[..digest_len]);
                h_buf[..digest_len].copy_from_slice(&d.finalize());
            }
        }
        HashAlgorithm::Sha384 => {
            let mut d = sha2::Sha384::new();
            d.update(salt);
            update_password_utf16le(&mut d, password);
            h_buf[..48].copy_from_slice(&d.finalize());
            for i in 0..spin {
                let mut d = sha2::Sha384::new();
                d.update(i.to_le_bytes());
                d.update(&h_buf[..digest_len]);
                h_buf[..digest_len].copy_from_slice(&d.finalize());
            }
        }
        HashAlgorithm::Sha512 => {
            let mut d = sha2::Sha512::new();
            d.update(salt);
            update_password_utf16le(&mut d, password);
            h_buf[..64].copy_from_slice(&d.finalize());
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
/// 2. Apply `TruncateHash`/normalization to `key_len` bytes.
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
    hash_alg.hash_two_into(h, block_key, &mut digest[..digest_len]);

    normalize_key_material(&digest[..digest_len], key_len)
}

/// Derive an IV of `iv_len` bytes per MS-OFFCRYPTO Agile.
///
/// Algorithm:
/// 1. `IV = Hash(salt || blockKey)`
/// 2. Apply `TruncateHash`/normalization to `iv_len` bytes.
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
    hash_alg.hash_two_into(salt, block_key, &mut digest[..digest_len]);

    normalize_key_material(&digest[..digest_len], iv_len)
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
    use super::super::decrypt_aes_cbc_no_padding;
    use std::time::{Duration, Instant};

    fn ct_eq(a: &[u8], b: &[u8]) -> bool {
        let mut diff = 0u8;
        let max_len = a.len().max(b.len());
        for idx in 0..max_len {
            let av = a.get(idx).copied().unwrap_or(0);
            let bv = b.get(idx).copied().unwrap_or(0);
            diff |= av ^ bv;
        }
        diff == 0 && a.len() == b.len()
    }

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
    fn derive_key_truncates_and_pads_with_0x36() {
        let h = vec![0x22u8; 32];
        let full = derive_key(&h, &VERIFIER_HASH_INPUT_BLOCK, 20, HashAlgorithm::Sha1).unwrap();
        assert_eq!(full.len(), 20);

        let trunc = derive_key(&h, &VERIFIER_HASH_INPUT_BLOCK, 16, HashAlgorithm::Sha1).unwrap();
        assert_eq!(trunc.len(), 16);
        assert_eq!(&full[..16], &trunc[..]);

        let padded = derive_key(&h, &VERIFIER_HASH_INPUT_BLOCK, 24, HashAlgorithm::Sha1).unwrap();
        assert_eq!(padded.len(), 24);
        assert_eq!(&padded[..20], &full[..]);
        assert_eq!(&padded[20..], &[0x36u8; 4]);
    }

    #[test]
    fn derive_iv_truncates_and_pads_with_0x36() {
        let salt = [0x33u8; 16];
        let full = derive_iv(&salt, &KEY_VALUE_BLOCK, 20, HashAlgorithm::Sha1).unwrap();
        assert_eq!(full.len(), 20);

        let trunc = derive_iv(&salt, &KEY_VALUE_BLOCK, 12, HashAlgorithm::Sha1).unwrap();
        assert_eq!(trunc.len(), 12);
        assert_eq!(&full[..12], &trunc[..]);

        let padded = derive_iv(&salt, &KEY_VALUE_BLOCK, 28, HashAlgorithm::Sha1).unwrap();
        assert_eq!(padded.len(), 28);
        assert_eq!(&padded[..20], &full[..]);
        assert_eq!(&padded[20..], &[0x36u8; 8]);
    }

    #[test]
    fn normalize_key_material_pads_with_0x36() {
        assert_eq!(
            normalize_key_material(&[0xAA, 0xBB], 5).unwrap(),
            vec![0xAA, 0xBB, 0x36, 0x36, 0x36]
        );
    }

    #[test]
    fn normalize_key_material_truncates() {
        assert_eq!(
            normalize_key_material(&[0xAA, 0xBB, 0xCC], 2).unwrap(),
            vec![0xAA, 0xBB]
        );
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
        assert_eq!(
            HashAlgorithm::parse_offcrypto_name("sha_256").unwrap(),
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

    #[test]
    fn agile_makekey_from_password_vector_matches_msoffcrypto_tool() {
        // Test vector sourced from `msoffcrypto-tool` (Python) `ecma376_agile.py` docstring for
        // `ECMA376Agile.makekey_from_password`.
        let password = "Password1234_";
        let salt_value: [u8; 16] = [
            0x4c, 0x72, 0x5d, 0x45, 0xdc, 0x61, 0x0f, 0x93, 0x94, 0x12, 0xa0, 0x4d, 0xa7, 0x91,
            0x04, 0x66,
        ];
        let encrypted_key_value: [u8; 32] = [
            0xa1, 0x6c, 0xd5, 0x16, 0x5a, 0x7a, 0xb9, 0xd2, 0x71, 0x11, 0x3e, 0xd3, 0x86, 0xa7,
            0x8c, 0xf4, 0x96, 0x92, 0xe8, 0xe5, 0x27, 0xb0, 0xc5, 0xfc, 0x00, 0x55, 0xed, 0x08,
            0x0b, 0x7c, 0xb9, 0x4b,
        ];
        let expected_key: [u8; 32] = [
            0x40, 0x20, 0x66, 0x09, 0xd9, 0xfa, 0xad, 0xf2, 0x4b, 0x07, 0x6a, 0xeb, 0xf2, 0xc4,
            0x35, 0xb7, 0x42, 0x92, 0xc8, 0xb8, 0xa7, 0xaa, 0x81, 0xbc, 0x67, 0x9b, 0xe8, 0x97,
            0x11, 0xb0, 0x2a, 0xc2,
        ];

        let hash_alg = HashAlgorithm::Sha512;
        let password_hash = hash_password(password, &salt_value, 100_000, hash_alg).unwrap();
        let encryption_key = derive_key(&password_hash, &KEY_VALUE_BLOCK, 256 / 8, hash_alg).unwrap();
        let key = decrypt_aes_cbc_no_padding(&encryption_key, &salt_value, &encrypted_key_value)
            .expect("decrypt encryptedKeyValue");

        assert_eq!(key.as_slice(), expected_key);
    }

    #[test]
    fn agile_verify_password_vector_matches_msoffcrypto_tool() {
        // Test vector sourced from `msoffcrypto-tool` (Python) `ecma376_agile.py` docstring for
        // `ECMA376Agile.verify_password`.
        let salt_value: [u8; 16] = [
            0xcb, 0xca, 0x1c, 0x99, 0x93, 0x43, 0xfb, 0xad, 0x92, 0x07, 0x56, 0x34, 0x15, 0x00,
            0x34, 0xb0,
        ];
        let encrypted_verifier_hash_input: [u8; 16] = [
            0x39, 0xee, 0xa5, 0x4e, 0x26, 0xe5, 0x14, 0x79, 0x8c, 0x28, 0x4b, 0xc7, 0x71, 0x4d,
            0x38, 0xac,
        ];
        let encrypted_verifier_hash_value: [u8; 64] = [
            0x14, 0x37, 0x6d, 0x6d, 0x81, 0x73, 0x34, 0xe6, 0xb0, 0xff, 0x4f, 0xd8, 0x22, 0x1a,
            0x7c, 0x67, 0x8e, 0x5d, 0x8a, 0x78, 0x4e, 0x8f, 0x99, 0x9f, 0x4c, 0x18, 0x89, 0x30,
            0xc3, 0x6a, 0x4b, 0x29, 0xc5, 0xb3, 0x33, 0x60, 0x5b, 0x5c, 0xd4, 0x03, 0xb0, 0x50,
            0x03, 0xad, 0xcf, 0x18, 0xcc, 0xa8, 0xcb, 0xab, 0x8d, 0xeb, 0xe3, 0x73, 0xc6, 0x56,
            0x04, 0xa0, 0xbe, 0xcf, 0xae, 0x5c, 0x0a, 0xd0,
        ];

        let hash_alg = HashAlgorithm::Sha512;
        let key_len = 256 / 8;
        let expected_hash_len = hash_alg.digest_len();

        let verify = |password: &str| -> bool {
            let password_hash = hash_password(password, &salt_value, 100_000, hash_alg).unwrap();

            let k1 = derive_key(&password_hash, &VERIFIER_HASH_INPUT_BLOCK, key_len, hash_alg).unwrap();
            let verifier_input_padded = decrypt_aes_cbc_no_padding(
                &k1,
                &salt_value,
                &encrypted_verifier_hash_input,
            )
            .expect("decrypt encryptedVerifierHashInput");
            let verifier_input = verifier_input_padded;

            let mut computed_hash = [0u8; MAX_DIGEST_LEN];
            hash_alg.hash_two_into(&verifier_input, &[], &mut computed_hash);

            let k2 = derive_key(&password_hash, &VERIFIER_HASH_VALUE_BLOCK, key_len, hash_alg).unwrap();
            let expected_hash_padded = decrypt_aes_cbc_no_padding(
                &k2,
                &salt_value,
                &encrypted_verifier_hash_value,
            )
            .expect("decrypt encryptedVerifierHashValue");
            let expected_hash = expected_hash_padded
                .get(..expected_hash_len)
                .expect("expected hash");

            ct_eq(&computed_hash[..expected_hash_len], expected_hash)
        };

        assert!(verify("Password1234_"), "expected verifier to accept correct password");
        assert!(!verify("wrong password"), "expected verifier to reject incorrect password");
    }

    #[test]
    fn standard_cryptoapi_sha1_password_hash_vector_spin_50k() {
        // Standard/CryptoAPI encryption uses a fixed 50,000-iteration password hashing loop.
        // This vector is documented in `docs/offcrypto-standard-cryptoapi.md` and is useful for
        // catching off-by-one and byte-order regressions.
        let password = "password";
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b, 0x0c, 0x0d,
            0x0e, 0x0f,
        ];

        let expected: [u8; 20] = [
            0x1b, 0x59, 0x72, 0x28, 0x4e, 0xab, 0x64, 0x81, 0xeb, 0x65, 0x65, 0xa0, 0x98, 0x5b,
            0x33, 0x4b, 0x3e, 0x65, 0xe0, 0x41,
        ];

        let derived = hash_password(password, &salt, 50_000, HashAlgorithm::Sha1).unwrap();
        assert_eq!(derived.as_slice(), expected.as_slice());
    }
}
