//! MS-OFFCRYPTO *Standard* (CryptoAPI) key derivation helpers.
//!
//! These primitives are used by the "Standard Encryption" mode described in
//! MS-OFFCRYPTO (CryptoAPI). Unlike the Agile encryption mode, the spin count
//! is fixed at **50,000** iterations and is not stored in the file.
//!
//! This module intentionally only implements the pieces needed for future
//! `EncryptedPackage` decryption support:
//! - password UTF-16LE encoding (no NUL terminator)
//! - fixed-spin password hashing (50k)
//! - final block hash (H || block)
//! - CryptoAPI `CryptDeriveKey` byte expansion (HMAC-like)
use md5::Md5;
use sha1::Sha1;
use sha1::Digest as _;

const MAX_DIGEST_LEN: usize = 20; // SHA-1

/// CryptoAPI algorithm identifier for MD5.
pub const CALG_MD5: u32 = 0x0000_8003;
/// CryptoAPI algorithm identifier for SHA-1.
pub const CALG_SHA1: u32 = 0x0000_8004;

/// Hash algorithms supported by the MS-OFFCRYPTO Standard (CryptoAPI) mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HashAlg {
    /// MD5 (16-byte digest).
    Md5,
    /// SHA-1 (20-byte digest).
    Sha1,
}

/// Errors returned by this module when an unsupported algorithm is requested.
#[derive(Debug, thiserror::Error)]
pub enum CryptoApiError {
    #[error(
        "unsupported hash algorithm CALG id {calg_id:#010x} (supported: CALG_MD5={CALG_MD5:#010x}, CALG_SHA1={CALG_SHA1:#010x})"
    )]
    UnsupportedHashAlg { calg_id: u32 },
}

impl HashAlg {
    /// Convert a CryptoAPI `CALG_*` id to a supported [`HashAlg`].
    pub fn from_calg_id(calg_id: u32) -> Result<Self, CryptoApiError> {
        match calg_id {
            CALG_MD5 => Ok(Self::Md5),
            CALG_SHA1 => Ok(Self::Sha1),
            _ => Err(CryptoApiError::UnsupportedHashAlg { calg_id }),
        }
    }

    /// Digest length in bytes.
    pub fn hash_len(self) -> usize {
        match self {
            HashAlg::Md5 => 16,
            HashAlg::Sha1 => 20,
        }
    }
}

enum Hasher {
    Md5(Md5),
    Sha1(Sha1),
}

impl Hasher {
    fn new(alg: HashAlg) -> Self {
        match alg {
            HashAlg::Md5 => Self::Md5(Md5::new()),
            HashAlg::Sha1 => Self::Sha1(Sha1::new()),
        }
    }

    fn update(&mut self, bytes: &[u8]) {
        match self {
            Hasher::Md5(h) => h.update(bytes),
            Hasher::Sha1(h) => h.update(bytes),
        }
    }

    fn finalize_into(self, out: &mut [u8]) {
        match self {
            Hasher::Md5(h) => {
                debug_assert_eq!(out.len(), 16);
                out.copy_from_slice(&h.finalize());
            }
            Hasher::Sha1(h) => {
                debug_assert_eq!(out.len(), 20);
                out.copy_from_slice(&h.finalize());
            }
        }
    }

    fn finalize_reset_into(&mut self, out: &mut [u8]) {
        match self {
            Hasher::Md5(h) => {
                debug_assert_eq!(out.len(), 16);
                out.copy_from_slice(&h.finalize_reset());
            }
            Hasher::Sha1(h) => {
                debug_assert_eq!(out.len(), 20);
                out.copy_from_slice(&h.finalize_reset());
            }
        }
    }
}

/// Encode a password as UTF-16LE bytes (no BOM, no NUL terminator).
pub fn password_to_utf16le(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

/// Hash `salt || password_utf16le`, then iterate a fixed 50,000 "spin" rounds:
///
/// ```text
/// H = Hash(salt || password)
/// for i in 0..50000:
///   H = Hash(LE32(i) || H)
/// ```
pub fn hash_password_fixed_spin(
    password_utf16le: &[u8],
    salt: &[u8],
    hash_alg: HashAlg,
) -> Vec<u8> {
    let hash_len = hash_alg.hash_len();
    let mut hasher = Hasher::new(hash_alg);
    hasher.update(salt);
    hasher.update(password_utf16le);
    let mut h = [0u8; MAX_DIGEST_LEN];
    hasher.finalize_reset_into(&mut h[..hash_len]);

    for i in 0u32..50_000u32 {
        hasher.update(&i.to_le_bytes());
        hasher.update(&h[..hash_len]);
        hasher.finalize_reset_into(&mut h[..hash_len]);
    }
    h[..hash_len].to_vec()
}

/// Compute `Hash(H || LE32(block))`.
pub fn final_hash(h: &[u8], block: u32, hash_alg: HashAlg) -> Vec<u8> {
    let mut hasher = Hasher::new(hash_alg);
    hasher.update(h);
    hasher.update(&block.to_le_bytes());
    let hash_len = hash_alg.hash_len();
    let mut out = [0u8; MAX_DIGEST_LEN];
    hasher.finalize_into(&mut out[..hash_len]);
    out[..hash_len].to_vec()
}

/// CryptoAPI `CryptDeriveKey` byte expansion used by MS-OFFCRYPTO Standard encryption.
///
/// CryptoAPI `CryptDeriveKey` is **not** PBKDF2 and does **not** use the digest bytes directly as
/// the key. For MD5/SHA-1 it expands the digest using an ipad/opad construction (HMAC-like) and
/// then truncates to the desired key length.
///
/// ```text
/// D = hash_value || 0x00*(64 - hash_len)
/// key = Hash(D XOR 0x36...) || Hash(D XOR 0x5C...)
/// return key[0..key_len_bytes]
/// ```
///
/// Note: For MS-OFFCRYPTO Standard AES encryption, Office/CryptoAPI use this derivation even when
/// `key_len_bytes <= hash_len` (i.e. AES-128 keys are **not** a direct truncation of the digest).
pub fn crypt_derive_key(hash_value: &[u8], key_len_bytes: usize, hash_alg: HashAlg) -> Vec<u8> {
    let hash_len = hash_alg.hash_len();
    debug_assert_eq!(
        hash_value.len(),
        hash_len,
        "hash_value len must match hash_alg.hash_len()"
    );
    assert!(
        key_len_bytes <= hash_len * 2,
        "CryptoAPI CryptDeriveKey only yields up to {} bytes of key material for {:?} (requested {})",
        hash_len * 2,
        hash_alg,
        key_len_bytes
    );

    // For MS-OFFCRYPTO Standard AES, CryptoAPI runs the ipad/opad expansion even when the requested
    // key length is <= the digest length (e.g. AES-128 derived from MD5/SHA-1).
    // The MS-OFFCRYPTO Standard mode only uses MD5/SHA-1, both of which have a 64-byte block size.
    // `hash_len` is guaranteed <= 64.
    let mut buf = [0u8; 64];
    buf[..hash_len].copy_from_slice(hash_value);

    let mut ipad = [0u8; 64];
    let mut opad = [0u8; 64];
    for i in 0..64 {
        ipad[i] = buf[i] ^ 0x36;
        opad[i] = buf[i] ^ 0x5C;
    }

    let mut key = [0u8; MAX_DIGEST_LEN * 2];
    let mut hasher = Hasher::new(hash_alg);
    hasher.update(&ipad);
    hasher.finalize_reset_into(&mut key[..hash_len]);
    hasher.update(&opad);
    hasher.finalize_reset_into(&mut key[hash_len..hash_len * 2]);

    key[..key_len_bytes].to_vec()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn crypt_derive_key_sha1_aes128_runs_expansion_even_when_key_len_lt_hash_len() {
        // MS-OFFCRYPTO Standard AES-128 `CryptDeriveKey` sanity check:
        //
        // Given a per-block SHA-1 digest `H_block0`, CryptoAPI's `CryptDeriveKey` performs the
        // ipad/opad expansion and *then* truncates to 16 bytes. It does **not** use
        // `H_block0[0:16]` directly.
        //
        // Test vector from `docs/offcrypto-standard-cryptoapi.md` (§8.2).
        let hash_value: [u8; 20] = [
            0xE2, 0xF8, 0xCD, 0xE4, 0x57, 0xE5, 0xD4, 0x49, 0xEB, 0x20, 0x50, 0x57, 0xC8,
            0x8D, 0x20, 0x1D, 0x14, 0x53, 0x1F, 0xF3,
        ];

        // Expected bytes computed independently with Python:
        //   buf = hash_value || 0x00*(64-hash_len)
        //   key = sha1(buf^0x36) || sha1(buf^0x5c)
        let key = crypt_derive_key(&hash_value, 16, HashAlg::Sha1);
        let expected: [u8; 16] = [
            0x40, 0xB1, 0x3A, 0x71, 0xF9, 0x0B, 0x96, 0x6E, 0x37, 0x54, 0x08, 0xF2, 0xD1,
            0x81, 0xA1, 0xAA,
        ];
        assert_eq!(key, expected);
    }

    #[test]
    fn crypt_derive_key_sha1_expands_for_aes256() {
        // SHA-1("hello") = aaf4c61ddcc5e8a2dabede0f3b482cd9aea9434d
        let hash_value: [u8; 20] = [
            0xAA, 0xF4, 0xC6, 0x1D, 0xDC, 0xC5, 0xE8, 0xA2, 0xDA, 0xBE, 0xDE, 0x0F, 0x3B,
            0x48, 0x2C, 0xD9, 0xAE, 0xA9, 0x43, 0x4D,
        ];

        // Expected bytes computed independently with Python:
        //   buf = hash_value || 0x00*(64-hash_len)
        //   key = sha1(buf^0x36) || sha1(buf^0x5c)
        let expected: [u8; 32] = [
            0xB1, 0xBF, 0x85, 0x34, 0x6E, 0xCA, 0xE4, 0x29, 0xC0, 0xB3, 0x50, 0x63, 0x5B,
            0xAA, 0x3F, 0x25, 0x32, 0x13, 0x59, 0x82, 0xC2, 0xBF, 0x71, 0x1E, 0x09, 0x13,
            0x4D, 0x00, 0x1E, 0xBB, 0x01, 0x2F,
        ];

        let key = crypt_derive_key(&hash_value, 32, HashAlg::Sha1);
        assert_eq!(key, expected);
    }

    #[test]
    fn crypt_derive_key_md5_applies_cryptderivekey_even_when_key_len_le_hash_len() {
        // MD5("hello") = 5d41402abc4b2a76b9719d911017c592
        let hash_value: [u8; 16] = [
            0x5D, 0x41, 0x40, 0x2A, 0xBC, 0x4B, 0x2A, 0x76, 0xB9, 0x71, 0x9D, 0x91, 0x10,
            0x17, 0xC5, 0x92,
        ];

        // Expected bytes computed independently with Python:
        //   buf = hash_value || 0x00*(64-hash_len)
        //   key = md5(buf^0x36) || md5(buf^0x5c)
        let key = crypt_derive_key(&hash_value, 16, HashAlg::Md5);
        // MD5 uses the CryptoAPI `CryptDeriveKey` expansion even when `key_len == hash_len`:
        // it returns the first half of `md5(buf^0x36) || md5(buf^0x5c)`, i.e. `md5(buf^0x36)`.
        let expected: [u8; 16] = [
            0x21, 0xA4, 0xF9, 0x3F, 0x30, 0xEF, 0x88, 0x60, 0x3B, 0x66, 0x15, 0x32, 0x4E,
            0x70, 0x90, 0x1B,
        ];
        assert_eq!(key, expected);
    }

    #[test]
    fn crypt_derive_key_md5_expands_for_aes256() {
        // MD5("hello") = 5d41402abc4b2a76b9719d911017c592
        let hash_value: [u8; 16] = [
            0x5D, 0x41, 0x40, 0x2A, 0xBC, 0x4B, 0x2A, 0x76, 0xB9, 0x71, 0x9D, 0x91, 0x10,
            0x17, 0xC5, 0x92,
        ];

        // Expected bytes computed independently with Python:
        //   buf = hash_value || 0x00*(64-hash_len)
        //   key = md5(buf^0x36) || md5(buf^0x5c)
        let expected: [u8; 32] = [
            0x21, 0xA4, 0xF9, 0x3F, 0x30, 0xEF, 0x88, 0x60, 0x3B, 0x66, 0x15, 0x32, 0x4E,
            0x70, 0x90, 0x1B, 0x47, 0xE2, 0xBB, 0x9D, 0x88, 0xB0, 0x9C, 0x98, 0xE4, 0x8C,
            0x25, 0xE3, 0x68, 0xAD, 0x45, 0x9E,
        ];

        let key = crypt_derive_key(&hash_value, 32, HashAlg::Md5);
        assert_eq!(key, expected);
    }

    #[test]
    fn hash_password_fixed_spin_is_deterministic_sha1() {
        // Test vector computed independently with Python.
        // password = "Pässwörd"
        // password_utf16le = 5000e400730073007700f60072006400
        // salt = 000102030405060708090a0b0c0d0e0f
        // H = sha1(salt || password_utf16le)
        // for i in range(50000): H = sha1(le32(i) || H)
        let password_utf16le = password_to_utf16le("Pässwörd");
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C,
            0x0D, 0x0E, 0x0F,
        ];

        let h = hash_password_fixed_spin(&password_utf16le, &salt, HashAlg::Sha1);
        let expected: [u8; 20] = [
            0x38, 0x0E, 0xEE, 0x94, 0xF0, 0x45, 0x4D, 0x44, 0xE1, 0x75, 0x85, 0x46, 0x57,
            0x1B, 0xEB, 0x9B, 0xE5, 0xE5, 0x38, 0x7C,
        ];

        assert_eq!(h, expected);
    }
}
