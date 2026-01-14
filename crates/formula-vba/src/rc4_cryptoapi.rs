#![allow(dead_code)]

use sha1::{Digest as _, Sha1};
use thiserror::Error;

/// Standard/CryptoAPI RC4 block size used for OOXML `EncryptedPackage` encryption.
///
/// This is **not** the same as the BIFF8 RC4 block size (0x400).
const ENCRYPTED_PACKAGE_BLOCK_SIZE: usize = 0x200;

/// MS-OFFCRYPTO specifies 50,000 iterations for the legacy CryptoAPI hashing spin count.
const CRYPTOAPI_SPIN_COUNT: u32 = 50_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum HashAlg {
    Sha1,
}

#[derive(Debug, Error)]
pub(crate) enum Rc4CryptoApiError {
    #[error("EncryptedPackage stream is truncated")]
    Truncated,
    #[error("EncryptedPackage declared plaintext length {0} does not fit into usize")]
    PlaintextLenTooLarge(u64),
    #[error("unsupported key length {0} (must be <= 20 for SHA-1)")]
    UnsupportedKeyLength(usize),
    #[error("unsupported hash algorithm")]
    UnsupportedHashAlg,
}

/// RC4 keystream generator (KSA + PRGA).
#[derive(Clone)]
struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Rc4 {
    fn new(key: &[u8]) -> Self {
        debug_assert!(!key.is_empty(), "RC4 key must be non-empty");
        let mut s = [0u8; 256];
        for (i, v) in s.iter_mut().enumerate() {
            *v = i as u8;
        }
        let mut j: u8 = 0;
        for i in 0..256usize {
            j = j.wrapping_add(s[i]).wrapping_add(key[i % key.len()]);
            s.swap(i, j as usize);
        }
        Self { s, i: 0, j: 0 }
    }

    fn apply_keystream(&mut self, data: &mut [u8]) {
        for b in data {
            self.i = self.i.wrapping_add(1);
            self.j = self.j.wrapping_add(self.s[self.i as usize]);
            self.s.swap(self.i as usize, self.j as usize);
            let t = self.s[self.i as usize].wrapping_add(self.s[self.j as usize]);
            let k = self.s[t as usize];
            *b ^= k;
        }
    }
}

fn password_utf16le(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for ch in password.encode_utf16() {
        out.extend_from_slice(&ch.to_le_bytes());
    }
    out
}

fn sha1_digest(chunks: &[&[u8]]) -> [u8; 20] {
    let mut h = Sha1::new();
    for chunk in chunks {
        h.update(chunk);
    }
    h.finalize().into()
}

fn cryptoapi_password_hash(password: &str, salt: &[u8]) -> [u8; 20] {
    let pw = password_utf16le(password);
    // H0 = SHA1(salt + password_utf16le)
    let mut h = sha1_digest(&[salt, &pw]);
    // Hi = SHA1(i_le + H(i-1)), for i in [0, spin_count)
    for i in 0..CRYPTOAPI_SPIN_COUNT {
        h = sha1_digest(&[&i.to_le_bytes(), &h]);
    }
    h
}

/// Decryptor for MS-OFFCRYPTO "Standard" / CryptoAPI RC4 per-block encryption of the OOXML
/// `EncryptedPackage` stream.
///
/// This decryptor operates directly on the `EncryptedPackage` stream bytes:
///
/// ```text
/// u64le(plaintext_len) || ciphertext
/// ```
///
/// The plaintext is encrypted in 0x200-byte blocks:
/// - For each block index `i`, derive an RC4 key from the CryptoAPI password hash + `i`.
/// - Reset the RC4 keystream for each block (re-initialize RC4).
pub(crate) struct Rc4CryptoApiDecryptor {
    password_hash: [u8; 20],
    key_len: usize,
    hash_alg: HashAlg,
}

impl Rc4CryptoApiDecryptor {
    pub(crate) fn new(
        password: &str,
        salt: &[u8],
        key_len: usize,
        hash_alg: HashAlg,
    ) -> Result<Self, Rc4CryptoApiError> {
        if key_len > 20 {
            // SHA-1 digest size.
            return Err(Rc4CryptoApiError::UnsupportedKeyLength(key_len));
        }
        let password_hash = match hash_alg {
            HashAlg::Sha1 => cryptoapi_password_hash(password, salt),
        };
        Ok(Self {
            password_hash,
            key_len,
            hash_alg,
        })
    }

    fn derive_key(&self, block_index: u32) -> Result<Vec<u8>, Rc4CryptoApiError> {
        match self.hash_alg {
            HashAlg::Sha1 => {
                let digest = sha1_digest(&[&self.password_hash, &block_index.to_le_bytes()]);
                Ok(digest[..self.key_len].to_vec())
            }
        }
    }

    pub(crate) fn decrypt_encrypted_package(
        &self,
        encrypted_package_stream: &[u8],
    ) -> Result<Vec<u8>, Rc4CryptoApiError> {
        let len_bytes: [u8; 8] = encrypted_package_stream
            .get(..8)
            .ok_or(Rc4CryptoApiError::Truncated)?
            .try_into()
            .map_err(|_| Rc4CryptoApiError::Truncated)?;
        // The `EncryptedPackage` stream begins with an 8-byte **plaintext** size header.
        //
        // MS-OFFCRYPTO describes this prefix as a `u64le`, but some producers/libraries treat it as
        // `(u32 totalSize, u32 reserved)` (often with `reserved = 0`).
        //
        // For compatibility, when the high DWORD is non-zero and the combined 64-bit value is not
        // plausible for the available ciphertext, fall back to the low DWORD **only when it is
        // non-zero** (so we don't misinterpret true 64-bit sizes that are exact multiples of 2^32).
        let lo = u32::from_le_bytes([len_bytes[0], len_bytes[1], len_bytes[2], len_bytes[3]]);
        let hi = u32::from_le_bytes([len_bytes[4], len_bytes[5], len_bytes[6], len_bytes[7]]);
        let plaintext_len_u64_raw = (lo as u64) | ((hi as u64) << 32);
        let ciphertext_len = encrypted_package_stream.len().saturating_sub(8) as u64;
        let plaintext_len_u64 =
            if lo != 0
                && hi != 0
                && plaintext_len_u64_raw > ciphertext_len
                && (lo as u64) <= ciphertext_len
            {
                lo as u64
            } else {
                plaintext_len_u64_raw
            };
        let plaintext_len = usize::try_from(plaintext_len_u64)
            .map_err(|_| Rc4CryptoApiError::PlaintextLenTooLarge(plaintext_len_u64))?;

        let ciphertext = encrypted_package_stream
            .get(8..)
            .ok_or(Rc4CryptoApiError::Truncated)?;
        if ciphertext.len() < plaintext_len {
            return Err(Rc4CryptoApiError::Truncated);
        }

        let mut out = ciphertext[..plaintext_len].to_vec();
        for (block_index, chunk) in out.chunks_mut(ENCRYPTED_PACKAGE_BLOCK_SIZE).enumerate() {
            let key = self.derive_key(block_index as u32)?;
            let mut rc4 = Rc4::new(&key);
            rc4.apply_keystream(chunk);
        }
        Ok(out)
    }
}

#[cfg(test)]
mod tests {
    use super::{sha1_digest, HashAlg, Rc4, Rc4CryptoApiDecryptor, CRYPTOAPI_SPIN_COUNT};

    const BLOCK_SIZE: usize = 0x200;

    struct TestCryptoApiRc4 {
        password_hash: [u8; 20],
        key_len: usize,
    }

    impl TestCryptoApiRc4 {
        fn new(password: &str, salt: &[u8], key_len: usize) -> Self {
            assert!(key_len <= 20);
            let pw_utf16: Vec<u8> = password
                .encode_utf16()
                .flat_map(|ch| ch.to_le_bytes())
                .collect();

            let mut h = sha1_digest(&[salt, &pw_utf16]);
            for i in 0..CRYPTOAPI_SPIN_COUNT {
                h = sha1_digest(&[&i.to_le_bytes(), &h]);
            }

            Self {
                password_hash: h,
                key_len,
            }
        }

        fn key_for_block(&self, block_index: u32) -> Vec<u8> {
            let digest = sha1_digest(&[&self.password_hash, &block_index.to_le_bytes()]);
            digest[..self.key_len].to_vec()
        }

        fn encrypt_encrypted_package(&self, plaintext: &[u8]) -> Vec<u8> {
            let mut ciphertext = plaintext.to_vec();
            for (block_index, chunk) in ciphertext.chunks_mut(BLOCK_SIZE).enumerate() {
                let key = self.key_for_block(block_index as u32);
                let mut rc4 = Rc4::new(&key);
                rc4.apply_keystream(chunk);
            }

            let mut out = Vec::with_capacity(8 + ciphertext.len());
            out.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
            out.extend_from_slice(&ciphertext);
            out
        }
    }

    #[test]
    fn encrypted_package_size_header_does_not_fall_back_when_low_dword_is_zero() {
        // Ensure we don't misinterpret a true 64-bit size like 4GiB (low=0, high=1) as a 0-byte
        // payload (which would incorrectly return Ok(empty)).
        let decryptor = Rc4CryptoApiDecryptor {
            password_hash: [0u8; 20],
            key_len: 5,
            hash_alg: HashAlg::Sha1,
        };

        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0u32.to_le_bytes()); // low DWORD
        bytes.extend_from_slice(&1u32.to_le_bytes()); // high DWORD

        let err = decryptor
            .decrypt_encrypted_package(&bytes)
            .expect_err("expected truncated ciphertext");
        assert!(matches!(err, super::Rc4CryptoApiError::Truncated));
    }

    fn make_plaintext_pattern(len: usize) -> Vec<u8> {
        let mut out = vec![0u8; len];
        for (i, b) in out.iter_mut().enumerate() {
            let i = i as u32;
            let x = i
                .wrapping_mul(31)
                .wrapping_add(i.rotate_left(13))
                .wrapping_add(0x9E37_79B9);
            *b = (x ^ (x >> 8) ^ (x >> 16) ^ (x >> 24)) as u8;
        }
        out
    }

    #[test]
    fn rc4_cryptoapi_encryptedpackage_block_boundary_regression() {
        // Fixed parameters; keep deterministic.
        let password = "correct horse battery staple";
        let salt: [u8; 16] = [
            0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x10, 0x32, 0x54, 0x76, 0x98, 0xBA,
            0xDC, 0xFE,
        ];
        for key_len in [5usize, 7, 16] {
            // Keep encryption helper independent of the production decryptor's key derivation.
            let helper = TestCryptoApiRc4::new(password, &salt, key_len);
            let decryptor = Rc4CryptoApiDecryptor::new(password, &salt, key_len, HashAlg::Sha1)
                .expect("decryptor");

            let lengths = [0usize, 1, 511, 512, 513, 1023, 1024, 1025, 10_000];
            for len in lengths {
                let plaintext = make_plaintext_pattern(len);
                let encrypted_package = helper.encrypt_encrypted_package(&plaintext);
                let decrypted = decryptor
                    .decrypt_encrypted_package(&encrypted_package)
                    .expect("decrypt");
                assert_eq!(
                    decrypted, plaintext,
                    "round-trip mismatch at len={len} (key_len={key_len})"
                );
            }
        }
    }

    #[test]
    fn rc4_cryptoapi_40_bit_key_is_5_bytes() {
        let password = "password";
        let salt: [u8; 16] = (0u8..16u8).collect::<Vec<_>>().try_into().unwrap();
        let key_len = 5; // 40-bit

        let decryptor =
            Rc4CryptoApiDecryptor::new(password, &salt, key_len, HashAlg::Sha1).expect("decryptor");
        let derived = decryptor.derive_key(0).expect("derive key");
        assert_eq!(derived.len(), 5);
        assert_eq!(derived, vec![0x6a, 0xd7, 0xde, 0xdf, 0x2d]);
    }

    #[test]
    fn rc4_cryptoapi_encryptedpackage_roundtrip_with_40_bit_key() {
        // Fixed parameters; keep deterministic.
        let password = "correct horse battery staple";
        let salt: [u8; 16] = [
            0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x10, 0x32, 0x54, 0x76, 0x98, 0xBA,
            0xDC, 0xFE,
        ];
        let key_len = 5; // 40-bit

        // Keep encryption helper independent of the production decryptor's key derivation.
        let helper = TestCryptoApiRc4::new(password, &salt, key_len);
        let decryptor =
            Rc4CryptoApiDecryptor::new(password, &salt, key_len, HashAlg::Sha1).expect("decryptor");

        // Ensure plaintext crosses a 0x200 boundary so we rekey.
        let plaintext = make_plaintext_pattern(10_000);
        let encrypted_package = helper.encrypt_encrypted_package(&plaintext);
        let decrypted = decryptor
            .decrypt_encrypted_package(&encrypted_package)
            .expect("decrypt");
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn rc4_cryptoapi_encryptedpackage_falls_back_to_low_dword_when_high_dword_is_reserved() {
        // Fixed parameters; keep deterministic.
        let password = "correct horse battery staple";
        let salt: [u8; 16] = [
            0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x10, 0x32, 0x54, 0x76, 0x98, 0xBA,
            0xDC, 0xFE,
        ];
        let key_len = 16; // 128-bit RC4 key

        let helper = TestCryptoApiRc4::new(password, &salt, key_len);
        let decryptor =
            Rc4CryptoApiDecryptor::new(password, &salt, key_len, HashAlg::Sha1).expect("decryptor");

        let plaintext = make_plaintext_pattern(1024);
        let mut encrypted_package = helper.encrypt_encrypted_package(&plaintext);

        // Mutate the size prefix high DWORD to a non-zero reserved value.
        encrypted_package[4..8].copy_from_slice(&1u32.to_le_bytes());

        let decrypted = decryptor
            .decrypt_encrypted_package(&encrypted_package)
            .expect("decrypt");
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn rc4_cryptoapi_encryptedpackage_does_not_fall_back_when_low_dword_is_zero() {
        let password = "correct horse battery staple";
        let salt: [u8; 16] = [
            0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x10, 0x32, 0x54, 0x76, 0x98, 0xBA,
            0xDC, 0xFE,
        ];
        let key_len = 16;

        let decryptor =
            Rc4CryptoApiDecryptor::new(password, &salt, key_len, HashAlg::Sha1).expect("decryptor");

        // True 64-bit size (2^32). With no ciphertext, this must be treated as truncated, not as an
        // empty package.
        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(&0u32.to_le_bytes()); // lo
        encrypted_package.extend_from_slice(&1u32.to_le_bytes()); // hi

        let err = decryptor
            .decrypt_encrypted_package(&encrypted_package)
            .expect_err("should be truncated");
        assert!(matches!(err, super::Rc4CryptoApiError::Truncated));
    }
}
