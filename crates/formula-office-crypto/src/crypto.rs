use crate::error::OfficeCryptoError;
use aes::{Aes128, Aes192, Aes256};
use cbc::Decryptor;
use cbc::Encryptor;
use cipher::block_padding::NoPadding;
use cipher::{BlockDecryptMut, BlockEncryptMut, KeyIvInit};
use sha2::Digest;
use zeroize::Zeroizing;

const MAX_DIGEST_LEN: usize = 64; // SHA-512

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HashAlgorithm {
    Sha1,
    Sha256,
    Sha384,
    Sha512,
}

impl HashAlgorithm {
    pub fn as_ooxml_name(&self) -> &'static str {
        match self {
            HashAlgorithm::Sha1 => "SHA1",
            HashAlgorithm::Sha256 => "SHA256",
            HashAlgorithm::Sha384 => "SHA384",
            HashAlgorithm::Sha512 => "SHA512",
        }
    }

    pub fn digest_len(&self) -> usize {
        match self {
            HashAlgorithm::Sha1 => 20,
            HashAlgorithm::Sha256 => 32,
            HashAlgorithm::Sha384 => 48,
            HashAlgorithm::Sha512 => 64,
        }
    }

    fn digest_two_into(&self, a: &[u8], b: &[u8], out: &mut [u8]) {
        debug_assert!(
            out.len() >= self.digest_len(),
            "hash output buffer too small"
        );
        match self {
            HashAlgorithm::Sha1 => {
                let mut hasher = sha1::Sha1::new();
                hasher.update(a);
                hasher.update(b);
                out[..20].copy_from_slice(&hasher.finalize());
            }
            HashAlgorithm::Sha256 => {
                let mut hasher = sha2::Sha256::new();
                hasher.update(a);
                hasher.update(b);
                out[..32].copy_from_slice(&hasher.finalize());
            }
            HashAlgorithm::Sha384 => {
                let mut hasher = sha2::Sha384::new();
                hasher.update(a);
                hasher.update(b);
                out[..48].copy_from_slice(&hasher.finalize());
            }
            HashAlgorithm::Sha512 => {
                let mut hasher = sha2::Sha512::new();
                hasher.update(a);
                hasher.update(b);
                out[..64].copy_from_slice(&hasher.finalize());
            }
        }
    }

    pub(crate) fn from_name(name: &str) -> Result<Self, OfficeCryptoError> {
        match name {
            "SHA1" | "SHA-1" => Ok(HashAlgorithm::Sha1),
            "SHA256" | "SHA-256" => Ok(HashAlgorithm::Sha256),
            "SHA384" | "SHA-384" => Ok(HashAlgorithm::Sha384),
            "SHA512" | "SHA-512" => Ok(HashAlgorithm::Sha512),
            other => Err(OfficeCryptoError::UnsupportedEncryption(format!(
                "unsupported hash algorithm {other}"
            ))),
        }
    }

    pub(crate) fn from_cryptoapi_alg_id_hash(alg_id: u32) -> Result<Self, OfficeCryptoError> {
        // https://learn.microsoft.com/en-us/windows/win32/seccrypto/alg-id
        match alg_id {
            0x0000_8004 => Ok(HashAlgorithm::Sha1),   // CALG_SHA1
            0x0000_800C => Ok(HashAlgorithm::Sha256), // CALG_SHA_256
            0x0000_800D => Ok(HashAlgorithm::Sha384), // CALG_SHA_384
            0x0000_800E => Ok(HashAlgorithm::Sha512), // CALG_SHA_512
            other => Err(OfficeCryptoError::UnsupportedEncryption(format!(
                "unsupported hash AlgIDHash {other:#x}"
            ))),
        }
    }

    pub(crate) fn digest(&self, data: &[u8]) -> Vec<u8> {
        match self {
            HashAlgorithm::Sha1 => {
                let mut hasher = sha1::Sha1::new();
                hasher.update(data);
                hasher.finalize().to_vec()
            }
            HashAlgorithm::Sha256 => {
                let mut hasher = sha2::Sha256::new();
                hasher.update(data);
                hasher.finalize().to_vec()
            }
            HashAlgorithm::Sha384 => {
                let mut hasher = sha2::Sha384::new();
                hasher.update(data);
                hasher.finalize().to_vec()
            }
            HashAlgorithm::Sha512 => {
                let mut hasher = sha2::Sha512::new();
                hasher.update(data);
                hasher.finalize().to_vec()
            }
        }
    }
}

pub(crate) fn password_to_utf16le(password: &str) -> Zeroizing<Vec<u8>> {
    let mut out = Vec::with_capacity(password.len() * 2);
    for cu in password.encode_utf16() {
        out.extend_from_slice(&cu.to_le_bytes());
    }
    Zeroizing::new(out)
}

pub(crate) fn hash_password(
    hash_alg: HashAlgorithm,
    salt: &[u8],
    password_utf16le: &[u8],
    spin_count: u32,
) -> Zeroizing<Vec<u8>> {
    let digest_len = hash_alg.digest_len();
    debug_assert!(digest_len <= MAX_DIGEST_LEN);

    // Avoid per-iteration allocations (spinCount is often 50k-100k):
    // keep the current digest in a fixed buffer and overwrite it each round.
    let mut h_buf: Zeroizing<[u8; MAX_DIGEST_LEN]> = Zeroizing::new([0u8; MAX_DIGEST_LEN]);
    hash_alg.digest_two_into(salt, password_utf16le, &mut h_buf[..digest_len]);

    match hash_alg {
        HashAlgorithm::Sha1 => {
            for i in 0..spin_count {
                let mut hasher = sha1::Sha1::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..digest_len]);
                h_buf[..digest_len].copy_from_slice(&hasher.finalize());
            }
        }
        HashAlgorithm::Sha256 => {
            for i in 0..spin_count {
                let mut hasher = sha2::Sha256::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..digest_len]);
                h_buf[..digest_len].copy_from_slice(&hasher.finalize());
            }
        }
        HashAlgorithm::Sha384 => {
            for i in 0..spin_count {
                let mut hasher = sha2::Sha384::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..digest_len]);
                h_buf[..digest_len].copy_from_slice(&hasher.finalize());
            }
        }
        HashAlgorithm::Sha512 => {
            for i in 0..spin_count {
                let mut hasher = sha2::Sha512::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..digest_len]);
                h_buf[..digest_len].copy_from_slice(&hasher.finalize());
            }
        }
    }

    let mut out = vec![0u8; digest_len];
    out.copy_from_slice(&h_buf[..digest_len]);
    Zeroizing::new(out)
}

pub(crate) fn derive_agile_key(
    hash_alg: HashAlgorithm,
    salt: &[u8],
    password_utf16le: &[u8],
    spin_count: u32,
    key_bytes: usize,
    block_key: &[u8],
) -> Zeroizing<Vec<u8>> {
    let h = hash_password(hash_alg, salt, password_utf16le, spin_count);

    // Avoid allocating a temporary `H || blockKey` buffer: hash with two updates into a stack
    // buffer, then truncate/pad.
    let digest_len = hash_alg.digest_len();
    debug_assert!(digest_len <= MAX_DIGEST_LEN);
    let mut digest: Zeroizing<[u8; MAX_DIGEST_LEN]> = Zeroizing::new([0u8; MAX_DIGEST_LEN]);
    hash_alg.digest_two_into(&h, block_key, &mut digest[..digest_len]);

    let mut key: Zeroizing<Vec<u8>> = Zeroizing::new(vec![0u8; key_bytes]);
    let take = key_bytes.min(digest_len);
    key[..take].copy_from_slice(&digest[..take]);
    key
}

pub(crate) fn derive_iv(
    hash_alg: HashAlgorithm,
    salt: &[u8],
    block_key: &[u8],
    iv_len: usize,
) -> Vec<u8> {
    // Avoid allocating a temporary `salt || blockKey` buffer: hash with two updates into a stack
    // buffer, then truncate.
    let digest_len = hash_alg.digest_len();
    debug_assert!(digest_len <= MAX_DIGEST_LEN);
    let mut digest = [0u8; MAX_DIGEST_LEN];
    hash_alg.digest_two_into(salt, block_key, &mut digest[..digest_len]);

    let out_len = iv_len.min(digest_len);
    digest[..out_len].to_vec()
}

pub(crate) fn aes_cbc_decrypt(
    key: &[u8],
    iv: &[u8],
    ciphertext: &[u8],
) -> Result<Vec<u8>, OfficeCryptoError> {
    if iv.len() != 16 {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "AES-CBC IV must be 16 bytes (got {})",
            iv.len()
        )));
    }
    if ciphertext.len() % 16 != 0 {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "AES-CBC ciphertext length must be multiple of 16 (got {})",
            ciphertext.len()
        )));
    }
    let mut buf = ciphertext.to_vec();
    match key.len() {
        16 => {
            let dec = Decryptor::<Aes128>::new_from_slices(key, iv).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid AES-128 key/iv".to_string())
            })?;
            dec.decrypt_padded_mut::<NoPadding>(&mut buf).map_err(|_| {
                OfficeCryptoError::InvalidFormat("AES-CBC decrypt failed".to_string())
            })?;
        }
        24 => {
            let dec = Decryptor::<Aes192>::new_from_slices(key, iv).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid AES-192 key/iv".to_string())
            })?;
            dec.decrypt_padded_mut::<NoPadding>(&mut buf).map_err(|_| {
                OfficeCryptoError::InvalidFormat("AES-CBC decrypt failed".to_string())
            })?;
        }
        32 => {
            let dec = Decryptor::<Aes256>::new_from_slices(key, iv).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid AES-256 key/iv".to_string())
            })?;
            dec.decrypt_padded_mut::<NoPadding>(&mut buf).map_err(|_| {
                OfficeCryptoError::InvalidFormat("AES-CBC decrypt failed".to_string())
            })?;
        }
        other => {
            return Err(OfficeCryptoError::UnsupportedEncryption(format!(
                "unsupported AES key length {other}"
            )))
        }
    }
    Ok(buf)
}

pub(crate) fn aes_cbc_encrypt(
    key: &[u8],
    iv: &[u8],
    plaintext: &[u8],
) -> Result<Vec<u8>, OfficeCryptoError> {
    if iv.len() != 16 {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "AES-CBC IV must be 16 bytes (got {})",
            iv.len()
        )));
    }
    if plaintext.len() % 16 != 0 {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "AES-CBC plaintext length must be multiple of 16 (got {})",
            plaintext.len()
        )));
    }
    let mut buf = plaintext.to_vec();
    match key.len() {
        16 => {
            let enc = Encryptor::<Aes128>::new_from_slices(key, iv).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid AES-128 key/iv".to_string())
            })?;
            enc.encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .map_err(|_| {
                    OfficeCryptoError::InvalidFormat("AES-CBC encrypt failed".to_string())
                })?;
        }
        24 => {
            let enc = Encryptor::<Aes192>::new_from_slices(key, iv).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid AES-192 key/iv".to_string())
            })?;
            enc.encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .map_err(|_| {
                    OfficeCryptoError::InvalidFormat("AES-CBC encrypt failed".to_string())
                })?;
        }
        32 => {
            let enc = Encryptor::<Aes256>::new_from_slices(key, iv).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid AES-256 key/iv".to_string())
            })?;
            enc.encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .map_err(|_| {
                    OfficeCryptoError::InvalidFormat("AES-CBC encrypt failed".to_string())
                })?;
        }
        other => {
            return Err(OfficeCryptoError::UnsupportedEncryption(format!(
                "unsupported AES key length {other}"
            )))
        }
    }
    Ok(buf)
}

/// Apply the RC4 keystream to `data` in-place using `key`.
///
/// RC4 encryption and decryption are the same operation: `ciphertext = plaintext XOR keystream`.
pub(crate) fn rc4_xor_in_place(key: &[u8], data: &mut [u8]) -> Result<(), OfficeCryptoError> {
    use rc4::cipher::{KeyInit, StreamCipher};
    use rc4::Rc4;

    // `rc4` uses a type-level key size, so we dispatch the key sizes used by Office (40-bit/56-bit/
    // 128-bit) plus the short keys used by canonical test vectors.
    use rc4::cipher::consts::{U16, U3, U4, U5, U6, U7};

    match key.len() {
        3 => {
            let mut cipher = Rc4::<U3>::new_from_slice(key)
                .map_err(|_| OfficeCryptoError::UnsupportedEncryption("invalid RC4 key".to_string()))?;
            cipher.apply_keystream(data);
        }
        4 => {
            let mut cipher = Rc4::<U4>::new_from_slice(key)
                .map_err(|_| OfficeCryptoError::UnsupportedEncryption("invalid RC4 key".to_string()))?;
            cipher.apply_keystream(data);
        }
        5 => {
            let mut cipher = Rc4::<U5>::new_from_slice(key)
                .map_err(|_| OfficeCryptoError::UnsupportedEncryption("invalid RC4 key".to_string()))?;
            cipher.apply_keystream(data);
        }
        6 => {
            let mut cipher = Rc4::<U6>::new_from_slice(key)
                .map_err(|_| OfficeCryptoError::UnsupportedEncryption("invalid RC4 key".to_string()))?;
            cipher.apply_keystream(data);
        }
        7 => {
            let mut cipher = Rc4::<U7>::new_from_slice(key)
                .map_err(|_| OfficeCryptoError::UnsupportedEncryption("invalid RC4 key".to_string()))?;
            cipher.apply_keystream(data);
        }
        16 => {
            let mut cipher = Rc4::<U16>::new_from_slice(key)
                .map_err(|_| OfficeCryptoError::UnsupportedEncryption("invalid RC4 key".to_string()))?;
            cipher.apply_keystream(data);
        }
        other => {
            return Err(OfficeCryptoError::UnsupportedEncryption(format!(
                "unsupported RC4 key length {other}"
            )))
        }
    }

    Ok(())
}

/// Implements the MS-OFFCRYPTO "Standard Encryption" password/key derivation that mimics
/// CryptoAPI's `CryptDeriveKey`.
///
/// We keep this in a dedicated struct so we can reuse the expensive password hash across blocks.
pub(crate) struct StandardKeyDeriver {
    hash_alg: HashAlgorithm,
    key_bytes: usize,
    password_hash: Zeroizing<Vec<u8>>,
}

impl StandardKeyDeriver {
    pub(crate) fn new(hash_alg: HashAlgorithm, key_bits: u32, salt: &[u8], password: &str) -> Self {
        let pw = password_to_utf16le(password);
        // Office Standard encryption uses a fixed spin count of 50k.
        let password_hash = hash_password(hash_alg, salt, &pw, 50_000);
        let key_bytes = (key_bits as usize) / 8;
        Self {
            hash_alg,
            key_bytes,
            password_hash,
        }
    }

    pub(crate) fn derive_key_for_block(
        &self,
        block_index: u32,
    ) -> Result<Zeroizing<Vec<u8>>, OfficeCryptoError> {
        let mut buf: Zeroizing<Vec<u8>> =
            Zeroizing::new(Vec::with_capacity(self.password_hash.len() + 4));
        buf.extend_from_slice(&self.password_hash);
        buf.extend_from_slice(&block_index.to_le_bytes());
        let h: Zeroizing<Vec<u8>> = Zeroizing::new(self.hash_alg.digest(&buf));
        Ok(crypt_derive_key(self.hash_alg, &h, self.key_bytes))
    }
}

fn crypt_derive_key(hash_alg: HashAlgorithm, hash: &[u8], key_len: usize) -> Zeroizing<Vec<u8>> {
    if key_len <= hash.len() {
        return Zeroizing::new(hash[..key_len].to_vec());
    }

    // MS-OFFCRYPTO's CryptoAPI key derivation extension: hash padded with 0x36/0x5c to 64 bytes,
    // then hashed again to produce additional material.
    let mut buf1: Zeroizing<Vec<u8>> = Zeroizing::new(Vec::with_capacity(64));
    buf1.extend_from_slice(hash);
    buf1.resize(64, 0x36);

    let mut buf2: Zeroizing<Vec<u8>> = Zeroizing::new(Vec::with_capacity(64));
    buf2.extend_from_slice(hash);
    buf2.resize(64, 0x5C);

    let h1: Zeroizing<Vec<u8>> = Zeroizing::new(hash_alg.digest(&buf1));
    let h2: Zeroizing<Vec<u8>> = Zeroizing::new(hash_alg.digest(&buf2));

    let mut out: Zeroizing<Vec<u8>> = Zeroizing::new(Vec::with_capacity(h1.len() + h2.len()));
    out.extend_from_slice(&h1);
    out.extend_from_slice(&h2);
    out.truncate(key_len);
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::time::{Duration, Instant};

    #[test]
    fn hash_password_perf_guard_spin_10k() {
        // Regression guard: the spinCount loop is the hot path for both Standard (50k) and Agile
        // (often 100k) password-based encryption.
        let salt = [0x11u8; 16];
        let pw = password_to_utf16le("password");

        let start = Instant::now();
        let _ = hash_password(HashAlgorithm::Sha256, &salt, &pw, 10_000);
        assert!(
            start.elapsed() < Duration::from_secs(2),
            "hash_password(spinCount=10_000) took too long: {:?}",
            start.elapsed()
        );
    }

    #[test]
    fn rc4_vectors_encrypt_decrypt_symmetry() {
        // Canonical raw RC4 test vectors (no drop):
        // - https://en.wikipedia.org/wiki/RC4#Test_vectors
        let cases: &[(&[u8], &[u8], &[u8])] = &[
            (
                b"Key",
                b"Plaintext",
                &[0xbb, 0xf3, 0x16, 0xe8, 0xd9, 0x40, 0xaf, 0x0a, 0xd3],
            ),
            (b"Wiki", b"pedia", &[0x10, 0x21, 0xbf, 0x04, 0x20]),
            (
                b"Secret",
                b"Attack at dawn",
                &[
                    0x45, 0xa0, 0x1f, 0x64, 0x5f, 0xc3, 0x5b, 0x38, 0x35, 0x52, 0x54, 0x4b,
                    0x9b, 0xf5,
                ],
            ),
        ];

        for (key, plaintext, expected_ciphertext) in cases {
            let mut ciphertext = plaintext.to_vec();
            rc4_xor_in_place(key, &mut ciphertext).expect("RC4 encrypt");
            assert_eq!(
                ciphertext.as_slice(),
                *expected_ciphertext,
                "encrypt key={:?} plaintext={:?}",
                std::str::from_utf8(key).ok(),
                std::str::from_utf8(plaintext).ok()
            );

            rc4_xor_in_place(key, &mut ciphertext).expect("RC4 decrypt");
            assert_eq!(
                ciphertext.as_slice(),
                *plaintext,
                "decrypt key={:?}",
                std::str::from_utf8(key).ok()
            );
        }
    }
}
