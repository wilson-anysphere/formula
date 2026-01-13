use crate::error::OfficeCryptoError;
use aes::{Aes128, Aes192, Aes256};
use cbc::Decryptor;
use cbc::Encryptor;
use cipher::block_padding::NoPadding;
use cipher::{BlockDecryptMut, BlockEncryptMut, KeyIvInit};
use sha2::Digest;
use zeroize::Zeroizing;

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
    let mut initial = Vec::with_capacity(salt.len() + password_utf16le.len());
    initial.extend_from_slice(salt);
    initial.extend_from_slice(password_utf16le);

    let mut h: Zeroizing<Vec<u8>> = Zeroizing::new(hash_alg.digest(&initial));
    for i in 0..spin_count {
        let mut buf = Vec::with_capacity(4 + h.len());
        buf.extend_from_slice(&i.to_le_bytes());
        buf.extend_from_slice(&h);
        *h = hash_alg.digest(&buf);
    }
    h
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
    let mut buf = Vec::with_capacity(h.len() + block_key.len());
    buf.extend_from_slice(&h);
    buf.extend_from_slice(block_key);
    let mut key: Zeroizing<Vec<u8>> = Zeroizing::new(hash_alg.digest(&buf));
    key.truncate(key_bytes);
    if key.len() < key_bytes {
        key.resize(key_bytes, 0);
    }
    key
}

pub(crate) fn derive_iv(hash_alg: HashAlgorithm, salt: &[u8], block_key: &[u8], iv_len: usize) -> Vec<u8> {
    let mut buf = Vec::with_capacity(salt.len() + block_key.len());
    buf.extend_from_slice(salt);
    buf.extend_from_slice(block_key);
    let hash = hash_alg.digest(&buf);
    hash[..iv_len.min(hash.len())].to_vec()
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
                .map_err(|_| OfficeCryptoError::InvalidFormat("AES-CBC encrypt failed".to_string()))?;
        }
        24 => {
            let enc = Encryptor::<Aes192>::new_from_slices(key, iv).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid AES-192 key/iv".to_string())
            })?;
            enc.encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .map_err(|_| OfficeCryptoError::InvalidFormat("AES-CBC encrypt failed".to_string()))?;
        }
        32 => {
            let enc = Encryptor::<Aes256>::new_from_slices(key, iv).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid AES-256 key/iv".to_string())
            })?;
            enc.encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .map_err(|_| OfficeCryptoError::InvalidFormat("AES-CBC encrypt failed".to_string()))?;
        }
        other => {
            return Err(OfficeCryptoError::UnsupportedEncryption(format!(
                "unsupported AES key length {other}"
            )))
        }
    }
    Ok(buf)
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
        let mut buf = Vec::with_capacity(self.password_hash.len() + 4);
        buf.extend_from_slice(&self.password_hash);
        buf.extend_from_slice(&block_index.to_le_bytes());
        let h = self.hash_alg.digest(&buf);
        Ok(Zeroizing::new(crypt_derive_key(
            self.hash_alg,
            &h,
            self.key_bytes,
        )))
    }
}

fn crypt_derive_key(hash_alg: HashAlgorithm, hash: &[u8], key_len: usize) -> Vec<u8> {
    if key_len <= hash.len() {
        return hash[..key_len].to_vec();
    }

    // MS-OFFCRYPTO's CryptoAPI key derivation extension: hash padded with 0x36/0x5c to 64 bytes,
    // then hashed again to produce additional material.
    let mut buf1 = Vec::with_capacity(64);
    buf1.extend_from_slice(hash);
    buf1.resize(64, 0x36);

    let mut buf2 = Vec::with_capacity(64);
    buf2.extend_from_slice(hash);
    buf2.resize(64, 0x5C);

    let h1 = hash_alg.digest(&buf1);
    let h2 = hash_alg.digest(&buf2);

    let mut out = Vec::with_capacity(h1.len() + h2.len());
    out.extend_from_slice(&h1);
    out.extend_from_slice(&h2);
    out.truncate(key_len);
    out
}
