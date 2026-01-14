use crate::error::OfficeCryptoError;
use aes::cipher::{generic_array::GenericArray, BlockDecrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use cbc::Decryptor;
use cbc::Encryptor;
use cipher::block_padding::NoPadding;
use cipher::{BlockDecryptMut, BlockEncryptMut, KeyIvInit};
use md5::Md5;
use sha2::Digest;
use zeroize::Zeroizing;

const MAX_DIGEST_LEN: usize = 64; // SHA-512
const MAX_HASH_BLOCK_LEN: usize = 128; // SHA-384/SHA-512 block size

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HashAlgorithm {
    Md5,
    Sha1,
    Sha256,
    Sha384,
    Sha512,
}

impl HashAlgorithm {
    pub fn as_ooxml_name(&self) -> &'static str {
        match self {
            HashAlgorithm::Md5 => "MD5",
            HashAlgorithm::Sha1 => "SHA1",
            HashAlgorithm::Sha256 => "SHA256",
            HashAlgorithm::Sha384 => "SHA384",
            HashAlgorithm::Sha512 => "SHA512",
        }
    }

    pub fn digest_len(&self) -> usize {
        match self {
            HashAlgorithm::Md5 => 16,
            HashAlgorithm::Sha1 => 20,
            HashAlgorithm::Sha256 => 32,
            HashAlgorithm::Sha384 => 48,
            HashAlgorithm::Sha512 => 64,
        }
    }

    pub(crate) fn block_len(&self) -> usize {
        // Hash block sizes (bytes). MD5/SHA-1/SHA-256 use 64-byte blocks; SHA-384/512 use 128-byte.
        match self {
            HashAlgorithm::Md5 | HashAlgorithm::Sha1 | HashAlgorithm::Sha256 => 64,
            HashAlgorithm::Sha384 | HashAlgorithm::Sha512 => 128,
        }
    }

    fn digest_two_into(&self, a: &[u8], b: &[u8], out: &mut [u8]) {
        debug_assert!(
            out.len() >= self.digest_len(),
            "hash output buffer too small"
        );
        match self {
            HashAlgorithm::Md5 => {
                let mut hasher = Md5::new();
                hasher.update(a);
                hasher.update(b);
                out[..16].copy_from_slice(&hasher.finalize());
            }
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
            "MD5" | "MD-5" => Ok(HashAlgorithm::Md5),
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
            0x0000_8003 => Ok(HashAlgorithm::Md5),    // CALG_MD5
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
            HashAlgorithm::Md5 => {
                let mut hasher = Md5::new();
                hasher.update(data);
                hasher.finalize().to_vec()
            }
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
        HashAlgorithm::Md5 => {
            for i in 0..spin_count {
                let mut hasher = Md5::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..digest_len]);
                h_buf[..digest_len].copy_from_slice(&hasher.finalize());
            }
        }
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

fn normalize_key_material(bytes: &[u8], out_len: usize) -> Vec<u8> {
    if bytes.len() >= out_len {
        return bytes[..out_len].to_vec();
    }

    // MS-OFFCRYPTO `TruncateHash` expansion: append 0x36 bytes (matches `msoffcrypto-tool`).
    let mut out = vec![0x36u8; out_len];
    out[..bytes.len()].copy_from_slice(bytes);
    out
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

    Zeroizing::new(normalize_key_material(&digest[..digest_len], key_bytes))
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

    normalize_key_material(&digest[..digest_len], iv_len)
}

#[cfg(test)]
pub(crate) fn aes_ecb_encrypt_in_place(
    key: &[u8],
    buf: &mut [u8],
) -> Result<(), OfficeCryptoError> {
    use aes::cipher::{generic_array::GenericArray, BlockEncrypt};

    if buf.len() % 16 != 0 {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "AES-ECB plaintext length must be multiple of 16 (got {})",
            buf.len()
        )));
    }

    fn encrypt_with<C>(key: &[u8], buf: &mut [u8]) -> Result<(), OfficeCryptoError>
    where
        C: BlockEncrypt + aes::cipher::KeyInit,
    {
        let cipher = C::new_from_slice(key)
            .map_err(|_| OfficeCryptoError::InvalidFormat("invalid AES key".to_string()))?;
        for block in buf.chunks_mut(16) {
            cipher.encrypt_block(GenericArray::from_mut_slice(block));
        }
        Ok(())
    }

    match key.len() {
        16 => encrypt_with::<Aes128>(key, buf),
        24 => encrypt_with::<Aes192>(key, buf),
        32 => encrypt_with::<Aes256>(key, buf),
        other => Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported AES key length {other}"
        ))),
    }
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

pub(crate) fn aes_ecb_decrypt_in_place(
    key: &[u8],
    buf: &mut [u8],
) -> Result<(), OfficeCryptoError> {
    if buf.is_empty() {
        return Ok(());
    }
    if buf.len() % 16 != 0 {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "AES-ECB ciphertext length must be multiple of 16 (got {})",
            buf.len()
        )));
    }

    fn decrypt_with<C>(key: &[u8], buf: &mut [u8]) -> Result<(), OfficeCryptoError>
    where
        C: BlockDecrypt + KeyInit,
    {
        let cipher = C::new_from_slice(key)
            .map_err(|_| OfficeCryptoError::InvalidFormat("invalid AES key".to_string()))?;
        for block in buf.chunks_mut(16) {
            cipher.decrypt_block(GenericArray::from_mut_slice(block));
        }
        Ok(())
    }

    match key.len() {
        16 => decrypt_with::<Aes128>(key, buf),
        24 => decrypt_with::<Aes192>(key, buf),
        32 => decrypt_with::<Aes256>(key, buf),
        other => Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported AES key length {other}"
        ))),
    }
}

#[allow(dead_code)]
pub(crate) fn aes_ecb_decrypt(
    key: &[u8],
    ciphertext: &[u8],
) -> Result<Vec<u8>, OfficeCryptoError> {
    let mut buf = ciphertext.to_vec();
    aes_ecb_decrypt_in_place(key, &mut buf)?;
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
#[allow(dead_code)]
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
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum StandardKeyDerivation {
    /// AES-based Standard encryption uses CryptoAPI `CryptDeriveKey` semantics (ipad/opad expansion).
    Aes,
    /// RC4-based Standard encryption uses key truncation (`key = H_block[..key_len]`).
    Rc4,
}

pub(crate) struct StandardKeyDeriver {
    hash_alg: HashAlgorithm,
    key_bytes: usize,
    password_hash: Zeroizing<Vec<u8>>,
    derivation: StandardKeyDerivation,
}

impl StandardKeyDeriver {
    pub(crate) fn new(
        hash_alg: HashAlgorithm,
        key_bits: u32,
        salt: &[u8],
        password: &str,
        derivation: StandardKeyDerivation,
    ) -> Self {
        Self::new_with_spin_count(hash_alg, key_bits, salt, password, derivation, 50_000)
    }

    pub(crate) fn new_with_spin_count(
        hash_alg: HashAlgorithm,
        key_bits: u32,
        salt: &[u8],
        password: &str,
        derivation: StandardKeyDerivation,
        spin_count: u32,
    ) -> Self {
        let pw = password_to_utf16le(password);
        let password_hash = hash_password(hash_alg, salt, &pw, spin_count);
        let key_bytes = (key_bits as usize) / 8;
        Self {
            hash_alg,
            key_bytes,
            password_hash,
            derivation,
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

        match self.derivation {
            StandardKeyDerivation::Aes => crypt_derive_key_aes(self.hash_alg, &h, self.key_bytes),
            StandardKeyDerivation::Rc4 => {
                if self.key_bytes > h.len() {
                    return Err(OfficeCryptoError::UnsupportedEncryption(format!(
                        "requested RC4 key length {} exceeds hash output length {}",
                        self.key_bytes,
                        h.len()
                    )));
                }
                Ok(Zeroizing::new(h[..self.key_bytes].to_vec()))
            }
        }
    }
}

fn crypt_derive_key_aes(
    hash_alg: HashAlgorithm,
    hash: &[u8],
    key_len: usize,
) -> Result<Zeroizing<Vec<u8>>, OfficeCryptoError> {
    // MS-OFFCRYPTO Standard encryption uses CryptoAPI `CryptDeriveKey`-style expansion:
    //
    //   D = hash padded with zeros to the hash block size (64 or 128 bytes)
    //   inner = Hash(D XOR 0x36)
    //   outer = Hash(D XOR 0x5c)
    //   derived = inner || outer
    //   key = derived[0..key_len]
    //
    // Notes:
    // - AES uses this derivation even for AES-128 (key_len < digest_len).
    // - The output length is 2*digest_len, which is always >= 32 for the hashes we support.
    let digest_len = hash_alg.digest_len();
    let block_len = hash_alg.block_len();
    debug_assert!(block_len <= MAX_HASH_BLOCK_LEN);

    let mut buf1: Zeroizing<Vec<u8>> = Zeroizing::new(vec![0x36u8; block_len]);
    let mut buf2: Zeroizing<Vec<u8>> = Zeroizing::new(vec![0x5Cu8; block_len]);
    let take = digest_len.min(hash.len()).min(block_len);
    for i in 0..take {
        buf1[i] ^= hash[i];
        buf2[i] ^= hash[i];
    }

    let h1: Zeroizing<Vec<u8>> = Zeroizing::new(hash_alg.digest(&buf1));
    let h2: Zeroizing<Vec<u8>> = Zeroizing::new(hash_alg.digest(&buf2));

    let mut out: Zeroizing<Vec<u8>> = Zeroizing::new(Vec::with_capacity(h1.len() + h2.len()));
    out.extend_from_slice(&h1);
    out.extend_from_slice(&h2);
    if key_len > out.len() {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "requested key length {} exceeds CryptoAPI derivation output length {}",
            key_len,
            out.len()
        )));
    }
    out.truncate(key_len);
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::time::{Duration, Instant};

    #[test]
    fn md5_digest_len_is_16() {
        assert_eq!(HashAlgorithm::Md5.digest_len(), 16);
    }

    #[test]
    fn hash_password_md5_spin_10_matches_vector() {
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];
        let pw = password_to_utf16le("password");
        let h = hash_password(HashAlgorithm::Md5, &salt, &pw, 10);
        assert_eq!(
            h.as_slice(),
            &[
                0x2B, 0x39, 0xE1, 0x55, 0x98, 0x6F, 0x47, 0x22, 0x96, 0x14, 0xE2, 0xBA, 0xED,
                0x8F, 0xB6, 0x0A
            ],
            "hash_password MD5 output mismatch"
        );
    }

    #[test]
    fn standard_key_derivation_md5_matches_vector() {
        // Test vectors match `formula-xls`'s CryptoAPI key derivation tests.
        let password = "password";
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];

        let expected: &[(u32, [u8; 16])] = &[
            (
                0,
                [
                    0x69, 0xBA, 0xDC, 0xAE, 0x24, 0x48, 0x68, 0xE2, 0x09, 0xD4, 0xE0, 0x53,
                    0xCC, 0xD2, 0xA3, 0xBC,
                ],
            ),
            (
                1,
                [
                    0x6F, 0x4D, 0x50, 0x2A, 0xB3, 0x77, 0x00, 0xFF, 0xDA, 0xB5, 0x70, 0x41,
                    0x60, 0x45, 0x5B, 0x47,
                ],
            ),
            (
                2,
                [
                    0xAC, 0x69, 0x02, 0x2E, 0x39, 0x6C, 0x77, 0x50, 0x87, 0x21, 0x33, 0xF3,
                    0x7E, 0x2C, 0x7A, 0xFC,
                ],
            ),
            (
                3,
                [
                    0x1B, 0x05, 0x6E, 0x71, 0x18, 0xAB, 0x8D, 0x35, 0xE9, 0xD6, 0x7A, 0xDE,
                    0xE8, 0xB1, 0x11, 0x04,
                ],
            ),
        ];

        // These vectors are from the legacy CryptoAPI RC4 derivation used by classic XLS.
        let deriver = StandardKeyDeriver::new(
            HashAlgorithm::Md5,
            128,
            &salt,
            password,
            StandardKeyDerivation::Rc4,
        );
        for (block, expected_key) in expected {
            let key = deriver
                .derive_key_for_block(*block)
                .unwrap_or_else(|_| panic!("derive block key {block}"));
            assert_eq!(key.as_slice(), expected_key, "block={block}");
        }
    }

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
                    0x45, 0xa0, 0x1f, 0x64, 0x5f, 0xc3, 0x5b, 0x38, 0x35, 0x52, 0x54,
                    0x4b, 0x9b, 0xf5,
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

    #[test]
    fn normalize_key_material_pads_with_0x36() {
        assert_eq!(
            normalize_key_material(&[0xAA, 0xBB], 5),
            vec![0xAA, 0xBB, 0x36, 0x36, 0x36]
        );
    }

    #[test]
    fn normalize_key_material_truncates() {
        assert_eq!(
            normalize_key_material(&[0xAA, 0xBB, 0xCC], 2),
            vec![0xAA, 0xBB]
        );
    }

    #[test]
    fn derive_agile_key_pads_with_0x36_when_longer_than_digest() {
        let salt = [0x11u8; 16];
        let pw_utf16 = password_to_utf16le("pw");
        let block_key = [0x22u8; 8];

        let key = derive_agile_key(HashAlgorithm::Sha1, &salt, &pw_utf16, 0, 24, &block_key);
        assert_eq!(key.len(), 24);
        assert_eq!(&key[20..], &[0x36u8; 4]);
    }

    #[test]
    fn derive_iv_pads_with_0x36_when_longer_than_digest() {
        let salt = [0x11u8; 16];
        let block_key = [0x22u8; 8];

        let iv = derive_iv(HashAlgorithm::Sha1, &salt, &block_key, 24);
        assert_eq!(iv.len(), 24);
        assert_eq!(&iv[20..], &[0x36u8; 4]);
    }
}
