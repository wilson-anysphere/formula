//! MS-OFFCRYPTO **Standard (CryptoAPI)** password verification for OOXML `EncryptedPackage`
//! containers.
//!
//! Office "Encrypt with Password" for OOXML files stores the real `.xlsx` ZIP package inside an OLE
//! compound file with:
//! - `EncryptionInfo`: encryption parameters + verifier
//! - `EncryptedPackage`: encrypted ZIP bytes
//!
//! This module implements parsing the **binary** `EncryptionInfo` payload for Standard encryption
//! (`EncryptionVersionInfo.versionMinor == 2`, commonly 3.2 but major 2/3/4 is observed) and
//! verifying a candidate password by decrypting the `EncryptionVerifier` fields and comparing the
//! verifier hash.
//!
//! Scope: password verification only (not full package decryption).

use aes::cipher::generic_array::GenericArray;
use aes::cipher::{BlockDecrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use md5::Md5;
use sha1::Sha1;

use formula_xlsx::offcrypto::{decrypt_aes_cbc_no_padding_in_place, AesCbcDecryptError, AES_BLOCK_SIZE};

/// CALG constants used by MS-OFFCRYPTO Standard (CryptoAPI) encryption.
///
/// Values are from `wincrypt.h` and MS-OFFCRYPTO.
pub const CALG_RC4: u32 = 0x0000_6801;
pub const CALG_AES_128: u32 = 0x0000_660E;
pub const CALG_AES_192: u32 = 0x0000_660F;
pub const CALG_AES_256: u32 = 0x0000_6610;
pub const CALG_SHA1: u32 = 0x0000_8004;
pub const CALG_MD5: u32 = 0x0000_8003;

// Standard EncryptionInfo is identified by `versionMinor == 2`; `versionMajor` varies in the wild.
//
// Keep the canonical (commonly-observed) major version constant for tests/fixtures.
#[cfg(test)]
const STANDARD_MAJOR_VERSION: u16 = 3;
const STANDARD_MINOR_VERSION: u16 = 2;
// MS-OFFCRYPTO identifies Standard (CryptoAPI) encryption via `versionMinor == 2`, but real-world
// files vary `versionMajor` across Office generations.
fn is_standard_cryptoapi_version(major: u16, minor: u16) -> bool {
    minor == STANDARD_MINOR_VERSION && matches!(major, 2 | 3 | 4)
}
const ENCRYPTION_HEADER_FIXED_LEN: usize = 8 * 4;
/// MS-OFFCRYPTO Standard uses a fixed spin count of 50,000 iterations for password hashing.
const STANDARD_SPIN_COUNT: u32 = 50_000;

#[derive(Debug, thiserror::Error)]
pub enum OffcryptoError {
    #[error(
        "unsupported EncryptionInfo version {major}.{minor} (expected Standard CryptoAPI versionMinor=2 with major=2/3/4)"
    )]
    UnsupportedEncryptionInfoVersion {
        major: u16,
        minor: u16,
        #[allow(dead_code)]
        expected_major: u16,
        expected_minor: u16,
    },

    #[error("truncated EncryptionInfo stream while reading {context}: needed {needed} bytes, only {available} available")]
    Truncated {
        context: &'static str,
        needed: usize,
        available: usize,
    },

    #[error("invalid EncryptionHeader size {header_size}: must be at least {min_size}")]
    InvalidHeaderSize { header_size: u32, min_size: usize },

    #[error("invalid CSPName length {len}: must be even (UTF-16LE)")]
    InvalidCspNameLength { len: usize },

    #[error("unsupported external Standard encryption (fExternal flag set)")]
    UnsupportedExternalEncryption,

    #[error("unsupported Standard encryption: fCryptoAPI flag not set")]
    UnsupportedNonCryptoApiStandardEncryption,

    #[error("invalid Standard EncryptionHeader flags for algId: flags={flags:#010x}, alg_id={alg_id:#010x}")]
    InvalidFlags { flags: u32, alg_id: u32 },

    #[error("invalid salt size {salt_size}: exceeds available bytes")]
    InvalidSaltSize { salt_size: u32 },

    #[error("invalid verifierHashSize {verifier_hash_size}")]
    InvalidVerifierHashSize { verifier_hash_size: u32 },

    #[error("unsupported encryption algorithm algId={alg_id:#010x}")]
    UnsupportedAlgId { alg_id: u32 },

    #[error("unsupported hash algorithm algIdHash={alg_id_hash:#010x}")]
    UnsupportedAlgIdHash { alg_id_hash: u32 },

    #[error("invalid key size {key_size_bits} bits")]
    InvalidKeySize { key_size_bits: u32 },

    #[error("AES ciphertext length {len} is not a multiple of the block size ({AES_BLOCK_SIZE})")]
    InvalidAesCiphertextLength { len: usize },

    #[error("cryptographic error: {message}")]
    Crypto { message: String },
}

impl OffcryptoError {
    fn crypto(message: impl Into<String>) -> Self {
        Self::Crypto {
            message: message.into(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StandardEncryptionInfo {
    pub header: EncryptionHeader,
    pub verifier: EncryptionVerifier,
}

/// Parsed `EncryptionHeader.Flags` bits for MS-OFFCRYPTO Standard (CryptoAPI) encryption.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncryptionHeaderFlags {
    pub raw: u32,
    pub f_cryptoapi: bool,
    pub f_doc_props: bool,
    pub f_external: bool,
    pub f_aes: bool,
}

impl EncryptionHeaderFlags {
    pub const F_CRYPTOAPI: u32 = 0x0000_0004;
    pub const F_DOCPROPS: u32 = 0x0000_0008;
    pub const F_EXTERNAL: u32 = 0x0000_0010;
    pub const F_AES: u32 = 0x0000_0020;

    pub fn from_raw(raw: u32) -> Self {
        Self {
            raw,
            f_cryptoapi: raw & Self::F_CRYPTOAPI != 0,
            f_doc_props: raw & Self::F_DOCPROPS != 0,
            f_external: raw & Self::F_EXTERNAL != 0,
            f_aes: raw & Self::F_AES != 0,
        }
    }
}

/// MS-OFFCRYPTO `EncryptionHeader` for Standard encryption.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EncryptionHeader {
    pub flags: EncryptionHeaderFlags,
    pub size_extra: u32,
    pub alg_id: u32,
    pub alg_id_hash: u32,
    /// Key size in *bits*.
    pub key_size: u32,
    pub provider_type: u32,
    pub reserved1: u32,
    pub reserved2: u32,
    pub csp_name: String,
}

/// MS-OFFCRYPTO `EncryptionVerifier` for Standard encryption.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EncryptionVerifier {
    pub salt: Vec<u8>,
    pub encrypted_verifier: [u8; 16],
    pub verifier_hash_size: u32,
    /// The remainder of the `EncryptionInfo` stream after `verifierHashSize`.
    ///
    /// For AES this is typically padded to a multiple of 16 bytes; for RC4 it is usually exactly
    /// `verifierHashSize` bytes.
    pub encrypted_verifier_hash: Vec<u8>,
}

struct Reader<'a> {
    bytes: &'a [u8],
    pos: usize,
}

impl<'a> Reader<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, pos: 0 }
    }

    fn remaining(&self) -> usize {
        self.bytes.len().saturating_sub(self.pos)
    }

    fn read_u16_le(&mut self, context: &'static str) -> Result<u16, OffcryptoError> {
        let needed = 2;
        if self.remaining() < needed {
            return Err(OffcryptoError::Truncated {
                context,
                needed,
                available: self.remaining(),
            });
        }
        let b = &self.bytes[self.pos..self.pos + 2];
        self.pos += 2;
        Ok(u16::from_le_bytes([b[0], b[1]]))
    }

    fn read_u32_le(&mut self, context: &'static str) -> Result<u32, OffcryptoError> {
        let needed = 4;
        if self.remaining() < needed {
            return Err(OffcryptoError::Truncated {
                context,
                needed,
                available: self.remaining(),
            });
        }
        let b = &self.bytes[self.pos..self.pos + 4];
        self.pos += 4;
        Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
    }

    fn read_bytes(&mut self, len: usize, context: &'static str) -> Result<&'a [u8], OffcryptoError> {
        if self.remaining() < len {
            return Err(OffcryptoError::Truncated {
                context,
                needed: len,
                available: self.remaining(),
            });
        }
        let out = &self.bytes[self.pos..self.pos + len];
        self.pos += len;
        Ok(out)
    }

    fn read_array<const N: usize>(
        &mut self,
        context: &'static str,
    ) -> Result<[u8; N], OffcryptoError> {
        let b = self.read_bytes(N, context)?;
        let mut out = [0u8; N];
        out.copy_from_slice(b);
        Ok(out)
    }
}

/// Parse a Standard (CryptoAPI) `EncryptionInfo` stream (`versionMinor == 2`, major 2/3/4).
pub fn parse_encryption_info_standard(bytes: &[u8]) -> Result<StandardEncryptionInfo, OffcryptoError> {
    let mut r = Reader::new(bytes);
    let major = r.read_u16_le("majorVersion")?;
    let minor = r.read_u16_le("minorVersion")?;
    if !is_standard_cryptoapi_version(major, minor) {
        return Err(OffcryptoError::UnsupportedEncryptionInfoVersion {
            major,
            minor,
            expected_major: 3,
            expected_minor: STANDARD_MINOR_VERSION,
        });
    }

    // MS-OFFCRYPTO `EncryptionInfo.Flags` (distinct from `EncryptionHeader.flags`).
    //
    // We currently do not interpret these bits for Standard encryption, but they are part of the
    // on-disk stream layout.
    let _flags = r.read_u32_le("flags")?;

    let header_size_u32 = r.read_u32_le("headerSize")?;
    if header_size_u32 < ENCRYPTION_HEADER_FIXED_LEN as u32 {
        return Err(OffcryptoError::InvalidHeaderSize {
            header_size: header_size_u32,
            min_size: ENCRYPTION_HEADER_FIXED_LEN,
        });
    }
    let header_size = header_size_u32 as usize;
    let header_bytes = r.read_bytes(header_size, "EncryptionHeader")?;
    let header = parse_encryption_header(header_bytes)?;

    let verifier_bytes = r.read_bytes(r.remaining(), "EncryptionVerifier")?;
    let verifier = parse_encryption_verifier(verifier_bytes)?;

    validate_parsed_standard_encryption_info(&header, &verifier)?;

    Ok(StandardEncryptionInfo { header, verifier })
}

fn validate_parsed_standard_encryption_info(
    header: &EncryptionHeader,
    verifier: &EncryptionVerifier,
) -> Result<(), OffcryptoError> {
    // Ensure we only return `Ok(_)` for structures that `verify_password_standard` can evaluate
    // without returning an error. This keeps the API ergonomic for callers and avoids
    // "successfully parsed but unverifiable" states.

    // Validate supported algorithms.
    match header.alg_id {
        CALG_RC4 | CALG_AES_128 | CALG_AES_192 | CALG_AES_256 => {}
        other => return Err(OffcryptoError::UnsupportedAlgId { alg_id: other }),
    }
    let hash_len = match header.alg_id_hash {
        CALG_SHA1 => 20usize,
        CALG_MD5 => 16usize,
        other => return Err(OffcryptoError::UnsupportedAlgIdHash { alg_id_hash: other }),
    };

    // Validate key size semantics against the downstream `CryptDeriveKey` implementation.
    let key_size_bits = header.key_size;
    if key_size_bits == 0 || key_size_bits % 8 != 0 {
        return Err(OffcryptoError::InvalidKeySize { key_size_bits });
    }
    let key_len = (key_size_bits / 8) as usize;
    if key_len > hash_len.saturating_mul(2) {
        return Err(OffcryptoError::InvalidKeySize { key_size_bits });
    }

    // AES requires an exact key size.
    match header.alg_id {
        CALG_AES_128 if key_size_bits != 128 => {
            return Err(OffcryptoError::InvalidKeySize { key_size_bits })
        }
        CALG_AES_192 if key_size_bits != 192 => {
            return Err(OffcryptoError::InvalidKeySize { key_size_bits })
        }
        CALG_AES_256 if key_size_bits != 256 => {
            return Err(OffcryptoError::InvalidKeySize { key_size_bits })
        }
        _ => {}
    }

    // Validate verifier hash sizing. The encrypted blob may include padding (AES), but the
    // declared verifierHashSize must fit in the decrypted bytes and be no longer than the hash
    // output.
    let verifier_hash_size = verifier.verifier_hash_size as usize;
    if verifier_hash_size == 0 || verifier_hash_size > hash_len {
        return Err(OffcryptoError::InvalidVerifierHashSize {
            verifier_hash_size: verifier.verifier_hash_size,
        });
    }
    if verifier_hash_size > verifier.encrypted_verifier_hash.len() {
        return Err(OffcryptoError::InvalidVerifierHashSize {
            verifier_hash_size: verifier.verifier_hash_size,
        });
    }

    // AES verifier material is encrypted as a single CBC stream and thus must be block-aligned.
    if matches!(header.alg_id, CALG_AES_128 | CALG_AES_192 | CALG_AES_256) {
        let ciphertext_len = 16 + verifier.encrypted_verifier_hash.len();
        if ciphertext_len % AES_BLOCK_SIZE != 0 {
            return Err(OffcryptoError::InvalidAesCiphertextLength { len: ciphertext_len });
        }
    }

    Ok(())
}

fn parse_encryption_header(bytes: &[u8]) -> Result<EncryptionHeader, OffcryptoError> {
    let mut r = Reader::new(bytes);

    let flags_raw = r.read_u32_le("EncryptionHeader.flags")?;
    let flags = EncryptionHeaderFlags::from_raw(flags_raw);
    let size_extra = r.read_u32_le("EncryptionHeader.sizeExtra")?;
    let alg_id = r.read_u32_le("EncryptionHeader.algId")?;
    let alg_id_hash = r.read_u32_le("EncryptionHeader.algIdHash")?;
    let key_size = r.read_u32_le("EncryptionHeader.keySize")?;
    let provider_type = r.read_u32_le("EncryptionHeader.providerType")?;
    let reserved1 = r.read_u32_le("EncryptionHeader.reserved1")?;
    let reserved2 = r.read_u32_le("EncryptionHeader.reserved2")?;

    // Validate `EncryptionHeader.flags` semantics. This is conservative and reduces false positives
    // when parsing arbitrary OLE streams as Standard encryption.
    if flags.f_external {
        return Err(OffcryptoError::UnsupportedExternalEncryption);
    }
    if !flags.f_cryptoapi {
        return Err(OffcryptoError::UnsupportedNonCryptoApiStandardEncryption);
    }
    let alg_is_aes = matches!(alg_id, CALG_AES_128 | CALG_AES_192 | CALG_AES_256);
    if flags.f_aes != alg_is_aes {
        return Err(OffcryptoError::InvalidFlags {
            flags: flags_raw,
            alg_id,
        });
    }

    let csp_name_bytes = r.read_bytes(r.remaining(), "EncryptionHeader.CSPName")?;
    if csp_name_bytes.len() % 2 != 0 {
        return Err(OffcryptoError::InvalidCspNameLength {
            len: csp_name_bytes.len(),
        });
    }
    let mut utf16: Vec<u16> = csp_name_bytes
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect();
    while utf16.last() == Some(&0) {
        utf16.pop();
    }
    let csp_name = String::from_utf16_lossy(&utf16);

    Ok(EncryptionHeader {
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

fn parse_encryption_verifier(bytes: &[u8]) -> Result<EncryptionVerifier, OffcryptoError> {
    let mut r = Reader::new(bytes);

    let salt_size_u32 = r.read_u32_le("EncryptionVerifier.saltSize")?;
    let salt_size = salt_size_u32 as usize;
    if r.remaining() < salt_size {
        return Err(OffcryptoError::InvalidSaltSize {
            salt_size: salt_size_u32,
        });
    }
    let salt = r.read_bytes(salt_size, "EncryptionVerifier.salt")?.to_vec();

    let encrypted_verifier = r.read_array::<16>("EncryptionVerifier.encryptedVerifier")?;
    let verifier_hash_size = r.read_u32_le("EncryptionVerifier.verifierHashSize")?;
    let encrypted_verifier_hash = r
        .read_bytes(r.remaining(), "EncryptionVerifier.encryptedVerifierHash")?
        .to_vec();

    Ok(EncryptionVerifier {
        salt,
        encrypted_verifier,
        verifier_hash_size,
        encrypted_verifier_hash,
    })
}

/// Verify a password against the Standard (CryptoAPI) `EncryptionVerifier` structure.
///
/// Returns `Ok(true)` when the password is correct, `Ok(false)` when the password is incorrect, and
/// `Err(_)` for malformed inputs or unsupported algorithms.
pub fn verify_password_standard(
    info: &StandardEncryptionInfo,
    password: &str,
) -> Result<bool, OffcryptoError> {
    let key = derive_file_key_standard(info, password)?;

    // Decrypt the concatenated verifier blob (`encryptedVerifier` || `encryptedVerifierHash`) as a
    // single stream.
    let mut ciphertext = Vec::with_capacity(16 + info.verifier.encrypted_verifier_hash.len());
    ciphertext.extend_from_slice(&info.verifier.encrypted_verifier);
    ciphertext.extend_from_slice(&info.verifier.encrypted_verifier_hash);

    match info.header.alg_id {
        CALG_AES_128 | CALG_AES_192 | CALG_AES_256 => {
            if ciphertext.len() % AES_BLOCK_SIZE != 0 {
                return Err(OffcryptoError::InvalidAesCiphertextLength {
                    len: ciphertext.len(),
                });
            }

            // Baseline MS-OFFCRYPTO Standard AES uses AES-ECB (no IV) for verifier fields.
            // However, some producers use CBC-style variants; fall back to the derived-IV CBC
            // scheme if ECB does not verify.
            let mut ecb_plaintext = ciphertext.clone();
            aes_ecb_decrypt_in_place(&key, &mut ecb_plaintext)?;
            if verifier_hash_matches(info, &ecb_plaintext)? {
                return Ok(true);
            }

            // Compatibility fallback: AES-CBC (no padding) with a derived IV.
            let iv = derive_standard_aes_iv(info)?;
            decrypt_aes_cbc_no_padding_in_place(&key, &iv, &mut ciphertext).map_err(|err| {
                let msg = match err {
                    AesCbcDecryptError::UnsupportedKeyLength(_) => "unsupported AES key length",
                    AesCbcDecryptError::InvalidIvLength(_) => "invalid AES IV length",
                    AesCbcDecryptError::InvalidCiphertextLength(_) => "invalid AES ciphertext length",
                };
                OffcryptoError::crypto(msg)
            })?;
            verifier_hash_matches(info, &ciphertext)
        }
        CALG_RC4 => {
            rc4_apply_keystream(&key, &mut ciphertext)?;
            verifier_hash_matches(info, &ciphertext)
        }
        other => Err(OffcryptoError::UnsupportedAlgId { alg_id: other }),
    }
}

fn verifier_hash_matches(
    info: &StandardEncryptionInfo,
    plaintext: &[u8],
) -> Result<bool, OffcryptoError> {
    let verifier_hash_size = info.verifier.verifier_hash_size as usize;
    if verifier_hash_size == 0 {
        return Err(OffcryptoError::InvalidVerifierHashSize {
            verifier_hash_size: info.verifier.verifier_hash_size,
        });
    }
    if plaintext.len() < 16 + verifier_hash_size {
        return Err(OffcryptoError::Truncated {
            context: "decrypted verifier blob",
            needed: 16 + verifier_hash_size,
            available: plaintext.len(),
        });
    }

    let verifier = &plaintext[0..16];
    let verifier_hash = &plaintext[16..16 + verifier_hash_size];

    let expected_full = hash(info.header.alg_id_hash, &[verifier])?;
    if verifier_hash_size > expected_full.len() {
        return Err(OffcryptoError::InvalidVerifierHashSize {
            verifier_hash_size: info.verifier.verifier_hash_size,
        });
    }
    Ok(expected_full[0..verifier_hash_size] == *verifier_hash)
}

fn aes_ecb_decrypt_in_place(key: &[u8], buf: &mut [u8]) -> Result<(), OffcryptoError> {
    if buf.len() % AES_BLOCK_SIZE != 0 {
        return Err(OffcryptoError::InvalidAesCiphertextLength { len: buf.len() });
    }

    fn decrypt_with<C>(key: &[u8], buf: &mut [u8]) -> Result<(), OffcryptoError>
    where
        C: BlockDecrypt + KeyInit,
    {
        let cipher = C::new_from_slice(key).map_err(|_| OffcryptoError::crypto("unsupported AES key length"))?;
        for block in buf.chunks_mut(AES_BLOCK_SIZE) {
            cipher.decrypt_block(GenericArray::from_mut_slice(block));
        }
        Ok(())
    }

    match key.len() {
        16 => decrypt_with::<Aes128>(key, buf),
        24 => decrypt_with::<Aes192>(key, buf),
        32 => decrypt_with::<Aes256>(key, buf),
        other => Err(OffcryptoError::crypto(format!(
            "unsupported AES key length {other} bytes"
        ))),
    }
}

fn derive_file_key_standard(
    info: &StandardEncryptionInfo,
    password: &str,
) -> Result<Vec<u8>, OffcryptoError> {
    let key_size_bits = info.header.key_size;
    if key_size_bits == 0 || key_size_bits % 8 != 0 {
        return Err(OffcryptoError::InvalidKeySize { key_size_bits });
    }
    let key_len = (key_size_bits / 8) as usize;

    let password_utf16le = utf16le_bytes(password);
    let h = hash_password_fixed_spin(&password_utf16le, &info.verifier.salt, info.header.alg_id_hash)?;

    let block = 0u32.to_le_bytes();
    let h_block0 = hash(info.header.alg_id_hash, &[&h, &block])?;

    match info.header.alg_id {
        CALG_RC4 => {
            if key_len > h_block0.len() {
                return Err(OffcryptoError::InvalidKeySize { key_size_bits });
            }
            // Standard RC4 key derivation truncates H_block0 directly. CryptoAPI/Office represent a
            // "40-bit" RC4 key as a padded 16-byte key (high 88 bits zero).
            let mut key = h_block0[..key_len].to_vec();
            if key_len == 5 {
                key.resize(16, 0);
            }
            Ok(key)
        }
        CALG_AES_128 | CALG_AES_192 | CALG_AES_256 => {
            crypt_derive_key(&h_block0, key_len, info.header.alg_id_hash)
        }
        other => Err(OffcryptoError::UnsupportedAlgId { alg_id: other }),
    }
}

fn derive_standard_aes_iv(info: &StandardEncryptionInfo) -> Result<[u8; AES_BLOCK_SIZE], OffcryptoError> {
    // MS-OFFCRYPTO Standard AES IV derivation for verifier checks:
    ////
    //   iv = Hash(salt || LE32(0))[..16]
    let block = 0u32.to_le_bytes();
    let iv_full = hash(info.header.alg_id_hash, &[&info.verifier.salt, &block])?;
    if iv_full.len() < AES_BLOCK_SIZE {
        return Err(OffcryptoError::crypto(format!(
            "hash output too short for AES IV: got {} bytes",
            iv_full.len()
        )));
    }
    let mut iv = [0u8; AES_BLOCK_SIZE];
    iv.copy_from_slice(&iv_full[..AES_BLOCK_SIZE]);
    Ok(iv)
}

fn utf16le_bytes(s: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(s.len().saturating_mul(2));
    for ch in s.encode_utf16() {
        out.extend_from_slice(&ch.to_le_bytes());
    }
    out
}

fn hash_password_fixed_spin(
    password_utf16le: &[u8],
    salt: &[u8],
    alg_id_hash: u32,
) -> Result<Vec<u8>, OffcryptoError> {
    // H0 = Hash(salt || password)
    let mut h = hash(alg_id_hash, &[salt, password_utf16le])?;

    // Hi = Hash(LE32(i) || H(i-1)), for i = 0..49999 (50,000 iterations).
    for i in 0..STANDARD_SPIN_COUNT {
        let i_le = (i as u32).to_le_bytes();
        h = hash(alg_id_hash, &[&i_le, &h])?;
    }

    Ok(h)
}

fn crypt_derive_key(
    hash_value: &[u8],
    key_len: usize,
    alg_id_hash: u32,
) -> Result<Vec<u8>, OffcryptoError> {
    if key_len == 0 {
        return Err(OffcryptoError::InvalidKeySize { key_size_bits: 0 });
    }

    // CryptoAPI `CryptDeriveKey` (MD5/SHA1): expand with an ipad/opad construction and truncate.
    //
    // This is **not** "truncate the hash": for Standard AES, Office uses this expansion even when
    // the requested key length is smaller than the digest length.
    //
    // Given H (hash_value), build a 64-byte buffer: H || 0x00.., then:
    //   X1 = Hash((buffer XOR 0x36..))
    //   X2 = Hash((buffer XOR 0x5C..))
    //   key = (X1 || X2)[..key_len]
    let hash_len = hash_value.len();
    let max_len = hash_len.checked_mul(2).unwrap_or(0);
    if key_len > max_len {
        return Err(OffcryptoError::InvalidKeySize {
            key_size_bits: (key_len as u32).saturating_mul(8),
        });
    }

    let mut buf = [0u8; 64];
    let copy_len = core::cmp::min(hash_len, buf.len());
    buf[..copy_len].copy_from_slice(&hash_value[..copy_len]);

    let mut ipad = [0x36u8; 64];
    let mut opad = [0x5Cu8; 64];
    for i in 0..64 {
        ipad[i] ^= buf[i];
        opad[i] ^= buf[i];
    }

    let x1 = hash(alg_id_hash, &[&ipad])?;
    let x2 = hash(alg_id_hash, &[&opad])?;
    let mut out = Vec::with_capacity(x1.len() + x2.len());
    out.extend_from_slice(&x1);
    out.extend_from_slice(&x2);
    out.truncate(key_len);
    Ok(out)
}

fn hash(alg_id_hash: u32, parts: &[&[u8]]) -> Result<Vec<u8>, OffcryptoError> {
    match alg_id_hash {
        CALG_SHA1 => {
            use sha1::Digest as _;
            let mut h = Sha1::new();
            for p in parts {
                h.update(p);
            }
            Ok(h.finalize().to_vec())
        }
        CALG_MD5 => {
            use md5::Digest as _;
            let mut h = Md5::new();
            for p in parts {
                h.update(p);
            }
            Ok(h.finalize().to_vec())
        }
        other => Err(OffcryptoError::UnsupportedAlgIdHash {
            alg_id_hash: other,
        }),
    }
}

fn rc4_apply_keystream(key: &[u8], buf: &mut [u8]) -> Result<(), OffcryptoError> {
    if key.is_empty() {
        return Err(OffcryptoError::crypto("invalid RC4 key length (empty)"));
    }

    let mut s = [0u8; 256];
    for (i, b) in s.iter_mut().enumerate() {
        *b = i as u8;
    }

    // Key-scheduling algorithm (KSA).
    let mut j: u8 = 0;
    for i in 0..256usize {
        j = j.wrapping_add(s[i]).wrapping_add(key[i % key.len()]);
        s.swap(i, j as usize);
    }

    // Pseudo-random generation algorithm (PRGA).
    let mut i: u8 = 0;
    j = 0;
    for b in buf.iter_mut() {
        i = i.wrapping_add(1);
        j = j.wrapping_add(s[i as usize]);
        s.swap(i as usize, j as usize);
        let t = s[i as usize].wrapping_add(s[j as usize]);
        let k = s[t as usize];
        *b ^= k;
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    use aes::Aes128;
    use cbc::cipher::{block_padding::Pkcs7, BlockEncryptMut, KeyIvInit};

    fn build_standard_encryption_info_bytes(
        header: &EncryptionHeader,
        verifier: &EncryptionVerifier,
    ) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&STANDARD_MAJOR_VERSION.to_le_bytes());
        out.extend_from_slice(&STANDARD_MINOR_VERSION.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes()); // EncryptionInfo.Flags

        let mut header_bytes = Vec::new();
        header_bytes.extend_from_slice(&header.flags.raw.to_le_bytes());
        header_bytes.extend_from_slice(&header.size_extra.to_le_bytes());
        header_bytes.extend_from_slice(&header.alg_id.to_le_bytes());
        header_bytes.extend_from_slice(&header.alg_id_hash.to_le_bytes());
        header_bytes.extend_from_slice(&header.key_size.to_le_bytes());
        header_bytes.extend_from_slice(&header.provider_type.to_le_bytes());
        header_bytes.extend_from_slice(&header.reserved1.to_le_bytes());
        header_bytes.extend_from_slice(&header.reserved2.to_le_bytes());
        header_bytes.extend_from_slice(&utf16le_bytes(&format!("{}\0", header.csp_name)));

        out.extend_from_slice(&(header_bytes.len() as u32).to_le_bytes());
        out.extend_from_slice(&header_bytes);

        out.extend_from_slice(&(verifier.salt.len() as u32).to_le_bytes());
        out.extend_from_slice(&verifier.salt);
        out.extend_from_slice(&verifier.encrypted_verifier);
        out.extend_from_slice(&verifier.verifier_hash_size.to_le_bytes());
        out.extend_from_slice(&verifier.encrypted_verifier_hash);
        out
    }

    #[test]
    fn verify_password_standard_aes_sha1() {
        // Test vector from `docs/offcrypto-standard-cryptoapi.md` (§8.2).
        let password = "Password1234_";
        let wrong_password = "Password1234!";
        let salt: [u8; 16] = [
            0xE8, 0x82, 0x66, 0x49, 0x0C, 0x5B, 0xD1, 0xEE, 0xBD, 0x2B, 0x43, 0x94, 0xE3,
            0xF8, 0x30, 0xEF,
        ];

        // Expected values computed from the MS-OFFCRYPTO Standard key/IV derivation.
        let expected_key: [u8; 16] = [
            0x40, 0xB1, 0x3A, 0x71, 0xF9, 0x0B, 0x96, 0x6E, 0x37, 0x54, 0x08, 0xF2, 0xD1,
            0x81, 0xA1, 0xAA,
        ];
        let expected_iv: [u8; 16] = [
            0xA1, 0xCD, 0xC2, 0x53, 0x36, 0x96, 0x4D, 0x31, 0x4D, 0xD9, 0x68, 0xDA, 0x99,
            0x8D, 0x05, 0xB8,
        ];

        let header = EncryptionHeader {
            flags: EncryptionHeaderFlags::from_raw(
                EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_AES,
            ),
            size_extra: 0,
            alg_id: CALG_AES_128,
            alg_id_hash: CALG_SHA1,
            key_size: 128,
            provider_type: 0,
            reserved1: 0,
            reserved2: 0,
            csp_name: "Microsoft Enhanced RSA and AES Cryptographic Provider".to_string(),
        };

        let verifier: [u8; 16] = [
            0x1F, 0xA2, 0x3B, 0x4C, 0x5D, 0x6E, 0x7F, 0x80, 0x91, 0xA0, 0xB1, 0xC2, 0xD3,
            0xE4, 0xF5, 0x06,
        ];
        let verifier_hash = hash(CALG_SHA1, &[&verifier]).expect("sha1 hash");
        assert_eq!(verifier_hash.len(), 20);

        // Encrypt verifier || verifier_hash as a single AES-CBC stream, with PKCS7 padding.
        let mut plaintext = Vec::new();
        plaintext.extend_from_slice(&verifier);
        plaintext.extend_from_slice(&verifier_hash);

        // Derive key and IV and assert they match the embedded expected constants.
        let password_utf16le = utf16le_bytes(password);
        let h = hash_password_fixed_spin(&password_utf16le, &salt, CALG_SHA1).unwrap();
        let block = 0u32.to_le_bytes();
        let h_final = hash(CALG_SHA1, &[&h, &block]).unwrap();
        let key = crypt_derive_key(&h_final, 16, CALG_SHA1).unwrap();
        assert_eq!(key.as_slice(), expected_key);

        let iv_full = hash(CALG_SHA1, &[&salt, &block]).unwrap();
        assert_eq!(&iv_full[..16], expected_iv);

        let mut buf = plaintext.clone();
        let pos = buf.len();
        buf.resize(pos + 16, 0);
        let ct = cbc::Encryptor::<Aes128>::new_from_slices(&key, expected_iv.as_slice())
            .unwrap()
            .encrypt_padded_mut::<Pkcs7>(&mut buf, pos)
            .unwrap();
        let ciphertext = ct.to_vec();
        assert!(ciphertext.len().is_multiple_of(16));

        let encrypted_verifier: [u8; 16] = ciphertext[0..16].try_into().unwrap();
        let encrypted_verifier_hash = ciphertext[16..].to_vec();

        let verifier_struct = EncryptionVerifier {
            salt: salt.to_vec(),
            encrypted_verifier,
            verifier_hash_size: verifier_hash.len() as u32,
            encrypted_verifier_hash,
        };

        let bytes = build_standard_encryption_info_bytes(&header, &verifier_struct);
        let parsed = parse_encryption_info_standard(&bytes).expect("parse");

        assert!(verify_password_standard(&parsed, password).unwrap());
        assert!(!verify_password_standard(&parsed, wrong_password).unwrap());
    }

    #[test]
    fn verify_password_standard_rc4_sha1() {
        let password = "hunter2";
        let wrong_password = "hunter3";
        let salt: [u8; 16] = [
            0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B, 0x1C,
            0x1D, 0x1E, 0x1F,
        ];

        let expected_key: [u8; 16] = [
            0x8F, 0x5C, 0x2B, 0x8A, 0xD0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00,
        ];

        let header = EncryptionHeader {
            flags: EncryptionHeaderFlags::from_raw(EncryptionHeaderFlags::F_CRYPTOAPI),
            size_extra: 0,
            alg_id: CALG_RC4,
            alg_id_hash: CALG_SHA1,
            key_size: 40,
            provider_type: 0,
            reserved1: 0,
            reserved2: 0,
            csp_name: "Microsoft Base Cryptographic Provider".to_string(),
        };

        let verifier: [u8; 16] = [
            0xAB, 0xCD, 0xEF, 0x01, 0x23, 0x45, 0x67, 0x89, 0x10, 0x32, 0x54, 0x76, 0x98,
            0xBA, 0xDC, 0xFE,
        ];
        let verifier_hash = hash(CALG_SHA1, &[&verifier]).expect("sha1 hash");

        let password_utf16le = utf16le_bytes(password);
        let h = hash_password_fixed_spin(&password_utf16le, &salt, CALG_SHA1).unwrap();
        let block = 0u32.to_le_bytes();
        let h_final = hash(CALG_SHA1, &[&h, &block]).unwrap();
        let mut key = h_final[..5].to_vec();
        key.resize(16, 0);
        assert_eq!(key.as_slice(), expected_key);

        // Encrypt verifier || verifier_hash using RC4 (symmetric).
        let mut plaintext = Vec::new();
        plaintext.extend_from_slice(&verifier);
        plaintext.extend_from_slice(&verifier_hash);

        let mut ciphertext = plaintext.clone();
        rc4_apply_keystream(&key, &mut ciphertext).unwrap();

        let encrypted_verifier: [u8; 16] = ciphertext[0..16].try_into().unwrap();
        let encrypted_verifier_hash = ciphertext[16..].to_vec();

        let verifier_struct = EncryptionVerifier {
            salt: salt.to_vec(),
            encrypted_verifier,
            verifier_hash_size: verifier_hash.len() as u32,
            encrypted_verifier_hash,
        };

        let bytes = build_standard_encryption_info_bytes(&header, &verifier_struct);
        let parsed = parse_encryption_info_standard(&bytes).expect("parse");

        assert!(verify_password_standard(&parsed, password).unwrap());
        assert!(!verify_password_standard(&parsed, wrong_password).unwrap());
    }

    fn minimal_encryption_info_header(flags: u32, alg_id: u32) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&STANDARD_MAJOR_VERSION.to_le_bytes());
        out.extend_from_slice(&STANDARD_MINOR_VERSION.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes()); // EncryptionInfo.Flags
        // header_size = fixed 8 DWORD header (no CSPName).
        out.extend_from_slice(&(ENCRYPTION_HEADER_FIXED_LEN as u32).to_le_bytes());

        // EncryptionHeader.
        out.extend_from_slice(&flags.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
        out.extend_from_slice(&alg_id.to_le_bytes()); // algId
        out.extend_from_slice(&CALG_SHA1.to_le_bytes()); // algIdHash
        out.extend_from_slice(&128u32.to_le_bytes()); // keySize
        out.extend_from_slice(&0u32.to_le_bytes()); // providerType
        out.extend_from_slice(&0u32.to_le_bytes()); // reserved1
        out.extend_from_slice(&0u32.to_le_bytes()); // reserved2

        out
    }

    #[test]
    fn rejects_external_standard_encryption_flag() {
        let flags =
            EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_EXTERNAL | EncryptionHeaderFlags::F_AES;
        let bytes = minimal_encryption_info_header(flags, CALG_AES_128);
        let err = parse_encryption_info_standard(&bytes).expect_err("expected error");
        assert!(matches!(err, OffcryptoError::UnsupportedExternalEncryption));
    }

    #[test]
    fn rejects_standard_without_cryptoapi_flag() {
        let flags = EncryptionHeaderFlags::F_AES;
        let bytes = minimal_encryption_info_header(flags, CALG_AES_128);
        let err = parse_encryption_info_standard(&bytes).expect_err("expected error");
        assert!(matches!(
            err,
            OffcryptoError::UnsupportedNonCryptoApiStandardEncryption
        ));
    }

    #[test]
    fn rejects_aes_algid_without_faes_flag() {
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI;
        let bytes = minimal_encryption_info_header(flags, CALG_AES_128);
        let err = parse_encryption_info_standard(&bytes).expect_err("expected error");
        assert!(matches!(
            err,
            OffcryptoError::InvalidFlags { flags: _, alg_id: _ }
        ));
    }

    #[test]
    fn rejects_faes_flag_with_non_aes_algid() {
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_AES;
        let bytes = minimal_encryption_info_header(flags, CALG_RC4);
        let err = parse_encryption_info_standard(&bytes).expect_err("expected error");
        assert!(matches!(
            err,
            OffcryptoError::InvalidFlags { flags: _, alg_id: _ }
        ));
    }
}
