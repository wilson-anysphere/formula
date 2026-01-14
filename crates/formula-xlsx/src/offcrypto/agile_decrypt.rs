//! MS-OFFCRYPTO Agile decryption for OOXML `EncryptedPackage`.

use digest::Digest as _;
use hmac::{Hmac, Mac};
use std::io::{Read, Seek, SeekFrom, Write};

use super::aes_cbc::{
    decrypt_aes_cbc_no_padding, decrypt_aes_cbc_no_padding_in_place, AES_BLOCK_SIZE,
};
use super::agile::DecryptOptions;
use super::crypto::{
    derive_iv, derive_key, hash_password, segment_block_key, HashAlgorithm, HMAC_KEY_BLOCK,
    HMAC_VALUE_BLOCK, KEY_VALUE_BLOCK, VERIFIER_HASH_INPUT_BLOCK, VERIFIER_HASH_VALUE_BLOCK,
};
use super::encryption_info::{
    decode_base64_field_limited, decode_encryption_info_xml_text, extract_encryption_info_xml,
    ParseOptions,
};
use super::error::{OffCryptoError, Result};
use super::warning::OffCryptoWarning;

const SEGMENT_SIZE: usize = 0x1000;
const KEY_ENCRYPTOR_URI_PASSWORD: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
const KEY_ENCRYPTOR_URI_CERTIFICATE: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate";

#[derive(Debug, Clone)]
struct KeyData {
    salt_value: Vec<u8>,
    hash_algorithm: HashAlgorithm,
    block_size: usize,
    key_bits: usize,
    hash_size: usize,
}

#[derive(Debug, Clone)]
struct DataIntegrity {
    encrypted_hmac_key: Vec<u8>,
    encrypted_hmac_value: Vec<u8>,
}

#[derive(Debug, Clone)]
struct PasswordKeyEncryptor {
    salt_value: Vec<u8>,
    hash_algorithm: HashAlgorithm,
    spin_count: u32,
    block_size: usize,
    key_bits: usize,
    hash_size: usize,
    encrypted_verifier_hash_input: Vec<u8>,
    encrypted_verifier_hash_value: Vec<u8>,
    encrypted_key_value: Vec<u8>,
}

#[derive(Debug, Clone)]
struct AgileEncryptionInfo {
    key_data: KeyData,
    data_integrity: Option<DataIntegrity>,
    password_key: PasswordKeyEncryptor,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PasswordKeyIvDerivation {
    /// MS-OFFCRYPTO spec behavior: use `p:encryptedKey.saltValue` truncated to `blockSize`.
    SaltValue,
    /// Compatibility behavior observed in some producers: derive the IV using the standard Agile
    /// `derive_iv(saltValue, blockKey, blockSize, hashAlgorithm)` scheme.
    Derived,
}

fn decrypt_agile_package_key_from_password(
    info: &AgileEncryptionInfo,
    password_hash: &[u8],
    key_encrypt_key_len: usize,
    package_key_len: usize,
    iv_derivation: PasswordKeyIvDerivation,
) -> Result<Vec<u8>> {
    decrypt_agile_package_key_from_password_with_iv_derivation(
        info,
        password_hash,
        key_encrypt_key_len,
        package_key_len,
        iv_derivation,
    )
}

fn decrypt_agile_package_key_from_password_with_iv_derivation(
    info: &AgileEncryptionInfo,
    password_hash: &[u8],
    key_encrypt_key_len: usize,
    package_key_len: usize,
    iv_derivation: PasswordKeyIvDerivation,
) -> Result<Vec<u8>> {
    let password_key = &info.password_key;
    // MS-OFFCRYPTO: for the password key encryptor (`p:encryptedKey`), the AES-CBC IV used for
    // `encryptedVerifierHashInput`, `encryptedVerifierHashValue`, and `encryptedKeyValue` is the
    // password `saltValue` itself (truncated to blockSize). Some producers vary from this; the
    // caller can request `PasswordKeyIvDerivation::Derived` to use the `derive_iv` scheme instead.
    let iv_for_block = |block_key: &[u8]| -> Result<Vec<u8>> {
        match iv_derivation {
            PasswordKeyIvDerivation::SaltValue => Ok(password_key
                .salt_value
                .get(..password_key.block_size)
                .ok_or_else(|| OffCryptoError::InvalidAttribute {
                    element: "p:encryptedKey".to_string(),
                    attr: "saltValue".to_string(),
                    reason: "saltValue shorter than blockSize".to_string(),
                })?
                .to_vec()),
            PasswordKeyIvDerivation::Derived => derive_iv_or_err(
                &password_key.salt_value,
                block_key,
                password_key.block_size,
                password_key.hash_algorithm,
            ),
        }
    };

    // Decrypt verifierHashInput.
    let verifier_input = {
        let verifier_iv = iv_for_block(&VERIFIER_HASH_INPUT_BLOCK)?;
        let k = derive_key_or_err(
            password_hash,
            &VERIFIER_HASH_INPUT_BLOCK,
            key_encrypt_key_len,
            password_key.hash_algorithm,
        )?;
        let decrypted = decrypt_aes_cbc_no_padding(
            &k,
            &verifier_iv,
            &password_key.encrypted_verifier_hash_input,
        )
        .map_err(|e| OffCryptoError::InvalidAttribute {
            element: "p:encryptedKey".to_string(),
            attr: "encryptedVerifierHashInput".to_string(),
            reason: e.to_string(),
        })?;
        decrypted
            .get(..password_key.block_size)
            .ok_or_else(|| OffCryptoError::InvalidAttribute {
                element: "p:encryptedKey".to_string(),
                attr: "encryptedVerifierHashInput".to_string(),
                reason: "decrypted verifierHashInput shorter than blockSize".to_string(),
            })?
            .to_vec()
    };

    // Decrypt verifierHashValue.
    let verifier_hash = {
        let verifier_iv = iv_for_block(&VERIFIER_HASH_VALUE_BLOCK)?;
        let k = derive_key_or_err(
            password_hash,
            &VERIFIER_HASH_VALUE_BLOCK,
            key_encrypt_key_len,
            password_key.hash_algorithm,
        )?;
        let decrypted = decrypt_aes_cbc_no_padding(
            &k,
            &verifier_iv,
            &password_key.encrypted_verifier_hash_value,
        )
        .map_err(|e| OffCryptoError::InvalidAttribute {
            element: "p:encryptedKey".to_string(),
            attr: "encryptedVerifierHashValue".to_string(),
            reason: e.to_string(),
        })?;
        decrypted
            .get(..password_key.hash_size)
            .ok_or_else(|| OffCryptoError::InvalidAttribute {
                element: "p:encryptedKey".to_string(),
                attr: "encryptedVerifierHashValue".to_string(),
                reason: "decrypted verifierHashValue shorter than hashSize".to_string(),
            })?
            .to_vec()
    };

    // Verify password.
    let computed_verifier_hash_full = hash_bytes(password_key.hash_algorithm, &verifier_input);
    let computed_verifier_hash = computed_verifier_hash_full
        .get(..password_key.hash_size)
        .ok_or_else(|| OffCryptoError::InvalidAttribute {
            element: "p:encryptedKey".to_string(),
            attr: "hashAlgorithm".to_string(),
            reason: "hash output shorter than hashSize".to_string(),
        })?;
    if !ct_eq(computed_verifier_hash, &verifier_hash) {
        return Err(OffCryptoError::WrongPassword);
    }

    // Decrypt the package key (encryptedKeyValue).
    let key_value = {
        let key_value_iv = iv_for_block(&KEY_VALUE_BLOCK)?;
        let k = derive_key_or_err(
            password_hash,
            &KEY_VALUE_BLOCK,
            key_encrypt_key_len,
            password_key.hash_algorithm,
        )?;
        let decrypted =
            decrypt_aes_cbc_no_padding(&k, &key_value_iv, &password_key.encrypted_key_value)
                .map_err(|e| OffCryptoError::InvalidAttribute {
                    element: "p:encryptedKey".to_string(),
                    attr: "encryptedKeyValue".to_string(),
                    reason: e.to_string(),
                })?;
        decrypted
            .get(..package_key_len)
            .ok_or_else(|| OffCryptoError::InvalidAttribute {
                element: "p:encryptedKey".to_string(),
                attr: "encryptedKeyValue".to_string(),
                reason: "decrypted keyValue shorter than keyData.keyBits".to_string(),
            })?
            .to_vec()
    };

    Ok(key_value)
}

fn decrypt_agile_package_key_from_password_best_effort(
    info: &AgileEncryptionInfo,
    password_hash: &[u8],
    key_encrypt_key_len: usize,
    package_key_len: usize,
) -> Result<Vec<u8>> {
    match decrypt_agile_package_key_from_password(
        info,
        password_hash,
        key_encrypt_key_len,
        package_key_len,
        PasswordKeyIvDerivation::SaltValue,
    ) {
        Ok(key) => Ok(key),
        Err(OffCryptoError::WrongPassword) => decrypt_agile_package_key_from_password(
            info,
            password_hash,
            key_encrypt_key_len,
            package_key_len,
            PasswordKeyIvDerivation::Derived,
        ),
        Err(other) => Err(other),
    }
}

#[derive(Debug)]
enum HmacCtx {
    Sha1(Hmac<sha1::Sha1>),
    Sha256(Hmac<sha2::Sha256>),
    Sha384(Hmac<sha2::Sha384>),
    Sha512(Hmac<sha2::Sha512>),
}

impl HmacCtx {
    fn new(alg: HashAlgorithm, key: &[u8]) -> Result<Self> {
        match alg {
            HashAlgorithm::Sha1 => Ok(Self::Sha1(Hmac::new_from_slice(key).map_err(|e| {
                OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "encryptedHmacKey".to_string(),
                    reason: e.to_string(),
                }
            })?)),
            HashAlgorithm::Sha256 => Ok(Self::Sha256(Hmac::new_from_slice(key).map_err(|e| {
                OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "encryptedHmacKey".to_string(),
                    reason: e.to_string(),
                }
            })?)),
            HashAlgorithm::Sha384 => Ok(Self::Sha384(Hmac::new_from_slice(key).map_err(|e| {
                OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "encryptedHmacKey".to_string(),
                    reason: e.to_string(),
                }
            })?)),
            HashAlgorithm::Sha512 => Ok(Self::Sha512(Hmac::new_from_slice(key).map_err(|e| {
                OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "encryptedHmacKey".to_string(),
                    reason: e.to_string(),
                }
            })?)),
        }
    }

    fn update(&mut self, data: &[u8]) {
        match self {
            HmacCtx::Sha1(mac) => mac.update(data),
            HmacCtx::Sha256(mac) => mac.update(data),
            HmacCtx::Sha384(mac) => mac.update(data),
            HmacCtx::Sha512(mac) => mac.update(data),
        }
    }

    fn finalize(self) -> Vec<u8> {
        match self {
            HmacCtx::Sha1(mac) => mac.finalize().into_bytes().to_vec(),
            HmacCtx::Sha256(mac) => mac.finalize().into_bytes().to_vec(),
            HmacCtx::Sha384(mac) => mac.finalize().into_bytes().to_vec(),
            HmacCtx::Sha512(mac) => mac.finalize().into_bytes().to_vec(),
        }
    }
}

/// Decrypt an MS-OFFCRYPTO Agile `EncryptedPackage` stream (OOXML password protection).
///
/// Inputs are the raw bytes of the CFB streams:
/// - `EncryptionInfo`
/// - `EncryptedPackage`
///
/// Returns the decrypted OOXML package bytes (a ZIP file).
///
/// When the `<dataIntegrity>` element is present in the `EncryptionInfo` XML, this function
/// verifies the HMAC integrity value as described by MS-OFFCRYPTO. Some real-world producers omit
/// `<dataIntegrity>` entirely; in that case, decryption proceeds but **no integrity verification**
/// is performed.
pub fn decrypt_agile_encrypted_package(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> Result<Vec<u8>> {
    decrypt_agile_encrypted_package_with_options(
        encryption_info,
        encrypted_package,
        password,
        &DecryptOptions::default(),
    )
}

/// Like [`decrypt_agile_encrypted_package`] but with configurable [`DecryptOptions`].
pub fn decrypt_agile_encrypted_package_with_options(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
    opts: &DecryptOptions,
) -> Result<Vec<u8>> {
    decrypt_agile_encrypted_package_impl(encryption_info, encrypted_package, password, opts, None)
}

/// Decrypt an MS-OFFCRYPTO Agile `EncryptedPackage` stream (OOXML password protection), collecting
/// non-fatal parse/decrypt warnings.
///
/// This is identical to [`decrypt_agile_encrypted_package`], but returns a vector of
/// [`OffCryptoWarning`] values that describe irregularities encountered while parsing the
/// `EncryptionInfo` XML. These warnings are intended for diagnostics/telemetry and never include
/// sensitive data (passwords, derived keys, decrypted bytes).
pub fn decrypt_agile_encrypted_package_with_warnings(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> Result<(Vec<u8>, Vec<OffCryptoWarning>)> {
    let mut warnings = Vec::new();
    let plaintext = decrypt_agile_encrypted_package_impl(
        encryption_info,
        encrypted_package,
        password,
        &DecryptOptions::default(),
        Some(&mut warnings),
    )?;
    Ok((plaintext, warnings))
}

fn decrypt_agile_encrypted_package_impl(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
    decrypt_opts: &DecryptOptions,
    mut warnings: Option<&mut Vec<OffCryptoWarning>>,
) -> Result<Vec<u8>> {
    let info = parse_agile_encryption_info(
        encryption_info,
        decrypt_opts,
        warnings.as_mut().map(|w| &mut **w),
    )?;

    // Validate AES-CBC ciphertext buffers up-front to avoid confusing crypto backend errors and to
    // ensure we can report which field was malformed.
    validate_ciphertext_block_aligned(
        "encryptedVerifierHashInput",
        &info.password_key.encrypted_verifier_hash_input,
    )?;
    validate_ciphertext_block_aligned(
        "encryptedVerifierHashValue",
        &info.password_key.encrypted_verifier_hash_value,
    )?;
    validate_ciphertext_block_aligned("encryptedKeyValue", &info.password_key.encrypted_key_value)?;
    if let Some(data_integrity) = &info.data_integrity {
        validate_ciphertext_block_aligned(
            "dataIntegrity.encryptedHmacKey",
            &data_integrity.encrypted_hmac_key,
        )?;
        validate_ciphertext_block_aligned(
            "dataIntegrity.encryptedHmacValue",
            &data_integrity.encrypted_hmac_value,
        )?;
    }

    // 1) Verify password and unwrap the package key ("keyValue").
    let password_hash = hash_password(
        password,
        &info.password_key.salt_value,
        info.password_key.spin_count,
        info.password_key.hash_algorithm,
    )
    .map_err(|e| OffCryptoError::InvalidAttribute {
        element: "p:encryptedKey".to_string(),
        attr: "hash_password".to_string(),
        reason: e.to_string(),
    })?;

    let key_encrypt_key_len =
        key_len_bytes(info.password_key.key_bits, "p:encryptedKey", "keyBits")?;
    let package_key_len = key_len_bytes(info.key_data.key_bits, "keyData", "keyBits")?;
    // Some producers vary how the AES-CBC IV is derived for the password-key-encryptor blobs.
    // Try both strategies for compatibility.
    let key_value = decrypt_agile_package_key_from_password_best_effort(
        &info,
        &password_hash,
        key_encrypt_key_len,
        package_key_len,
    )?;

    // 2) Decrypt EncryptedPackage stream to plaintext ZIP bytes.
    let (declared_len, ciphertext) = parse_encrypted_package_stream(encrypted_package)?;

    if ciphertext.len() % AES_BLOCK_SIZE != 0 {
        return Err(OffCryptoError::CiphertextNotBlockAligned {
            field: "EncryptedPackage",
            len: ciphertext.len(),
        });
    }

    let mut plaintext = Vec::with_capacity(ciphertext.len());
    for (idx, chunk) in ciphertext.chunks(SEGMENT_SIZE).enumerate() {
        if chunk.len() % AES_BLOCK_SIZE != 0 {
            return Err(OffCryptoError::CiphertextNotBlockAligned {
                field: "EncryptedPackage",
                len: chunk.len(),
            });
        }
        let block_key = segment_block_key(idx as u32);
        let iv = derive_iv_or_err(
            &info.key_data.salt_value,
            &block_key,
            info.key_data.block_size,
            info.key_data.hash_algorithm,
        )?;
        let decrypted = decrypt_aes_cbc_no_padding(&key_value, &iv, chunk).map_err(|e| {
            OffCryptoError::InvalidAttribute {
                element: "EncryptedPackage".to_string(),
                attr: "ciphertext".to_string(),
                reason: e.to_string(),
            }
        })?;
        plaintext.extend_from_slice(&decrypted);
    }

    if plaintext.len() < declared_len {
        return Err(OffCryptoError::DecryptedLengthShorterThanHeader {
            declared_len,
            available_len: plaintext.len(),
        });
    }
    plaintext.truncate(declared_len);

    // 3) Validate dataIntegrity HMAC (when present).
    if let Some(data_integrity) = &info.data_integrity {
        let hmac_key = {
            let iv = derive_iv_or_err(
                &info.key_data.salt_value,
                &HMAC_KEY_BLOCK,
                info.key_data.block_size,
                info.key_data.hash_algorithm,
            )?;
            let decrypted =
                decrypt_aes_cbc_no_padding(&key_value, &iv, &data_integrity.encrypted_hmac_key)
                    .map_err(|e| OffCryptoError::InvalidAttribute {
                        element: "dataIntegrity".to_string(),
                        attr: "encryptedHmacKey".to_string(),
                        reason: e.to_string(),
                    })?;
            // HMAC accepts any key length; some producers emit a shorter decrypted key than
            // `hashSize`. Be tolerant and use as much key material as available.
            let key_len = std::cmp::min(info.key_data.hash_size, decrypted.len());
            if key_len == 0 {
                return Err(OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "encryptedHmacKey".to_string(),
                    reason: "decrypted HMAC key is empty".to_string(),
                });
            }
            decrypted[..key_len].to_vec()
        };

        let expected_hmac = {
            let iv = derive_iv_or_err(
                &info.key_data.salt_value,
                &HMAC_VALUE_BLOCK,
                info.key_data.block_size,
                info.key_data.hash_algorithm,
            )?;
            let decrypted = decrypt_aes_cbc_no_padding(
                &key_value,
                &iv,
                &data_integrity.encrypted_hmac_value,
            )
            .map_err(|e| OffCryptoError::InvalidAttribute {
                element: "dataIntegrity".to_string(),
                attr: "encryptedHmacValue".to_string(),
                reason: e.to_string(),
            })?;
            decrypted
                .get(..info.key_data.hash_size)
                .ok_or_else(|| OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "encryptedHmacValue".to_string(),
                    reason: "decrypted HMAC value shorter than hashSize".to_string(),
                })?
                .to_vec()
        };

        // MS-OFFCRYPTO describes `dataIntegrity` as an HMAC over the **EncryptedPackage stream bytes**
        // (length prefix + ciphertext). This matches Excel and the `ms-offcrypto-writer` crate.
        //
        // However, some producers appear to compute the HMAC over the **decrypted package bytes**
        // (plaintext ZIP bytes) instead. To be compatible with both, accept either hash target.
        let actual_hmac_ciphertext =
            compute_hmac(info.key_data.hash_algorithm, &hmac_key, encrypted_package)?;
        let actual_hmac_ciphertext = actual_hmac_ciphertext
            .get(..info.key_data.hash_size)
            .ok_or_else(|| OffCryptoError::InvalidAttribute {
                element: "dataIntegrity".to_string(),
                attr: "hashAlgorithm".to_string(),
                reason: "HMAC output shorter than hashSize".to_string(),
            })?;

        if !ct_eq(actual_hmac_ciphertext, &expected_hmac) {
            let actual_hmac_plaintext =
                compute_hmac(info.key_data.hash_algorithm, &hmac_key, &plaintext)?;
            let actual_hmac_plaintext = actual_hmac_plaintext
                .get(..info.key_data.hash_size)
                .ok_or_else(|| OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "hashAlgorithm".to_string(),
                    reason: "HMAC output shorter than hashSize".to_string(),
                })?;
            if !ct_eq(actual_hmac_plaintext, &expected_hmac) {
                return Err(OffCryptoError::IntegrityMismatch);
            }
        }
    }

    Ok(plaintext)
}

fn validate_ciphertext_block_aligned(field: &'static str, ciphertext: &[u8]) -> Result<()> {
    if ciphertext.len() % AES_BLOCK_SIZE != 0 {
        return Err(OffCryptoError::CiphertextNotBlockAligned {
            field,
            len: ciphertext.len(),
        });
    }
    Ok(())
}

/// Decrypt an MS-OFFCRYPTO Agile `EncryptedPackage` stream incrementally (OOXML password protection).
///
/// Unlike [`decrypt_agile_encrypted_package`], this function avoids loading the full ciphertext into
/// memory by streaming the `EncryptedPackage` bytes from `encrypted_package_stream` while
/// simultaneously computing and validating the `dataIntegrity` HMAC over the encrypted bytes (when
/// `<dataIntegrity>` is present).
///
/// Some real-world producers omit `<dataIntegrity>` entirely; in that case, this function still
/// decrypts the package but performs **no integrity verification**.
///
/// Returns the declared plaintext length (from the `EncryptedPackage` 8-byte header).
pub fn decrypt_agile_encrypted_package_stream<R: Read + Seek, W: Write>(
    encryption_info: &[u8],
    encrypted_package_stream: &mut R,
    password: &str,
    out: &mut W,
) -> Result<u64> {
    decrypt_agile_encrypted_package_stream_with_options(
        encryption_info,
        encrypted_package_stream,
        password,
        out,
        &DecryptOptions::default(),
    )
}

/// Like [`decrypt_agile_encrypted_package_stream`] but with configurable [`DecryptOptions`].
pub fn decrypt_agile_encrypted_package_stream_with_options<R: Read + Seek, W: Write>(
    encryption_info: &[u8],
    encrypted_package_stream: &mut R,
    password: &str,
    out: &mut W,
    opts: &DecryptOptions,
) -> Result<u64> {
    let info = parse_agile_encryption_info(encryption_info, opts, None)?;

    // Validate AES-CBC ciphertext buffers up-front to avoid confusing crypto backend errors and to
    // ensure we can report which field was malformed.
    validate_ciphertext_block_aligned(
        "encryptedVerifierHashInput",
        &info.password_key.encrypted_verifier_hash_input,
    )?;
    validate_ciphertext_block_aligned(
        "encryptedVerifierHashValue",
        &info.password_key.encrypted_verifier_hash_value,
    )?;
    validate_ciphertext_block_aligned("encryptedKeyValue", &info.password_key.encrypted_key_value)?;
    if let Some(data_integrity) = &info.data_integrity {
        validate_ciphertext_block_aligned(
            "dataIntegrity.encryptedHmacKey",
            &data_integrity.encrypted_hmac_key,
        )?;
        validate_ciphertext_block_aligned(
            "dataIntegrity.encryptedHmacValue",
            &data_integrity.encrypted_hmac_value,
        )?;
    }

    // 1) Verify password and unwrap the package key ("keyValue").
    let password_hash = hash_password(
        password,
        &info.password_key.salt_value,
        info.password_key.spin_count,
        info.password_key.hash_algorithm,
    )
    .map_err(|e| OffCryptoError::InvalidAttribute {
        element: "p:encryptedKey".to_string(),
        attr: "hash_password".to_string(),
        reason: e.to_string(),
    })?;

    let key_encrypt_key_len =
        key_len_bytes(info.password_key.key_bits, "p:encryptedKey", "keyBits")?;
    let package_key_len = key_len_bytes(info.key_data.key_bits, "keyData", "keyBits")?;
    // Some producers vary how the AES-CBC IV is derived for the password-key-encryptor blobs.
    // Try both strategies for compatibility.
    let key_value = decrypt_agile_package_key_from_password_best_effort(
        &info,
        &password_hash,
        key_encrypt_key_len,
        package_key_len,
    )?;

    // 2) Decrypt the integrity HMAC key/value (encrypted with the package key), when present.
    let mut ciphertext_mac: Option<HmacCtx> = None;
    let mut plaintext_mac: Option<HmacCtx> = None;
    let mut expected_hmac: Option<Vec<u8>> = None;

    if let Some(data_integrity) = &info.data_integrity {
        let hmac_key = {
            let iv = derive_iv_or_err(
                &info.key_data.salt_value,
                &HMAC_KEY_BLOCK,
                info.key_data.block_size,
                info.key_data.hash_algorithm,
            )?;
            let decrypted =
                decrypt_aes_cbc_no_padding(&key_value, &iv, &data_integrity.encrypted_hmac_key)
                    .map_err(|e| OffCryptoError::InvalidAttribute {
                        element: "dataIntegrity".to_string(),
                        attr: "encryptedHmacKey".to_string(),
                        reason: e.to_string(),
                    })?;
            let key_len = std::cmp::min(info.key_data.hash_size, decrypted.len());
            if key_len == 0 {
                return Err(OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "encryptedHmacKey".to_string(),
                    reason: "decrypted HMAC key is empty".to_string(),
                });
            }
            decrypted[..key_len].to_vec()
        };

        let expected = {
            let iv = derive_iv_or_err(
                &info.key_data.salt_value,
                &HMAC_VALUE_BLOCK,
                info.key_data.block_size,
                info.key_data.hash_algorithm,
            )?;
            let decrypted = decrypt_aes_cbc_no_padding(
                &key_value,
                &iv,
                &data_integrity.encrypted_hmac_value,
            )
            .map_err(|e| OffCryptoError::InvalidAttribute {
                element: "dataIntegrity".to_string(),
                attr: "encryptedHmacValue".to_string(),
                reason: e.to_string(),
            })?;
            decrypted
                .get(..info.key_data.hash_size)
                .ok_or_else(|| OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "encryptedHmacValue".to_string(),
                    reason: "decrypted HMAC value shorter than hashSize".to_string(),
                })?
                .to_vec()
        };

        ciphertext_mac = Some(HmacCtx::new(info.key_data.hash_algorithm, &hmac_key)?);
        plaintext_mac = Some(HmacCtx::new(info.key_data.hash_algorithm, &hmac_key)?);
        expected_hmac = Some(expected);
    }

    // 3) Stream-decrypt the ciphertext while optionally computing the HMAC over the encrypted bytes.
    let encrypted_package_len = encrypted_package_stream
        .seek(SeekFrom::End(0))
        .map_err(|source| OffCryptoError::Io {
            context: "seeking EncryptedPackage to end",
            source,
        })?;
    encrypted_package_stream
        .seek(SeekFrom::Start(0))
        .map_err(|source| OffCryptoError::Io {
            context: "seeking EncryptedPackage to start",
            source,
        })?;

    let mut header = [0u8; 8];
    encrypted_package_stream
        .read_exact(&mut header)
        .map_err(|source| OffCryptoError::Io {
            context: "reading EncryptedPackage length header",
            source,
        })?;
    if let Some(mac) = ciphertext_mac.as_mut() {
        mac.update(&header);
    }
    // `original_package_size` is an 8-byte plaintext prefix. While MS-OFFCRYPTO describes it as a
    // `u64le`, some producers/libraries treat it as `u32 totalSize` + `u32 reserved` (often 0).
    //
    // For compatibility, parse as two DWORDs and fall back to the low DWORD when the combined
    // 64-bit value is not plausible for the available ciphertext.
    //
    // Avoid falling back when the low DWORD is zero: some real files may have true 64-bit sizes
    // that are exact multiples of 2^32 (lo=0, hi!=0).
    let len_lo = u32::from_le_bytes(header[..4].try_into().expect("slice length checked")) as u64;
    let len_hi = u32::from_le_bytes(header[4..].try_into().expect("slice length checked")) as u64;
    let declared_len_u64 = len_lo | (len_hi << 32);
    let ciphertext_len = encrypted_package_len.saturating_sub(8);
    let declared_len =
        if len_lo != 0 && len_hi != 0 && declared_len_u64 > ciphertext_len && len_lo <= ciphertext_len {
            len_lo
        } else {
            declared_len_u64
        };

    let mut remaining_to_write = declared_len;
    let mut written: u64 = 0;

    let mut segment_index: u32 = 0;
    let mut buf = [0u8; SEGMENT_SIZE];

    loop {
        let mut filled = 0usize;
        while filled < SEGMENT_SIZE {
            let n = encrypted_package_stream
                .read(&mut buf[filled..])
                .map_err(|source| OffCryptoError::Io {
                    context: "reading EncryptedPackage ciphertext",
                    source,
                })?;
            if n == 0 {
                break;
            }
            filled += n;
        }

        if filled == 0 {
            break;
        }

        validate_ciphertext_block_aligned("EncryptedPackage", &buf[..filled])?;

        if let Some(mac) = ciphertext_mac.as_mut() {
            mac.update(&buf[..filled]);
        }

        if remaining_to_write > 0 {
            let block_key = segment_block_key(segment_index);
            let iv = derive_iv_or_err(
                &info.key_data.salt_value,
                &block_key,
                info.key_data.block_size,
                info.key_data.hash_algorithm,
            )?;
            decrypt_aes_cbc_no_padding_in_place(&key_value, &iv, &mut buf[..filled]).map_err(
                |e| OffCryptoError::InvalidAttribute {
                    element: "EncryptedPackage".to_string(),
                    attr: "ciphertext".to_string(),
                    reason: e.to_string(),
                },
            )?;

            let to_write = std::cmp::min(remaining_to_write, filled as u64) as usize;
            if let Some(mac) = plaintext_mac.as_mut() {
                mac.update(&buf[..to_write]);
            }
            out.write_all(&buf[..to_write]).map_err(|source| OffCryptoError::Io {
                context: "writing decrypted plaintext",
                source,
            })?;
            remaining_to_write -= to_write as u64;
            written += to_write as u64;
        }

        segment_index =
            segment_index
                .checked_add(1)
                .ok_or_else(|| OffCryptoError::InvalidAttribute {
                    element: "EncryptedPackage".to_string(),
                    attr: "segmentIndex".to_string(),
                    reason: "EncryptedPackage segment index overflow".to_string(),
                })?;
    }

    if remaining_to_write != 0 {
        return Err(OffCryptoError::DecryptedLengthShorterThanHeader {
            declared_len: usize::try_from(declared_len).unwrap_or(usize::MAX),
            available_len: usize::try_from(written).unwrap_or(usize::MAX),
        });
    }

    if let Some(expected_hmac) = expected_hmac {
        let ciphertext_mac = ciphertext_mac
            .take()
            .expect("expected ciphertext HMAC context when expected_hmac is present");
        let plaintext_mac = plaintext_mac
            .take()
            .expect("expected plaintext HMAC context when expected_hmac is present");

        let actual_ciphertext_hmac_full = ciphertext_mac.finalize();
        let actual_ciphertext_hmac = actual_ciphertext_hmac_full
            .get(..info.key_data.hash_size)
            .ok_or_else(|| OffCryptoError::InvalidAttribute {
                element: "dataIntegrity".to_string(),
                attr: "hashAlgorithm".to_string(),
                reason: "HMAC output shorter than hashSize".to_string(),
            })?;

        if !ct_eq(actual_ciphertext_hmac, &expected_hmac) {
            let actual_plaintext_hmac_full = plaintext_mac.finalize();
            let actual_plaintext_hmac = actual_plaintext_hmac_full
                .get(..info.key_data.hash_size)
                .ok_or_else(|| OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "hashAlgorithm".to_string(),
                    reason: "HMAC output shorter than hashSize".to_string(),
                })?;
            if !ct_eq(actual_plaintext_hmac, &expected_hmac) {
                return Err(OffCryptoError::IntegrityMismatch);
            }
        }
    }

    Ok(declared_len)
}

fn parse_encrypted_package_stream(encrypted_package: &[u8]) -> Result<(usize, &[u8])> {
    if encrypted_package.len() < 8 {
        return Err(OffCryptoError::EncryptedPackageTooShort {
            len: encrypted_package.len(),
        });
    }

    // `original_package_size` is an 8-byte plaintext prefix. While MS-OFFCRYPTO describes it as a
    // `u64le`, some producers/libraries treat it as `u32 totalSize` + `u32 reserved` (often 0).
    //
    // To improve compatibility, parse as two DWORDs and apply a small heuristic:
    // - if the combined 64-bit size is larger than the available ciphertext, but the low DWORD is
    //   plausible, treat the upper DWORD as "reserved" and fall back to the low DWORD.
    //
    // Avoid falling back when the low DWORD is zero: some real files may have true 64-bit sizes
    // that are exact multiples of 2^32 (lo=0, hi!=0).
    let len_lo = u32::from_le_bytes(
        encrypted_package[..4]
            .try_into()
            .expect("slice length already checked"),
    ) as u64;
    let len_hi = u32::from_le_bytes(
        encrypted_package[4..8]
            .try_into()
            .expect("slice length already checked"),
    ) as u64;
    let declared_len_u64 = len_lo | (len_hi << 32);

    let ciphertext_len = encrypted_package.len() - 8;
    let effective_len_u64 = if len_lo != 0
        && len_hi != 0
        && declared_len_u64 > ciphertext_len as u64
        && len_lo <= ciphertext_len as u64
    {
        len_lo
    } else {
        declared_len_u64
    };

    let declared_len =
        usize::try_from(effective_len_u64).map_err(|_| OffCryptoError::InvalidAttribute {
            element: "EncryptedPackage".to_string(),
            attr: "original_package_size".to_string(),
            reason: format!("declared size {effective_len_u64} does not fit in usize"),
        })?;

    Ok((declared_len, &encrypted_package[8..]))
}

fn parse_agile_encryption_info(
    encryption_info: &[u8],
    decrypt_opts: &DecryptOptions,
    mut warnings: Option<&mut Vec<OffCryptoWarning>>,
) -> Result<AgileEncryptionInfo> {
    let parse_opts = ParseOptions::default();
    if encryption_info.len() < 8 {
        return Err(OffCryptoError::MissingRequiredElement {
            element: "EncryptionInfoHeader".to_string(),
        });
    }

    let major = u16::from_le_bytes([encryption_info[0], encryption_info[1]]);
    let minor = u16::from_le_bytes([encryption_info[2], encryption_info[3]]);
    if (major, minor) != (4, 4) {
        return Err(OffCryptoError::UnsupportedEncryptionVersion { major, minor });
    }

    let xml_bytes = extract_encryption_info_xml(encryption_info, &parse_opts)?;
    let xml = decode_encryption_info_xml_text(xml_bytes)?;
    let doc = roxmltree::Document::parse(xml.as_ref())?;

    if let Some(w) = warnings.as_deref_mut() {
        collect_xml_warnings(&doc, w);
    }

    let key_data_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "keyData")
        .ok_or_else(|| OffCryptoError::MissingRequiredElement {
            element: "keyData".to_string(),
        })?;
    let data_integrity_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "dataIntegrity");

    let key_encryptors_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "keyEncryptors")
        .ok_or_else(|| OffCryptoError::MissingRequiredElement {
            element: "keyEncryptors".to_string(),
        })?;

    // Office can emit multiple key encryptors (e.g. password + certificate). We only support
    // password-based encryption; select the first password encryptor deterministically.
    let mut available_uris: Vec<String> = Vec::new();
    let mut selected_password_encryptor: Option<roxmltree::Node<'_, '_>> = None;
    let mut password_encryptor_count = 0usize;
    for key_encryptor in key_encryptors_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "keyEncryptor")
    {
        let uri = key_encryptor.attribute("uri").ok_or_else(|| {
            OffCryptoError::MissingRequiredAttribute {
                element: "keyEncryptor".to_string(),
                attr: "uri".to_string(),
            }
        })?;

        if !available_uris.iter().any(|u| u == uri) {
            available_uris.push(uri.to_string());
        }

        if uri == KEY_ENCRYPTOR_URI_PASSWORD {
            password_encryptor_count += 1;
            if selected_password_encryptor.is_none() {
                selected_password_encryptor = Some(key_encryptor);
            }
        }
    }

    if let Some(w) = warnings.as_deref_mut() {
        if password_encryptor_count > 1 {
            push_warning_dedup(
                w,
                OffCryptoWarning::MultiplePasswordKeyEncryptors {
                    count: password_encryptor_count,
                },
            );
        }
    }

    let Some(key_encryptor_node) = selected_password_encryptor else {
        let mut msg = String::new();
        msg.push_str("unsupported key encryptor in Agile EncryptionInfo: ");
        msg.push_str("Formula currently supports only password-based encryption. ");

        if available_uris.is_empty() {
            msg.push_str("No `<keyEncryptor>` entries were found.");
        } else {
            msg.push_str("Found keyEncryptor URIs: ");
            msg.push_str(&available_uris.join(", "));
            msg.push('.');
        }

        if available_uris
            .iter()
            .any(|u| u == KEY_ENCRYPTOR_URI_CERTIFICATE)
        {
            msg.push_str(" This file appears to be certificate-encrypted (public/private key) rather than password-encrypted. Re-save the workbook in Excel using “Encrypt with Password”.");
        } else {
            msg.push_str(" Re-save the workbook in Excel using “Encrypt with Password” (not certificate-based protection).");
        }

        return Err(OffCryptoError::UnsupportedKeyEncryptor {
            available_uris,
            message: msg,
        });
    };

    let encrypted_key_node = key_encryptor_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "encryptedKey")
        .ok_or_else(|| OffCryptoError::MissingRequiredElement {
            element: "encryptedKey".to_string(),
        })?;

    let key_data = parse_key_data(key_data_node, &parse_opts, warnings.as_deref_mut())?;
    let data_integrity = data_integrity_node
        .map(|node| parse_data_integrity(node, &parse_opts))
        .transpose()?;
    if data_integrity.is_none() {
        if let Some(w) = warnings.as_deref_mut() {
            push_warning_dedup(w, OffCryptoWarning::MissingDataIntegrity);
        }
    }
    let password_key = parse_password_key_encryptor(
        encrypted_key_node,
        &parse_opts,
        decrypt_opts,
        warnings.as_deref_mut(),
    )?;

    Ok(AgileEncryptionInfo {
        key_data,
        data_integrity,
        password_key,
    })
}

fn parse_key_data(
    node: roxmltree::Node<'_, '_>,
    opts: &ParseOptions,
    warnings: Option<&mut Vec<OffCryptoWarning>>,
) -> Result<KeyData> {
    validate_cipher_settings(node)?;

    let salt_size = parse_usize_attr(node, "saltSize")?;
    if salt_size == 0 {
        return Err(OffCryptoError::InvalidAttribute {
            element: "keyData".to_string(),
            attr: "saltSize".to_string(),
            reason: "saltSize must be non-zero".to_string(),
        });
    }

    let salt_value = parse_base64_attr(node, "saltValue", opts)?;
    if salt_value.len() != salt_size {
        return Err(OffCryptoError::InvalidAttribute {
            element: "keyData".to_string(),
            attr: "saltValue".to_string(),
            reason: format!(
                "decoded saltValue length {} does not match saltSize {}",
                salt_value.len(),
                salt_size
            ),
        });
    }
    let hash_algorithm = parse_hash_algorithm(node, "hashAlgorithm")?;
    let block_size = parse_usize_attr(node, "blockSize")?;
    let key_bits = parse_usize_attr(node, "keyBits")?;
    let hash_size = parse_usize_attr(node, "hashSize")?;

    if let Some(w) = warnings {
        maybe_warn_hash_size(w, "keyData", hash_algorithm, hash_size);
        maybe_warn_salt_size(w, "keyData", node.attribute("saltSize"), salt_value.len());
    }

    validate_block_size(block_size)?;
    validate_hash_size(node, "hashSize", hash_algorithm, hash_size)?;
    // Prevent unbounded allocations later in the decrypt path (key derivation).
    key_len_bytes(key_bits, "keyData", "keyBits")?;

    Ok(KeyData {
        salt_value,
        hash_algorithm,
        block_size,
        key_bits,
        hash_size,
    })
}

fn parse_data_integrity(
    node: roxmltree::Node<'_, '_>,
    opts: &ParseOptions,
) -> Result<DataIntegrity> {
    Ok(DataIntegrity {
        encrypted_hmac_key: parse_base64_attr(node, "encryptedHmacKey", opts)?,
        encrypted_hmac_value: parse_base64_attr(node, "encryptedHmacValue", opts)?,
    })
}

fn parse_password_key_encryptor(
    node: roxmltree::Node<'_, '_>,
    parse_opts: &ParseOptions,
    decrypt_opts: &DecryptOptions,
    warnings: Option<&mut Vec<OffCryptoWarning>>,
) -> Result<PasswordKeyEncryptor> {
    validate_cipher_settings(node)?;

    // `spinCount` is attacker-controlled; enforce limits before decoding any base64 attributes so we
    // can fail fast on malicious inputs.
    let spin_count = parse_u32_attr(node, "spinCount")?;
    if spin_count > decrypt_opts.max_spin_count {
        return Err(OffCryptoError::SpinCountTooLarge {
            spin_count,
            max: decrypt_opts.max_spin_count,
        });
    }

    let salt_value = parse_base64_attr(node, "saltValue", parse_opts)?;
    let hash_algorithm = parse_hash_algorithm(node, "hashAlgorithm")?;
    let block_size = parse_usize_attr(node, "blockSize")?;
    let key_bits = parse_usize_attr(node, "keyBits")?;
    let hash_size = parse_usize_attr(node, "hashSize")?;
    let encrypted_verifier_hash_input =
        parse_base64_attr_or_child(node, "encryptedVerifierHashInput", parse_opts)?;
    let encrypted_verifier_hash_value =
        parse_base64_attr_or_child(node, "encryptedVerifierHashValue", parse_opts)?;
    let encrypted_key_value = parse_base64_attr_or_child(node, "encryptedKeyValue", parse_opts)?;

    let salt_size = parse_usize_attr(node, "saltSize")?;
    if salt_size == 0 {
        return Err(OffCryptoError::InvalidAttribute {
            element: "encryptedKey".to_string(),
            attr: "saltSize".to_string(),
            reason: "saltSize must be non-zero".to_string(),
        });
    }

    if salt_value.len() != salt_size {
        return Err(OffCryptoError::InvalidAttribute {
            element: "encryptedKey".to_string(),
            attr: "saltValue".to_string(),
            reason: format!(
                "decoded saltValue length {} does not match saltSize {}",
                salt_value.len(),
                salt_size
            ),
        });
    }

    if let Some(w) = warnings {
        maybe_warn_hash_size(w, "encryptedKey", hash_algorithm, hash_size);
        maybe_warn_salt_size(
            w,
            "encryptedKey",
            node.attribute("saltSize"),
            salt_value.len(),
        );
    }
    validate_block_size(block_size)?;
    validate_hash_size(node, "hashSize", hash_algorithm, hash_size)?;
    key_len_bytes(key_bits, "p:encryptedKey", "keyBits")?;

    Ok(PasswordKeyEncryptor {
        salt_value,
        hash_algorithm,
        spin_count,
        block_size,
        key_bits,
        hash_size,
        encrypted_verifier_hash_input,
        encrypted_verifier_hash_value,
        encrypted_key_value,
    })
}

fn push_warning_dedup(warnings: &mut Vec<OffCryptoWarning>, warning: OffCryptoWarning) {
    if warnings.iter().any(|w| w == &warning) {
        return;
    }
    warnings.push(warning);
}

fn maybe_warn_hash_size(
    warnings: &mut Vec<OffCryptoWarning>,
    element: &'static str,
    hash_alg: HashAlgorithm,
    hash_size: usize,
) {
    let expected = match hash_alg {
        HashAlgorithm::Sha1 => 20,
        HashAlgorithm::Sha256 => 32,
        HashAlgorithm::Sha384 => 48,
        HashAlgorithm::Sha512 => 64,
    };
    if hash_size > 0 && hash_size <= expected && hash_size != expected {
        push_warning_dedup(
            warnings,
            OffCryptoWarning::NonStandardHashSize {
                element,
                hash_algorithm: hash_alg,
                hash_size,
                expected_size: expected,
            },
        );
    }
}

fn maybe_warn_salt_size(
    warnings: &mut Vec<OffCryptoWarning>,
    element: &'static str,
    salt_size_attr: Option<&str>,
    salt_value_len: usize,
) {
    const DEFAULT_SALT_SIZE: usize = 16;
    let Some(raw) = salt_size_attr else {
        return;
    };
    let Ok(salt_size) = raw.trim().parse::<usize>() else {
        return;
    };

    if salt_size != salt_value_len {
        push_warning_dedup(
            warnings,
            OffCryptoWarning::SaltSizeMismatch {
                element,
                declared_salt_size: salt_size,
                salt_value_len,
            },
        );
        return;
    }

    if salt_size != DEFAULT_SALT_SIZE {
        push_warning_dedup(
            warnings,
            OffCryptoWarning::NonStandardSaltSize {
                element,
                salt_size,
                expected_size: DEFAULT_SALT_SIZE,
            },
        );
    }
}

fn collect_xml_warnings(doc: &roxmltree::Document<'_>, warnings: &mut Vec<OffCryptoWarning>) {
    // Best-effort schema/attribute validation.
    //
    // This intentionally does not attempt to validate values (which might contain sensitive base64
    // data). It only reports *names* of elements/attributes we don't understand so callers can
    // surface non-fatal anomalies for debugging/telemetry.

    fn is_allowed_element(name: &str) -> bool {
        matches!(
            name,
            "encryption"
                | "keyData"
                | "dataIntegrity"
                | "keyEncryptors"
                | "keyEncryptor"
                | "encryptedKey"
                | "encryptedVerifierHashInput"
                | "encryptedVerifierHashValue"
                | "encryptedKeyValue"
        )
    }

    fn allowed_attrs(name: &str) -> &'static [&'static str] {
        match name {
            "keyData" => &[
                "saltSize",
                "blockSize",
                "keyBits",
                "hashSize",
                "cipherAlgorithm",
                "cipherChaining",
                "hashAlgorithm",
                "saltValue",
            ],
            "dataIntegrity" => &["encryptedHmacKey", "encryptedHmacValue"],
            "keyEncryptor" => &["uri"],
            "encryptedKey" => &[
                "saltSize",
                "blockSize",
                "keyBits",
                "hashSize",
                "spinCount",
                "cipherAlgorithm",
                "cipherChaining",
                "hashAlgorithm",
                "saltValue",
                "encryptedVerifierHashInput",
                "encryptedVerifierHashValue",
                "encryptedKeyValue",
            ],
            // `encryption` and `keyEncryptors` typically only carry namespace declarations.
            _ => &[],
        }
    }

    for node in doc.descendants().filter(|n| n.is_element()) {
        let element = node.tag_name().name();
        if !is_allowed_element(element) {
            push_warning_dedup(
                warnings,
                OffCryptoWarning::UnrecognizedXmlElement {
                    element: element.to_string(),
                },
            );
        }

        let allowed = allowed_attrs(element);
        for attr in node.attributes() {
            if attr.namespace() == Some("http://www.w3.org/2000/xmlns/")
                || attr.name().starts_with("xmlns")
            {
                continue;
            }
            if allowed.iter().any(|a| a == &attr.name()) {
                continue;
            }
            push_warning_dedup(
                warnings,
                OffCryptoWarning::UnrecognizedXmlAttribute {
                    element: element.to_string(),
                    attr: attr.name().to_string(),
                },
            );
        }
    }
}

fn validate_cipher_settings(node: roxmltree::Node<'_, '_>) -> Result<()> {
    let cipher_alg = required_attr(node, "cipherAlgorithm")?.trim();
    if !cipher_alg.eq_ignore_ascii_case("AES") {
        return Err(OffCryptoError::UnsupportedCipherAlgorithm {
            cipher: cipher_alg.to_string(),
        });
    }
    let chaining = required_attr(node, "cipherChaining")?.trim();
    if !chaining.eq_ignore_ascii_case("ChainingModeCBC") {
        return Err(OffCryptoError::UnsupportedCipherChaining {
            chaining: chaining.to_string(),
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use aes::Aes128;
    use base64::engine::general_purpose::STANDARD as BASE64;
    use base64::Engine as _;
    use cbc::cipher::{block_padding::NoPadding, BlockEncryptMut, KeyIvInit};

    fn wrap_encryption_info(xml: &str) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&4u16.to_le_bytes()); // major
        out.extend_from_slice(&4u16.to_le_bytes()); // minor
        out.extend_from_slice(&0u32.to_le_bytes()); // flags
        out.extend_from_slice(xml.as_bytes());
        out
    }

    fn zero_pad(mut bytes: Vec<u8>) -> Vec<u8> {
        if bytes.is_empty() {
            return bytes;
        }
        let rem = bytes.len() % AES_BLOCK_SIZE;
        if rem == 0 {
            return bytes;
        }
        bytes.extend(std::iter::repeat(0u8).take(AES_BLOCK_SIZE - rem));
        bytes
    }

    #[test]
    fn parse_encrypted_package_stream_falls_back_to_low_dword_when_high_dword_is_reserved() {
        // Some producers store the 8-byte size prefix as `u32 totalSize` + `u32 reserved`, and the
        // reserved field may be non-zero. If the combined `u64` size is not plausible for the
        // available ciphertext, we should fall back to the low DWORD.
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&16u32.to_le_bytes()); // lo
        bytes.extend_from_slice(&1u32.to_le_bytes()); // hi (reserved/non-zero)
        bytes.extend_from_slice(&[0u8; 16]); // ciphertext

        let (declared_len, ciphertext) = parse_encrypted_package_stream(&bytes).expect("parse");
        assert_eq!(declared_len, 16);
        assert_eq!(ciphertext.len(), 16);
    }

    #[test]
    fn parse_encrypted_package_stream_does_not_fall_back_when_low_dword_is_zero() {
        // The "high DWORD reserved" fallback should not misinterpret true 64-bit sizes that are
        // exact multiples of 2^32 (low DWORD = 0).
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0u32.to_le_bytes()); // lo
        bytes.extend_from_slice(&1u32.to_le_bytes()); // hi
        bytes.extend_from_slice(&[0u8; 16]); // ciphertext

        let res = parse_encrypted_package_stream(&bytes);
        if usize::BITS < 64 {
            // 4GiB does not fit in usize on 32-bit targets.
            match &res {
                Err(OffCryptoError::InvalidAttribute { element, .. })
                    if element == "EncryptedPackage" => {}
                other => panic!("expected InvalidAttribute on 32-bit, got {other:?}"),
            }
        } else {
            let (declared_len, ciphertext) = res.expect("parse");
            assert_eq!(declared_len, (1u64 << 32) as usize);
            assert_eq!(ciphertext.len(), 16);
        }
    }

    fn encrypt_aes128_cbc_no_padding(key: &[u8], iv: &[u8], plaintext: &[u8]) -> Vec<u8> {
        assert_eq!(key.len(), 16, "AES-128 key required for test helper");
        assert_eq!(iv.len(), 16, "AES block-sized IV required");
        assert!(
            plaintext.len() % AES_BLOCK_SIZE == 0,
            "plaintext must be block-aligned"
        );

        let mut buf = plaintext.to_vec();
        let len = buf.len();
        cbc::Encryptor::<Aes128>::new_from_slices(key, iv)
            .unwrap()
            .encrypt_padded_mut::<NoPadding>(&mut buf, len)
            .unwrap();
        buf
    }

    #[test]
    fn decrypts_without_data_integrity_element() {
        // Some real-world producers omit `<dataIntegrity>` entirely. Ensure we can still decrypt
        // (without HMAC verification) as long as the password verifier blobs are present.
        let password = "password";

        // keyData (package encryption parameters).
        let key_data_salt = (0u8..=15).collect::<Vec<_>>();
        let key_data_salt_size = key_data_salt.len();
        let key_data_key_bits = 128usize;
        let key_data_block_size = 16usize;
        let key_data_hash_alg = HashAlgorithm::Sha1;
        let key_data_hash_size = 20usize;

        // password key encryptor parameters.
        let ke_salt = (16u8..=31).collect::<Vec<_>>();
        let ke_salt_size = ke_salt.len();
        let ke_spin = 10u32;
        let ke_key_bits = 128usize;
        let ke_block_size = 16usize;
        let ke_hash_alg = HashAlgorithm::Sha1;
        let ke_hash_size = 20usize;

        // Generate a deterministic package key and plaintext.
        let package_key = b"0123456789ABCDEF".to_vec(); // 16 bytes
        let plaintext = (0..5000u32).map(|i| (i % 251) as u8).collect::<Vec<_>>();

        // --- Encrypt EncryptedPackage stream (segment-wise) -----------------------------------
        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        let padded_plaintext = zero_pad(plaintext.clone());
        for (i, chunk) in padded_plaintext.chunks(SEGMENT_SIZE).enumerate() {
            let block_key = segment_block_key(i as u32);
            let iv =
                derive_iv(&key_data_salt, &block_key, key_data_block_size, key_data_hash_alg).unwrap();
            let ct = encrypt_aes128_cbc_no_padding(&package_key, &iv, chunk);
            encrypted_package.extend_from_slice(&ct);
        }

        // --- Encrypt password key-encryptor blobs ---------------------------------------------
        let pw_hash = hash_password(password, &ke_salt, ke_spin, ke_hash_alg).unwrap();

        let verifier_hash_input = b"abcdefghijklmnop".to_vec(); // 16 bytes
        let verifier_hash_value = hash_bytes(ke_hash_alg, &verifier_hash_input); // 20 bytes for SHA1

        fn encrypt_ke_blob(
            pw_hash: &[u8],
            ke_salt: &[u8],
            ke_key_bits: usize,
            ke_block_size: usize,
            ke_hash_alg: HashAlgorithm,
            block_key: &[u8],
            plaintext: &[u8],
        ) -> Vec<u8> {
            let key_len = ke_key_bits / 8;
            let key = derive_key(pw_hash, block_key, key_len, ke_hash_alg).unwrap();
            let iv = &ke_salt[..ke_block_size];
            let padded = zero_pad(plaintext.to_vec());
            encrypt_aes128_cbc_no_padding(&key, iv, &padded)
        }

        let encrypted_verifier_hash_input = encrypt_ke_blob(
            &pw_hash,
            &ke_salt,
            ke_key_bits,
            ke_block_size,
            ke_hash_alg,
            &VERIFIER_HASH_INPUT_BLOCK,
            &verifier_hash_input,
        );
        let encrypted_verifier_hash_value = encrypt_ke_blob(
            &pw_hash,
            &ke_salt,
            ke_key_bits,
            ke_block_size,
            ke_hash_alg,
            &VERIFIER_HASH_VALUE_BLOCK,
            &verifier_hash_value,
        );
        let encrypted_key_value = encrypt_ke_blob(
            &pw_hash,
            &ke_salt,
            ke_key_bits,
            ke_block_size,
            ke_hash_alg,
            &KEY_VALUE_BLOCK,
            &package_key,
        );

        // Build the EncryptionInfo XML *without* `<dataIntegrity>`.
        let xml = format!(
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltValue="{key_data_salt_b64}" saltSize="{key_data_salt_size}"
                       hashAlgorithm="SHA1" hashSize="{key_data_hash_size}"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       keyBits="{key_data_key_bits}" blockSize="{key_data_block_size}" />
              <keyEncryptors>
                <keyEncryptor uri="{KEY_ENCRYPTOR_URI_PASSWORD}">
                  <p:encryptedKey saltValue="{ke_salt_b64}" saltSize="{ke_salt_size}"
                                  spinCount="{ke_spin}" hashAlgorithm="SHA1" hashSize="{ke_hash_size}"
                                  cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  keyBits="{ke_key_bits}" blockSize="{ke_block_size}"
                                  encryptedVerifierHashInput="{evhi_b64}"
                                  encryptedVerifierHashValue="{evhv_b64}"
                                  encryptedKeyValue="{ekv_b64}"/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>"#,
            key_data_salt_b64 = BASE64.encode(&key_data_salt),
            ke_salt_b64 = BASE64.encode(&ke_salt),
            evhi_b64 = BASE64.encode(&encrypted_verifier_hash_input),
            evhv_b64 = BASE64.encode(&encrypted_verifier_hash_value),
            ekv_b64 = BASE64.encode(&encrypted_key_value),
        );

        let encryption_info = wrap_encryption_info(&xml);

        let (decrypted, warnings) =
            decrypt_agile_encrypted_package_with_warnings(&encryption_info, &encrypted_package, password)
                .expect("decrypt should succeed even without dataIntegrity");
        assert_eq!(decrypted, plaintext);
        assert!(
            warnings.contains(&OffCryptoWarning::MissingDataIntegrity),
            "expected MissingDataIntegrity warning, got: {warnings:?}"
        );
    }

    #[test]
    fn decrypts_when_data_integrity_hmac_targets_plaintext() {
        // Some non-Excel producers compute the `dataIntegrity` HMAC over the decrypted package bytes
        // (plaintext ZIP) rather than the EncryptedPackage stream bytes. Ensure we accept that
        // variant for compatibility.
        let password = "password";

        // keyData (package encryption parameters).
        let key_data_salt = (0u8..=15).collect::<Vec<_>>();
        let key_data_salt_size = key_data_salt.len();
        let key_data_key_bits = 128usize;
        let key_data_block_size = 16usize;
        let key_data_hash_alg = HashAlgorithm::Sha1;
        let key_data_hash_size = 20usize;

        // password key encryptor parameters.
        let ke_salt = (16u8..=31).collect::<Vec<_>>();
        let ke_salt_size = ke_salt.len();
        let ke_spin = 10u32;
        let ke_key_bits = 128usize;
        let ke_block_size = 16usize;
        let ke_hash_alg = HashAlgorithm::Sha1;
        let ke_hash_size = 20usize;

        // Deterministic package key + plaintext.
        let package_key = b"0123456789ABCDEF".to_vec(); // 16 bytes
        let plaintext = (0..5000u32).map(|i| (i % 251) as u8).collect::<Vec<_>>();

        // --- Encrypt EncryptedPackage stream (segment-wise) -----------------------------------
        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        let padded_plaintext = zero_pad(plaintext.clone());
        for (i, chunk) in padded_plaintext.chunks(SEGMENT_SIZE).enumerate() {
            let block_key = segment_block_key(i as u32);
            let iv =
                derive_iv(&key_data_salt, &block_key, key_data_block_size, key_data_hash_alg).unwrap();
            let ct = encrypt_aes128_cbc_no_padding(&package_key, &iv, chunk);
            encrypted_package.extend_from_slice(&ct);
        }

        // --- Encrypt password key-encryptor blobs ---------------------------------------------
        let pw_hash = hash_password(password, &ke_salt, ke_spin, ke_hash_alg).unwrap();

        let verifier_hash_input = b"abcdefghijklmnop".to_vec(); // 16 bytes
        let verifier_hash_value = hash_bytes(ke_hash_alg, &verifier_hash_input); // 20 bytes for SHA1

        fn encrypt_ke_blob(
            pw_hash: &[u8],
            ke_salt: &[u8],
            ke_key_bits: usize,
            ke_block_size: usize,
            ke_hash_alg: HashAlgorithm,
            block_key: &[u8],
            plaintext: &[u8],
        ) -> Vec<u8> {
            let key_len = ke_key_bits / 8;
            let key = derive_key(pw_hash, block_key, key_len, ke_hash_alg).unwrap();
            let iv = &ke_salt[..ke_block_size];
            let padded = zero_pad(plaintext.to_vec());
            encrypt_aes128_cbc_no_padding(&key, iv, &padded)
        }

        let encrypted_verifier_hash_input = encrypt_ke_blob(
            &pw_hash,
            &ke_salt,
            ke_key_bits,
            ke_block_size,
            ke_hash_alg,
            &VERIFIER_HASH_INPUT_BLOCK,
            &verifier_hash_input,
        );
        let encrypted_verifier_hash_value = encrypt_ke_blob(
            &pw_hash,
            &ke_salt,
            ke_key_bits,
            ke_block_size,
            ke_hash_alg,
            &VERIFIER_HASH_VALUE_BLOCK,
            &verifier_hash_value,
        );
        let encrypted_key_value = encrypt_ke_blob(
            &pw_hash,
            &ke_salt,
            ke_key_bits,
            ke_block_size,
            ke_hash_alg,
            &KEY_VALUE_BLOCK,
            &package_key,
        );

        // --- dataIntegrity (HMAC over plaintext, not ciphertext stream) ------------------------
        let hmac_key_plain = vec![0x22u8; key_data_hash_size];
        let hmac_value_plain = compute_hmac(key_data_hash_alg, &hmac_key_plain, &plaintext).unwrap();

        // Sanity: ensure the non-standard plaintext target differs from the spec target so this
        // test actually exercises the fallback.
        let hmac_value_stream =
            compute_hmac(key_data_hash_alg, &hmac_key_plain, &encrypted_package).unwrap();
        assert_ne!(
            hmac_value_stream.get(..key_data_hash_size),
            Some(hmac_value_plain.as_slice()),
            "expected plaintext-target HMAC to differ from stream-target HMAC"
        );

        let iv_hmac_key =
            derive_iv(&key_data_salt, &HMAC_KEY_BLOCK, key_data_block_size, key_data_hash_alg).unwrap();
        let encrypted_hmac_key =
            encrypt_aes128_cbc_no_padding(&package_key, &iv_hmac_key, &zero_pad(hmac_key_plain));

        let iv_hmac_val = derive_iv(
            &key_data_salt,
            &HMAC_VALUE_BLOCK,
            key_data_block_size,
            key_data_hash_alg,
        )
        .unwrap();
        let encrypted_hmac_value = encrypt_aes128_cbc_no_padding(
            &package_key,
            &iv_hmac_val,
            &zero_pad(hmac_value_plain),
        );

        // Build the EncryptionInfo XML (with `<dataIntegrity>` present).
        let xml = format!(
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltValue="{key_data_salt_b64}" saltSize="{key_data_salt_size}"
                       hashAlgorithm="SHA1" hashSize="{key_data_hash_size}"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       keyBits="{key_data_key_bits}" blockSize="{key_data_block_size}" />
              <dataIntegrity encryptedHmacKey="{hmac_key_b64}" encryptedHmacValue="{hmac_value_b64}" />
              <keyEncryptors>
                <keyEncryptor uri="{KEY_ENCRYPTOR_URI_PASSWORD}">
                  <p:encryptedKey saltValue="{ke_salt_b64}" saltSize="{ke_salt_size}"
                                  spinCount="{ke_spin}" hashAlgorithm="SHA1" hashSize="{ke_hash_size}"
                                  cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  keyBits="{ke_key_bits}" blockSize="{ke_block_size}"
                                  encryptedVerifierHashInput="{evhi_b64}"
                                  encryptedVerifierHashValue="{evhv_b64}"
                                  encryptedKeyValue="{ekv_b64}"/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>"#,
            key_data_salt_b64 = BASE64.encode(&key_data_salt),
            ke_salt_b64 = BASE64.encode(&ke_salt),
            evhi_b64 = BASE64.encode(&encrypted_verifier_hash_input),
            evhv_b64 = BASE64.encode(&encrypted_verifier_hash_value),
            ekv_b64 = BASE64.encode(&encrypted_key_value),
            hmac_key_b64 = BASE64.encode(&encrypted_hmac_key),
            hmac_value_b64 = BASE64.encode(&encrypted_hmac_value),
        );

        let encryption_info = wrap_encryption_info(&xml);
        let decrypted = decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password)
            .expect("decrypt should succeed with plaintext-target HMAC");
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn decrypts_when_password_key_encryptor_blobs_use_derived_ivs() {
        // Some producers appear to derive per-blob IVs for the password-key-encryptor blobs
        // (`encryptedVerifierHashInput`, `encryptedVerifierHashValue`, `encryptedKeyValue`) instead
        // of using `saltValue` directly. Ensure we can still decrypt via the best-effort IV retry.
        let password = "password";

        // keyData (package encryption parameters).
        let key_data_salt = (0u8..=15).collect::<Vec<_>>();
        let key_data_salt_size = key_data_salt.len();
        let key_data_key_bits = 128usize;
        let key_data_block_size = 16usize;
        let key_data_hash_alg = HashAlgorithm::Sha1;
        let key_data_hash_size = 20usize;

        // password key encryptor parameters.
        let ke_salt = (16u8..=31).collect::<Vec<_>>();
        let ke_salt_size = ke_salt.len();
        let ke_spin = 10u32;
        let ke_key_bits = 128usize;
        let ke_block_size = 16usize;
        let ke_hash_alg = HashAlgorithm::Sha1;
        let ke_hash_size = 20usize;

        // Generate a deterministic package key and plaintext.
        let package_key = b"0123456789ABCDEF".to_vec(); // 16 bytes
        let plaintext = (0..5000u32).map(|i| (i % 251) as u8).collect::<Vec<_>>();

        // --- Encrypt EncryptedPackage stream (segment-wise) -----------------------------------
        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        let padded_plaintext = zero_pad(plaintext.clone());
        for (i, chunk) in padded_plaintext.chunks(SEGMENT_SIZE).enumerate() {
            let block_key = segment_block_key(i as u32);
            let iv = derive_iv(&key_data_salt, &block_key, key_data_block_size, key_data_hash_alg)
                .unwrap();
            let ct = encrypt_aes128_cbc_no_padding(&package_key, &iv, chunk);
            encrypted_package.extend_from_slice(&ct);
        }

        // --- Encrypt password key-encryptor blobs (with derived per-blob IVs) ------------------
        let pw_hash = hash_password(password, &ke_salt, ke_spin, ke_hash_alg).unwrap();

        let verifier_hash_input = b"abcdefghijklmnop".to_vec(); // 16 bytes
        let verifier_hash_value = hash_bytes(ke_hash_alg, &verifier_hash_input); // 20 bytes for SHA1

        fn encrypt_ke_blob(
            pw_hash: &[u8],
            ke_salt: &[u8],
            ke_key_bits: usize,
            ke_block_size: usize,
            ke_hash_alg: HashAlgorithm,
            block_key: &[u8],
            plaintext: &[u8],
        ) -> Vec<u8> {
            let key_len = ke_key_bits / 8;
            let key = derive_key(pw_hash, block_key, key_len, ke_hash_alg).unwrap();
            let iv = derive_iv(ke_salt, block_key, ke_block_size, ke_hash_alg).unwrap();
            let padded = zero_pad(plaintext.to_vec());
            encrypt_aes128_cbc_no_padding(&key, &iv, &padded)
        }

        let encrypted_verifier_hash_input = encrypt_ke_blob(
            &pw_hash,
            &ke_salt,
            ke_key_bits,
            ke_block_size,
            ke_hash_alg,
            &VERIFIER_HASH_INPUT_BLOCK,
            &verifier_hash_input,
        );
        let encrypted_verifier_hash_value = encrypt_ke_blob(
            &pw_hash,
            &ke_salt,
            ke_key_bits,
            ke_block_size,
            ke_hash_alg,
            &VERIFIER_HASH_VALUE_BLOCK,
            &verifier_hash_value,
        );
        let encrypted_key_value = encrypt_ke_blob(
            &pw_hash,
            &ke_salt,
            ke_key_bits,
            ke_block_size,
            ke_hash_alg,
            &KEY_VALUE_BLOCK,
            &package_key,
        );

        // Build the EncryptionInfo XML *without* `<dataIntegrity>`.
        let xml = format!(
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltValue="{key_data_salt_b64}" saltSize="{key_data_salt_size}"
                       hashAlgorithm="SHA1" hashSize="{key_data_hash_size}"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       keyBits="{key_data_key_bits}" blockSize="{key_data_block_size}" />
              <keyEncryptors>
                <keyEncryptor uri="{KEY_ENCRYPTOR_URI_PASSWORD}">
                  <p:encryptedKey saltValue="{ke_salt_b64}" saltSize="{ke_salt_size}"
                                  spinCount="{ke_spin}" hashAlgorithm="SHA1" hashSize="{ke_hash_size}"
                                  cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  keyBits="{ke_key_bits}" blockSize="{ke_block_size}"
                                  encryptedVerifierHashInput="{evhi_b64}"
                                  encryptedVerifierHashValue="{evhv_b64}"
                                  encryptedKeyValue="{ekv_b64}"/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>"#,
            key_data_salt_b64 = BASE64.encode(&key_data_salt),
            ke_salt_b64 = BASE64.encode(&ke_salt),
            evhi_b64 = BASE64.encode(&encrypted_verifier_hash_input),
            evhv_b64 = BASE64.encode(&encrypted_verifier_hash_value),
            ekv_b64 = BASE64.encode(&encrypted_key_value),
        );

        let encryption_info = wrap_encryption_info(&xml);

        let (decrypted, warnings) =
            decrypt_agile_encrypted_package_with_warnings(&encryption_info, &encrypted_package, password)
                .expect("decrypt should succeed with derived password-key IVs");
        assert_eq!(decrypted, plaintext);
        assert!(
            warnings.contains(&OffCryptoWarning::MissingDataIntegrity),
            "expected MissingDataIntegrity warning, got: {warnings:?}"
        );
    }

    #[test]
    fn rejects_cfb_cipher_chaining_in_key_data() {
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltValue="AA==" saltSize="1" hashAlgorithm="SHA1" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCFB"
                        keyBits="128" blockSize="16" />
              <dataIntegrity encryptedHmacKey="AA==" encryptedHmacValue="AA==" />
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey saltValue="AA==" saltSize="1" spinCount="1" hashAlgorithm="SHA1" hashSize="20"
                                  cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  keyBits="128" blockSize="16"
                                  encryptedVerifierHashInput="AA=="
                                  encryptedVerifierHashValue="AA=="
                                  encryptedKeyValue="AA=="/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let encryption_info = wrap_encryption_info(xml);
        let err = decrypt_agile_encrypted_package(&encryption_info, &[], "pw").unwrap_err();
        assert!(
            matches!(err, OffCryptoError::UnsupportedCipherChaining { ref chaining } if chaining == "ChainingModeCFB"),
            "unexpected error: {err:?}"
        );
        let msg = err.to_string();
        assert!(
            msg.contains("only") && msg.contains("ChainingModeCBC"),
            "expected message to mention only CBC is supported, got: {msg}"
        );
    }

    #[test]
    fn rejects_cfb_cipher_chaining_in_encrypted_key() {
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltValue="AA==" saltSize="1" hashAlgorithm="SHA1" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                        keyBits="128" blockSize="16" />
              <dataIntegrity encryptedHmacKey="AA==" encryptedHmacValue="AA==" />
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey saltValue="AA==" saltSize="1" spinCount="1" hashAlgorithm="SHA1" hashSize="20"
                                  cipherAlgorithm="AES" cipherChaining="ChainingModeCFB"
                                  keyBits="128" blockSize="16"
                                  encryptedVerifierHashInput="AA=="
                                  encryptedVerifierHashValue="AA=="
                                  encryptedKeyValue="AA=="/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let encryption_info = wrap_encryption_info(xml);
        let err = decrypt_agile_encrypted_package(&encryption_info, &[], "pw").unwrap_err();
        assert!(
            matches!(err, OffCryptoError::UnsupportedCipherChaining { ref chaining } if chaining == "ChainingModeCFB"),
            "unexpected error: {err:?}"
        );
    }

    #[test]
    fn rejects_aes_block_size_not_16() {
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltValue="AAAAAAAAAAAAAAAAAAAAAA==" saltSize="16" hashAlgorithm="SHA1" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       keyBits="128" blockSize="32" />
              <dataIntegrity encryptedHmacKey="AA==" encryptedHmacValue="AA==" />
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey saltValue="AAAAAAAAAAAAAAAAAAAAAA==" saltSize="16" spinCount="1"
                                  hashAlgorithm="SHA1" hashSize="20"
                                  cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  keyBits="128" blockSize="16"
                                  encryptedVerifierHashInput="AA=="
                                  encryptedVerifierHashValue="AA=="
                                  encryptedKeyValue="AA=="/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let encryption_info = wrap_encryption_info(xml);
        let err = decrypt_agile_encrypted_package(&encryption_info, &[], "pw").unwrap_err();
        assert!(
            matches!(err, OffCryptoError::InvalidBlockSize { block_size: 32 }),
            "unexpected error: {err:?}"
        );
    }

    #[test]
    fn rejects_salt_value_len_mismatch() {
        // Declares an 8-byte salt but provides a 16-byte saltValue.
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltValue="AAAAAAAAAAAAAAAAAAAAAA==" saltSize="8" hashAlgorithm="SHA1" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       keyBits="128" blockSize="16" />
              <dataIntegrity encryptedHmacKey="AA==" encryptedHmacValue="AA==" />
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey saltValue="AAAAAAAAAAAAAAAAAAAAAA==" saltSize="16" spinCount="1"
                                  hashAlgorithm="SHA1" hashSize="20"
                                  cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  keyBits="128" blockSize="16"
                                  encryptedVerifierHashInput="AA=="
                                  encryptedVerifierHashValue="AA=="
                                  encryptedKeyValue="AA=="/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let encryption_info = wrap_encryption_info(xml);
        let err = decrypt_agile_encrypted_package(&encryption_info, &[], "pw").unwrap_err();
        assert!(
            matches!(err, OffCryptoError::InvalidAttribute { .. }),
            "expected InvalidAttribute, got {err:?}"
        );
        let msg = err.to_string();
        assert!(
            msg.contains("saltSize") && msg.contains("8") && msg.contains("16"),
            "expected message to mention saltSize mismatch, got: {msg}"
        );
    }
}

fn hash_output_len(alg: HashAlgorithm) -> usize {
    match alg {
        HashAlgorithm::Sha1 => 20,
        HashAlgorithm::Sha256 => 32,
        HashAlgorithm::Sha384 => 48,
        HashAlgorithm::Sha512 => 64,
    }
}

fn validate_block_size(block_size: usize) -> Result<()> {
    if block_size != AES_BLOCK_SIZE {
        return Err(OffCryptoError::InvalidBlockSize { block_size });
    }
    Ok(())
}

fn validate_hash_size(
    node: roxmltree::Node<'_, '_>,
    attr: &'static str,
    hash_alg: HashAlgorithm,
    hash_size: usize,
) -> Result<()> {
    let expected = hash_output_len(hash_alg);
    if hash_size != expected {
        return Err(OffCryptoError::InvalidAttribute {
            element: node.tag_name().name().to_string(),
            attr: attr.to_string(),
            reason: format!(
                "hashSize must match hashAlgorithm output length ({expected}, got {hash_size})"
            ),
        });
    }
    Ok(())
}
fn required_attr<'a>(node: roxmltree::Node<'a, '_>, attr: &str) -> Result<&'a str> {
    node.attribute(attr)
        .ok_or_else(|| OffCryptoError::MissingRequiredAttribute {
            element: node.tag_name().name().to_string(),
            attr: attr.to_string(),
        })
}

fn parse_usize_attr(node: roxmltree::Node<'_, '_>, attr: &str) -> Result<usize> {
    let val = required_attr(node, attr)?;
    val.trim()
        .parse::<usize>()
        .map_err(|e| OffCryptoError::InvalidAttribute {
            element: node.tag_name().name().to_string(),
            attr: attr.to_string(),
            reason: e.to_string(),
        })
}

fn parse_u32_attr(node: roxmltree::Node<'_, '_>, attr: &str) -> Result<u32> {
    let val = required_attr(node, attr)?;
    val.trim()
        .parse::<u32>()
        .map_err(|e| OffCryptoError::InvalidAttribute {
            element: node.tag_name().name().to_string(),
            attr: attr.to_string(),
            reason: e.to_string(),
        })
}

fn parse_base64_attr(
    node: roxmltree::Node<'_, '_>,
    attr: &'static str,
    opts: &ParseOptions,
) -> Result<Vec<u8>> {
    let val = required_attr(node, attr)?;
    decode_base64_field_limited(node.tag_name().name(), attr, val, opts)
}

fn parse_base64_attr_or_child(
    node: roxmltree::Node<'_, '_>,
    field: &'static str,
    opts: &ParseOptions,
) -> Result<Vec<u8>> {
    // Prefer the attribute form for deterministic behavior when both are present.
    if let Some(raw) = node.attribute(field) {
        return decode_base64_field_limited(node.tag_name().name(), field, raw, opts);
    }

    // Some producers encode these blobs as child elements with base64 text content:
    //   <p:encryptedKey ...>
    //     <p:encryptedVerifierHashInput>...</p:encryptedVerifierHashInput>
    //     ...
    //   </p:encryptedKey>
    //
    // Match by local name so namespace prefixes don't matter.
    if let Some(child) = node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == field)
    {
        let raw = child.text().unwrap_or("");
        return decode_base64_field_limited(node.tag_name().name(), field, raw, opts);
    }

    // Preserve the original error semantics for missing ciphertext fields.
    Err(OffCryptoError::MissingRequiredAttribute {
        element: node.tag_name().name().to_string(),
        attr: field.to_string(),
    })
}

fn parse_hash_algorithm(node: roxmltree::Node<'_, '_>, attr: &str) -> Result<HashAlgorithm> {
    let val = required_attr(node, attr)?;
    HashAlgorithm::parse_offcrypto_name(val).map_err(|_| OffCryptoError::UnsupportedHashAlgorithm {
        hash: val.to_string(),
    })
}

fn key_len_bytes(key_bits: usize, element: &'static str, attr: &'static str) -> Result<usize> {
    if key_bits % 8 != 0 {
        return Err(OffCryptoError::InvalidAttribute {
            element: element.to_string(),
            attr: attr.to_string(),
            reason: "keyBits must be divisible by 8".to_string(),
        });
    }
    let key_len = key_bits / 8;
    if !matches!(key_len, 16 | 24 | 32) {
        return Err(OffCryptoError::InvalidAttribute {
            element: element.to_string(),
            attr: attr.to_string(),
            reason: format!(
                "unsupported keyBits value {key_bits} (expected 128, 192, or 256)"
            ),
        });
    }
    Ok(key_len)
}

fn derive_key_or_err(
    h: &[u8],
    block_key: &[u8],
    key_len: usize,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>> {
    derive_key(h, block_key, key_len, hash_alg).map_err(|e| OffCryptoError::InvalidAttribute {
        element: "crypto".to_string(),
        attr: "derive_key".to_string(),
        reason: e.to_string(),
    })
}

fn derive_iv_or_err(
    salt: &[u8],
    block_key: &[u8],
    iv_len: usize,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>> {
    derive_iv(salt, block_key, iv_len, hash_alg).map_err(|e| OffCryptoError::InvalidAttribute {
        element: "crypto".to_string(),
        attr: "derive_iv".to_string(),
        reason: e.to_string(),
    })
}

fn hash_bytes(alg: HashAlgorithm, data: &[u8]) -> Vec<u8> {
    match alg {
        HashAlgorithm::Sha1 => sha1::Sha1::digest(data).to_vec(),
        HashAlgorithm::Sha256 => sha2::Sha256::digest(data).to_vec(),
        HashAlgorithm::Sha384 => sha2::Sha384::digest(data).to_vec(),
        HashAlgorithm::Sha512 => sha2::Sha512::digest(data).to_vec(),
    }
}

fn compute_hmac(alg: HashAlgorithm, key: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    match alg {
        HashAlgorithm::Sha1 => {
            let mut mac: Hmac<sha1::Sha1> =
                Hmac::new_from_slice(key).map_err(|e| OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "encryptedHmacKey".to_string(),
                    reason: e.to_string(),
                })?;
            mac.update(data);
            Ok(mac.finalize().into_bytes().to_vec())
        }
        HashAlgorithm::Sha256 => {
            let mut mac: Hmac<sha2::Sha256> =
                Hmac::new_from_slice(key).map_err(|e| OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "encryptedHmacKey".to_string(),
                    reason: e.to_string(),
                })?;
            mac.update(data);
            Ok(mac.finalize().into_bytes().to_vec())
        }
        HashAlgorithm::Sha384 => {
            let mut mac: Hmac<sha2::Sha384> =
                Hmac::new_from_slice(key).map_err(|e| OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "encryptedHmacKey".to_string(),
                    reason: e.to_string(),
                })?;
            mac.update(data);
            Ok(mac.finalize().into_bytes().to_vec())
        }
        HashAlgorithm::Sha512 => {
            let mut mac: Hmac<sha2::Sha512> =
                Hmac::new_from_slice(key).map_err(|e| OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "encryptedHmacKey".to_string(),
                    reason: e.to_string(),
                })?;
            mac.update(data);
            Ok(mac.finalize().into_bytes().to_vec())
        }
    }
}

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

#[cfg(test)]
mod key_encryptor_tests {
    use super::*;

    fn build_encryption_info_stream(xml: &str) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&4u16.to_le_bytes()); // major
        out.extend_from_slice(&4u16.to_le_bytes()); // minor
        out.extend_from_slice(&0u32.to_le_bytes()); // flags (ignored by parser)
        out.extend_from_slice(xml.as_bytes());
        out
    }

    #[test]
    fn parses_password_key_encryptor_when_multiple_key_encryptors_present() {
        let xml = format!(
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"
                xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
              <keyData saltValue="AA==" saltSize="1" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       keyBits="128" blockSize="16" hashSize="20"/>
              <dataIntegrity encryptedHmacKey="AA==" encryptedHmacValue="AA=="/>
              <keyEncryptors>
                <keyEncryptor uri="{KEY_ENCRYPTOR_URI_CERTIFICATE}">
                  <c:encryptedKey/>
                </keyEncryptor>
                <keyEncryptor uri="{KEY_ENCRYPTOR_URI_PASSWORD}">
                  <p:encryptedKey saltValue="AA==" saltSize="1" hashAlgorithm="SHA1" spinCount="1" cipherAlgorithm="AES"
                                  cipherChaining="ChainingModeCBC" keyBits="128" blockSize="16" hashSize="20"
                                  encryptedVerifierHashInput="AA==" encryptedVerifierHashValue="AA==" encryptedKeyValue="AA=="/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>"#
        );

        let stream = build_encryption_info_stream(&xml);
        let info = parse_agile_encryption_info(&stream, &DecryptOptions::default(), None)
            .expect("parse should succeed");
        assert_eq!(info.password_key.spin_count, 1);
    }

    #[test]
    fn errors_when_password_key_encryptor_is_missing() {
        let xml = format!(
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
              <keyData saltValue="AA==" saltSize="1" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       keyBits="128" blockSize="16" hashSize="20"/>
              <dataIntegrity encryptedHmacKey="AA==" encryptedHmacValue="AA=="/>
              <keyEncryptors>
                <keyEncryptor uri="{KEY_ENCRYPTOR_URI_CERTIFICATE}">
                  <c:encryptedKey/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>"#
        );

        let stream = build_encryption_info_stream(&xml);
        let err = parse_agile_encryption_info(&stream, &DecryptOptions::default(), None)
            .expect_err("expected error");
        match err {
            OffCryptoError::UnsupportedKeyEncryptor { available_uris, .. } => {
                assert!(
                    available_uris
                        .iter()
                        .any(|u| u == KEY_ENCRYPTOR_URI_CERTIFICATE),
                    "expected certificate URI to be listed, got {available_uris:?}"
                );
            }
            other => panic!("expected UnsupportedKeyEncryptor, got {other:?}"),
        }
    }
}

#[cfg(all(test, not(target_arch = "wasm32")))]
mod fuzz_tests {
    #![allow(unexpected_cfgs)]

    use super::*;
    use proptest::prelude::*;
    use std::io::Cursor;
    use std::panic::{catch_unwind, AssertUnwindSafe};
    use std::sync::OnceLock;

    #[cfg(fuzzing)]
    const CASES: u32 = 512;
    #[cfg(not(fuzzing))]
    const CASES: u32 = 32;

    #[cfg(fuzzing)]
    const MAX_LEN: usize = 256 * 1024;
    #[cfg(not(fuzzing))]
    const MAX_LEN: usize = 32 * 1024;

    fn invalid_agile_encryption_info(mut tail: Vec<u8>) -> Vec<u8> {
        // Force the parser down the Agile branch (4.4) and ensure the XML slice is not UTF-8.
        //
        // This avoids flaky tests where randomly generated bytes accidentally form a valid
        // EncryptionInfo XML descriptor (extremely unlikely, but possible).
        //
        // Note: the Agile `EncryptionInfo` stream header is 8 bytes:
        // `major (u16le), minor (u16le), flags (u32le)`. The XML payload begins at byte offset 8.
        let mut out = Vec::with_capacity(8 + 2 + tail.len());
        out.extend_from_slice(&[0x04, 0x00, 0x04, 0x00]); // major=4, minor=4
        out.extend_from_slice(&0u32.to_le_bytes()); // flags
        out.push(b'<');
        out.push(0xFF); // invalid UTF-8
        out.append(&mut tail);
        out
    }

    fn valid_agile_encryption_info() -> &'static Vec<u8> {
        static CACHE: OnceLock<Vec<u8>> = OnceLock::new();
        CACHE.get_or_init(|| {
            use cfb::CompoundFile;
            use formula_office_crypto::{encrypt_package_to_ole, EncryptOptions, EncryptionScheme};
            use std::io::{Cursor, Read, Write};
            use zip::write::FileOptions;

            // Keep this tiny so CI runtime stays low (the password KDF runs per proptest case).
            let plain_zip = {
                let cursor = Cursor::new(Vec::new());
                let mut writer = zip::ZipWriter::new(cursor);
                writer
                    .start_file("hello.txt", FileOptions::<()>::default())
                    .expect("start zip file");
                writer.write_all(b"hello").expect("write zip contents");
                writer.finish().expect("finish zip").into_inner()
            };

            let password = "pw";
            let opts = EncryptOptions {
                scheme: EncryptionScheme::Agile,
                key_bits: 128,
                hash_algorithm: formula_office_crypto::HashAlgorithm::Sha1,
                spin_count: 1,
            };

            let ole_bytes = encrypt_package_to_ole(&plain_zip, password, opts).expect("encrypt");
            let mut ole = CompoundFile::open(Cursor::new(ole_bytes)).expect("open cfb");

            let mut buf = Vec::new();
            if let Ok(mut stream) = ole.open_stream("EncryptionInfo") {
                stream.read_to_end(&mut buf).expect("read EncryptionInfo");
                return buf;
            }
            let mut stream = ole
                .open_stream("/EncryptionInfo")
                .expect("open /EncryptionInfo");
            stream.read_to_end(&mut buf).expect("read /EncryptionInfo");
            buf
        })
    }

    proptest! {
        #![proptest_config(ProptestConfig {
            cases: CASES,
            max_shrink_iters: 0,
            .. ProptestConfig::default()
        })]

        #[test]
        fn parse_agile_encryption_info_is_panic_free_and_rejects_garbage(
            tail in prop::collection::vec(any::<u8>(), 0..=MAX_LEN),
        ) {
            let bytes = invalid_agile_encryption_info(tail);
            let outcome = catch_unwind(AssertUnwindSafe(|| {
                let opts = DecryptOptions::default();
                super::parse_agile_encryption_info(&bytes, &opts, None)
            }));
            prop_assert!(outcome.is_ok(), "parse_agile_encryption_info panicked");
            prop_assert!(outcome.unwrap().is_err(), "garbage input should not parse");

            // Also cover the public Agile parser (`offcrypto::agile`) to ensure both entry points are
            // panic-free on hostile inputs.
            let outcome = catch_unwind(AssertUnwindSafe(|| {
                crate::offcrypto::parse_agile_encryption_info_stream(&bytes)
            }));
            prop_assert!(
                outcome.is_ok(),
                "parse_agile_encryption_info_stream panicked"
            );
            prop_assert!(
                outcome.unwrap().is_err(),
                "garbage input should not parse via parse_agile_encryption_info_stream"
            );

            // Also cover the lightweight Agile XML parser (used for key-encryptor preflight
            // diagnostics).
            let outcome = catch_unwind(AssertUnwindSafe(|| {
                let opts = ParseOptions::default();
                let xml = extract_encryption_info_xml(&bytes, &opts)?;
                crate::offcrypto::parse_agile_encryption_info_xml(xml)
            }));
            prop_assert!(
                outcome.is_ok(),
                "parse_agile_encryption_info_xml panicked"
            );
            prop_assert!(
                outcome.unwrap().is_err(),
                "garbage input should not parse via parse_agile_encryption_info_xml"
            );
        }

        #[test]
        fn decrypt_agile_encrypted_package_is_panic_free_and_rejects_garbage(
            info_tail in prop::collection::vec(any::<u8>(), 0..=MAX_LEN),
            encrypted_package in prop::collection::vec(any::<u8>(), 0..=MAX_LEN),
        ) {
            let encryption_info = invalid_agile_encryption_info(info_tail);
            let outcome = catch_unwind(AssertUnwindSafe(|| {
                decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, "pw")
            }));
            prop_assert!(outcome.is_ok(), "decrypt_agile_encrypted_package panicked");
            prop_assert!(outcome.unwrap().is_err(), "garbage input should not decrypt");
        }

        #[test]
        fn decrypt_agile_encrypted_package_with_valid_info_is_panic_free_and_rejects_garbage_ciphertext(
            declared_len in any::<u64>(),
            mut ciphertext in prop::collection::vec(any::<u8>(), 0..=MAX_LEN),
        ) {
            // Ensure ciphertext (after the 8-byte original-size header) is AES-block aligned so
            // we exercise the full decrypt path instead of failing immediately.
            let new_len = ciphertext.len() - (ciphertext.len() % AES_BLOCK_SIZE);
            ciphertext.truncate(new_len);

            // Ensure `declared_len <= ciphertext.len()` so we reach the integrity/HMAC checks.
            let declared_len = if ciphertext.is_empty() {
                0u64
            } else {
                declared_len % (ciphertext.len() as u64 + 1)
            };

            let mut encrypted_package = Vec::with_capacity(8 + ciphertext.len());
            encrypted_package.extend_from_slice(&declared_len.to_le_bytes());
            encrypted_package.extend_from_slice(&ciphertext);

            let encryption_info = valid_agile_encryption_info();
            let outcome = catch_unwind(AssertUnwindSafe(|| {
                decrypt_agile_encrypted_package(encryption_info, &encrypted_package, "pw")
            }));
            prop_assert!(outcome.is_ok(), "decrypt_agile_encrypted_package panicked");
            prop_assert!(outcome.unwrap().is_err(), "garbage ciphertext should not decrypt");
        }

        #[test]
        fn decrypt_agile_encrypted_package_stream_with_valid_info_is_panic_free_and_rejects_garbage_ciphertext(
            len_matches in any::<bool>(),
            declared_len in any::<u64>(),
            mut ciphertext in prop::collection::vec(any::<u8>(), 0..=MAX_LEN),
        ) {
            // Ensure ciphertext (after the 8-byte original-size header) is AES-block aligned so
            // we exercise the full decrypt path instead of failing immediately.
            let new_len = ciphertext.len() - (ciphertext.len() % AES_BLOCK_SIZE);
            ciphertext.truncate(new_len);

            let declared_len = if len_matches {
                // Ensure `declared_len <= ciphertext.len()` so we reach the integrity/HMAC checks.
                if ciphertext.is_empty() {
                    0u64
                } else {
                    declared_len % (ciphertext.len() as u64 + 1)
                }
            } else {
                // Ensure `declared_len > ciphertext.len()` so we exercise the truncated-stream error
                // path without needing huge allocations.
                declared_len.saturating_add(ciphertext.len() as u64 + 1)
            };

            let mut encrypted_package = Vec::with_capacity(8 + ciphertext.len());
            encrypted_package.extend_from_slice(&declared_len.to_le_bytes());
            encrypted_package.extend_from_slice(&ciphertext);

            let encryption_info = valid_agile_encryption_info();
            let mut cursor = Cursor::new(encrypted_package);
            let mut out = Vec::new();
            let outcome = catch_unwind(AssertUnwindSafe(|| {
                decrypt_agile_encrypted_package_stream(encryption_info, &mut cursor, "pw", &mut out)
            }));
            prop_assert!(
                outcome.is_ok(),
                "decrypt_agile_encrypted_package_stream panicked"
            );
            prop_assert!(
                outcome.unwrap().is_err(),
                "garbage ciphertext should not decrypt via streaming API"
            );
        }
    }

    #[test]
    fn keydata_block_size_zero_is_rejected_without_panicking() {
        use cfb::CompoundFile;
        use ms_offcrypto_writer::Ecma376AgileWriter;
        use rand::{rngs::StdRng, SeedableRng as _};
        use std::io::{Cursor, Read, Write};
        use zip::write::FileOptions;

        fn build_tiny_zip() -> Vec<u8> {
            let cursor = Cursor::new(Vec::new());
            let mut writer = zip::ZipWriter::new(cursor);
            writer
                .start_file("hello.txt", FileOptions::<()>::default())
                .expect("start zip file");
            writer.write_all(b"hello").expect("write zip contents");
            writer.finish().expect("finish zip").into_inner()
        }

        fn encrypt_zip_with_password(plain_zip: &[u8], password: &str) -> Vec<u8> {
            let mut cursor = Cursor::new(Vec::new());
            let mut rng = StdRng::from_seed([0u8; 32]);
            let mut agile = Ecma376AgileWriter::create(&mut rng, password, &mut cursor)
                .expect("create agile");
            agile
                .write_all(plain_zip)
                .expect("write plaintext zip to agile writer");
            agile.finalize().expect("finalize agile writer");
            cursor.into_inner()
        }

        fn extract_stream_bytes(cfb_bytes: &[u8], stream_name: &str) -> Vec<u8> {
            let mut ole = CompoundFile::open(Cursor::new(cfb_bytes)).expect("open cfb");
            let mut stream = ole.open_stream(stream_name).expect("open stream");
            let mut buf = Vec::new();
            stream.read_to_end(&mut buf).expect("read stream");
            buf
        }

        fn replace_tag_attr(xml: &str, tag: &str, attr: &str, new_value: &str) -> String {
            let start = xml
                .find(&format!("<{tag}"))
                .unwrap_or_else(|| panic!("missing <{tag}> tag"));
            let end = xml[start..]
                .find('>')
                .map(|i| start + i)
                .unwrap_or_else(|| panic!("unterminated <{tag}> start tag"));
            let head = &xml[..start];
            let tag_contents = &xml[start..end];
            let tail = &xml[end..];

            let needle = format!("{attr}=\"");
            let attr_pos = tag_contents
                .find(&needle)
                .unwrap_or_else(|| panic!("missing {attr} attribute on <{tag}>"));
            let value_start = attr_pos + needle.len();
            let value_end = tag_contents[value_start..]
                .find('"')
                .map(|i| value_start + i)
                .unwrap_or_else(|| panic!("unterminated {attr} attribute on <{tag}>"));

            format!(
                "{}{}{}{}{}",
                head,
                &tag_contents[..value_start],
                new_value,
                &tag_contents[value_end..],
                tail
            )
        }

        let password = "correct horse battery staple";
        let plain_zip = build_tiny_zip();

        let encrypted_cfb = encrypt_zip_with_password(&plain_zip, password);
        let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
        let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

        let xml_start = encryption_info
            .iter()
            .position(|b| *b == b'<')
            .expect("EncryptionInfo should contain XML");
        let xml = std::str::from_utf8(&encryption_info[xml_start..]).expect("fixture XML should be UTF-8");
        let xml = replace_tag_attr(xml, "keyData", "blockSize", "0");

        let mut modified = encryption_info[..xml_start].to_vec();
        modified.extend_from_slice(xml.as_bytes());

        let outcome = catch_unwind(AssertUnwindSafe(|| {
            decrypt_agile_encrypted_package(&modified, &encrypted_package, password)
        }));
        assert!(outcome.is_ok(), "decrypt_agile_encrypted_package panicked");
        let err = outcome.unwrap().expect_err("expected failure due to invalid blockSize");
        match err {
            OffCryptoError::InvalidBlockSize { block_size: 0 } => {}
            OffCryptoError::InvalidAttribute { element, attr, .. }
                if element == "keyData" && attr == "blockSize" => {}
            other => panic!(
                "expected InvalidBlockSize(0) or InvalidAttribute(keyData, blockSize), got {other:?}"
            ),
        }
    }
}
