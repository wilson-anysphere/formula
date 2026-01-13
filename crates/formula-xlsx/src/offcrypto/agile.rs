#[cfg(test)]
use base64::engine::general_purpose::STANDARD as BASE64;
#[cfg(test)]
use base64::Engine as _;
use digest::Digest as _;

use crate::offcrypto::{
    decrypt_aes_cbc_no_padding_in_place, derive_iv, derive_key, hash_password, AesCbcDecryptError,
    decode_base64_field_limited, extract_encryption_info_xml, HashAlgorithm, OffCryptoError,
    ParseOptions, Result, HMAC_KEY_BLOCK, HMAC_VALUE_BLOCK, KEY_VALUE_BLOCK, VERIFIER_HASH_INPUT_BLOCK,
    VERIFIER_HASH_VALUE_BLOCK, AES_BLOCK_SIZE,
};
use super::encryption_info::decode_encryption_info_xml_text;

const OOXML_PASSWORD_KEY_ENCRYPTOR_URI: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
const OOXML_CERTIFICATE_KEY_ENCRYPTOR_URI: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate";

/// Non-fatal warnings surfaced while parsing an Agile `EncryptionInfo` XML descriptor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AgileEncryptionInfoWarning {
    /// Multiple password `<keyEncryptor>` entries were present.
    ///
    /// Resolution is deterministic: the first password key encryptor wins.
    MultiplePasswordKeyEncryptors { count: usize },
}

/// Parsed `<keyData>` parameters from an Agile Encryption descriptor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AgileKeyData {
    pub salt_value: Vec<u8>,
    pub hash_algorithm: HashAlgorithm,
    pub cipher_algorithm: String,
    pub cipher_chaining: String,
    pub key_bits: u32,
    pub block_size: u32,
    pub hash_size: u32,
}

/// Parsed `<dataIntegrity>` parameters from an Agile Encryption descriptor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AgileDataIntegrity {
    pub encrypted_hmac_key: Vec<u8>,
    pub encrypted_hmac_value: Vec<u8>,
}

/// Parsed password `<encryptedKey>` parameters from an Agile Encryption descriptor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AgilePasswordKeyEncryptor {
    pub salt_value: Vec<u8>,
    pub spin_count: u32,
    pub hash_algorithm: HashAlgorithm,
    pub cipher_algorithm: String,
    pub cipher_chaining: String,
    pub key_bits: u32,
    pub block_size: u32,
    pub hash_size: u32,
    pub encrypted_verifier_hash_input: Vec<u8>,
    pub encrypted_verifier_hash_value: Vec<u8>,
    pub encrypted_key_value: Vec<u8>,
}

/// Parsed Agile Encryption descriptor (MS-OFFCRYPTO).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AgileEncryptionInfo {
    pub key_data: AgileKeyData,
    pub data_integrity: Option<AgileDataIntegrity>,
    pub password_key_encryptor: AgilePasswordKeyEncryptor,
    pub warnings: Vec<AgileEncryptionInfoWarning>,
}

/// Decrypted key material from the Agile password key encryptor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AgileDecryptedKeys {
    /// The decrypted package encryption key (`keyValue`), used to decrypt the `EncryptedPackage` stream.
    pub package_key: Vec<u8>,
    /// Decrypted `dataIntegrity/encryptedHmacKey` (when present).
    pub hmac_key: Option<Vec<u8>>,
    /// Decrypted `dataIntegrity/encryptedHmacValue` (when present).
    pub hmac_value: Option<Vec<u8>>,
}

fn parse_required_attr<'a>(
    element: &str,
    node: roxmltree::Node<'a, 'a>,
    attr: &str,
) -> Result<&'a str> {
    node.attribute(attr).ok_or_else(|| OffCryptoError::MissingRequiredAttribute {
        element: element.to_string(),
        attr: attr.to_string(),
    })
}

fn parse_u32_attr(element: &str, node: roxmltree::Node<'_, '_>, attr: &str) -> Result<u32> {
    let raw = parse_required_attr(element, node, attr)?;
    raw.trim().parse::<u32>().map_err(|e| OffCryptoError::InvalidAttribute {
        element: element.to_string(),
        attr: attr.to_string(),
        reason: format!("expected u32, got {raw:?}: {e}"),
    })
}

fn decode_b64_attr(
    element: &'static str,
    node: roxmltree::Node<'_, '_>,
    attr: &'static str,
    opts: &ParseOptions,
) -> Result<Vec<u8>> {
    let raw = parse_required_attr(element, node, attr)?;
    decode_base64_field_limited(element, attr, raw, opts)
}

fn parse_hash_algorithm_attr(
    element: &str,
    node: roxmltree::Node<'_, '_>,
    attr: &str,
) -> Result<HashAlgorithm> {
    let raw = parse_required_attr(element, node, attr)?;
    HashAlgorithm::parse_offcrypto_name(raw).map_err(|_| OffCryptoError::UnsupportedHashAlgorithm {
        hash: raw.to_string(),
    })
}

fn validate_aes_cbc_params(
    element: &str,
    cipher_algorithm: &str,
    cipher_chaining: &str,
    key_bits: u32,
    block_size: u32,
) -> Result<()> {
    if !cipher_algorithm.trim().eq_ignore_ascii_case("AES") {
        return Err(OffCryptoError::UnsupportedCipherAlgorithm {
            cipher: cipher_algorithm.to_string(),
        });
    }
    if !cipher_chaining
        .trim()
        .eq_ignore_ascii_case("ChainingModeCBC")
    {
        return Err(OffCryptoError::UnsupportedCipherChaining {
            chaining: cipher_chaining.to_string(),
        });
    }

    let key_len = (key_bits / 8) as usize;
    if !matches!(key_len, 16 | 24 | 32) {
        return Err(OffCryptoError::InvalidAttribute {
            element: element.to_string(),
            attr: "keyBits".to_string(),
            reason: format!("unsupported AES keyBits={key_bits} (expected 128/192/256)"),
        });
    }
    if block_size as usize != AES_BLOCK_SIZE {
        return Err(OffCryptoError::InvalidAttribute {
            element: element.to_string(),
            attr: "blockSize".to_string(),
            reason: format!("unsupported AES blockSize={block_size} (expected {AES_BLOCK_SIZE})"),
        });
    }

    Ok(())
}

/// Parse an Agile Encryption `EncryptionInfo` stream (MS-OFFCRYPTO version 4.4).
///
/// The caller must pass the full `EncryptionInfo` stream bytes (including the version header).
pub fn parse_agile_encryption_info_stream(
    encryption_info_stream: &[u8],
) -> Result<AgileEncryptionInfo> {
    parse_agile_encryption_info_stream_with_options(encryption_info_stream, &ParseOptions::default())
}

/// Parse an Agile Encryption `EncryptionInfo` stream with explicit parsing limits.
pub fn parse_agile_encryption_info_stream_with_options(
    encryption_info_stream: &[u8],
    opts: &ParseOptions,
) -> Result<AgileEncryptionInfo> {
    if encryption_info_stream.len() < 8 {
        return Err(OffCryptoError::MissingRequiredElement {
            element: "EncryptionInfoHeader".to_string(),
        });
    }

    let major = u16::from_le_bytes([encryption_info_stream[0], encryption_info_stream[1]]);
    let minor = u16::from_le_bytes([encryption_info_stream[2], encryption_info_stream[3]]);
    if (major, minor) != (4, 4) {
        return Err(OffCryptoError::UnsupportedEncryptionVersion { major, minor });
    }

    let xml_bytes = extract_encryption_info_xml(encryption_info_stream, opts)?;
    let xml = decode_encryption_info_xml_text(xml_bytes)?;
    let doc = roxmltree::Document::parse(xml.as_ref())?;

    let key_data_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "keyData")
        .ok_or_else(|| OffCryptoError::MissingRequiredElement {
            element: "keyData".to_string(),
        })?;

    let key_data_cipher_algorithm = parse_required_attr("keyData", key_data_node, "cipherAlgorithm")?
        .to_string();
    let key_data_cipher_chaining = parse_required_attr("keyData", key_data_node, "cipherChaining")?
        .to_string();
    let key_data_key_bits = parse_u32_attr("keyData", key_data_node, "keyBits")?;
    let key_data_block_size = parse_u32_attr("keyData", key_data_node, "blockSize")?;
    validate_aes_cbc_params(
        "keyData",
        &key_data_cipher_algorithm,
        &key_data_cipher_chaining,
        key_data_key_bits,
        key_data_block_size,
    )?;

    let key_data = AgileKeyData {
        salt_value: decode_b64_attr("keyData", key_data_node, "saltValue", opts)?,
        hash_algorithm: parse_hash_algorithm_attr("keyData", key_data_node, "hashAlgorithm")?,
        cipher_algorithm: key_data_cipher_algorithm,
        cipher_chaining: key_data_cipher_chaining,
        key_bits: key_data_key_bits,
        block_size: key_data_block_size,
        hash_size: parse_u32_attr("keyData", key_data_node, "hashSize")?,
    };

    let data_integrity = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "dataIntegrity")
        .map(|node| -> Result<AgileDataIntegrity> {
            Ok(AgileDataIntegrity {
                encrypted_hmac_key: decode_b64_attr("dataIntegrity", node, "encryptedHmacKey", opts)?,
                encrypted_hmac_value: decode_b64_attr(
                    "dataIntegrity",
                    node,
                    "encryptedHmacValue",
                    opts,
                )?,
            })
        })
        .transpose()?;

    let key_encryptors_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "keyEncryptors")
        .ok_or_else(|| OffCryptoError::MissingRequiredElement {
            element: "keyEncryptors".to_string(),
        })?;

    // Office can emit multiple key encryptors (e.g. password + certificate). We currently support
    // only the password key encryptor, so select it deterministically while capturing which URIs
    // were present for actionable errors.
    let mut available_uris: Vec<String> = Vec::new();
    let mut selected_password_encryptor: Option<roxmltree::Node<'_, '_>> = None;
    let mut password_encryptor_count = 0usize;
    for key_encryptor in key_encryptors_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "keyEncryptor")
    {
        let uri = key_encryptor.attribute("uri").ok_or_else(|| OffCryptoError::MissingRequiredAttribute {
            element: "keyEncryptor".to_string(),
            attr: "uri".to_string(),
        })?;

        if !available_uris.iter().any(|u| u == uri) {
            available_uris.push(uri.to_string());
        }

        if uri.trim() == OOXML_PASSWORD_KEY_ENCRYPTOR_URI {
            password_encryptor_count += 1;
            if selected_password_encryptor.is_none() {
                selected_password_encryptor = Some(key_encryptor);
            }
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
            .any(|u| u == OOXML_CERTIFICATE_KEY_ENCRYPTOR_URI)
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

    let key_encryptor_cipher_algorithm =
        parse_required_attr("encryptedKey", encrypted_key_node, "cipherAlgorithm")?.to_string();
    let key_encryptor_cipher_chaining =
        parse_required_attr("encryptedKey", encrypted_key_node, "cipherChaining")?.to_string();
    let key_encryptor_key_bits = parse_u32_attr("encryptedKey", encrypted_key_node, "keyBits")?;
    let key_encryptor_block_size = parse_u32_attr("encryptedKey", encrypted_key_node, "blockSize")?;
    validate_aes_cbc_params(
        "encryptedKey",
        &key_encryptor_cipher_algorithm,
        &key_encryptor_cipher_chaining,
        key_encryptor_key_bits,
        key_encryptor_block_size,
    )?;

    let password_key_encryptor = AgilePasswordKeyEncryptor {
        salt_value: decode_b64_attr("encryptedKey", encrypted_key_node, "saltValue", opts)?,
        spin_count: parse_u32_attr("encryptedKey", encrypted_key_node, "spinCount")?,
        hash_algorithm: parse_hash_algorithm_attr("encryptedKey", encrypted_key_node, "hashAlgorithm")?,
        cipher_algorithm: key_encryptor_cipher_algorithm,
        cipher_chaining: key_encryptor_cipher_chaining,
        key_bits: key_encryptor_key_bits,
        block_size: key_encryptor_block_size,
        hash_size: parse_u32_attr("encryptedKey", encrypted_key_node, "hashSize")?,
        encrypted_verifier_hash_input: decode_b64_attr(
            "encryptedKey",
            encrypted_key_node,
            "encryptedVerifierHashInput",
            opts,
        )?,
        encrypted_verifier_hash_value: decode_b64_attr(
            "encryptedKey",
            encrypted_key_node,
            "encryptedVerifierHashValue",
            opts,
        )?,
        encrypted_key_value: decode_b64_attr("encryptedKey", encrypted_key_node, "encryptedKeyValue", opts)?,
    };

    let mut warnings = Vec::new();
    if password_encryptor_count > 1 {
        warnings.push(AgileEncryptionInfoWarning::MultiplePasswordKeyEncryptors {
            count: password_encryptor_count,
        });
    }

    Ok(AgileEncryptionInfo {
        key_data,
        data_integrity,
        password_key_encryptor,
        warnings,
    })
}

fn hash_bytes(hash_alg: HashAlgorithm, bytes: &[u8]) -> Vec<u8> {
    match hash_alg {
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

fn decrypt_key_encryptor_blob(
    password_hash: &[u8],
    key_encryptor: &AgilePasswordKeyEncryptor,
    block_key: &[u8],
    ciphertext: &[u8],
) -> Result<Vec<u8>> {
    let key_len = (key_encryptor.key_bits / 8) as usize;
    let iv_len = key_encryptor.block_size as usize;

    let key =
        derive_key(password_hash, block_key, key_len, key_encryptor.hash_algorithm).map_err(|e| match e {
            crate::offcrypto::CryptoError::UnsupportedHashAlgorithm(name) => {
                OffCryptoError::UnsupportedHashAlgorithm { hash: name }
            }
            crate::offcrypto::CryptoError::InvalidParameter(reason) => OffCryptoError::InvalidAttribute {
                element: "encryptedKey".to_string(),
                attr: "keyBits".to_string(),
                reason: reason.to_string(),
            },
        })?;
    let iv = derive_iv(
        &key_encryptor.salt_value,
        block_key,
        iv_len,
        key_encryptor.hash_algorithm,
    )
    .map_err(|e| match e {
        crate::offcrypto::CryptoError::UnsupportedHashAlgorithm(name) => {
            OffCryptoError::UnsupportedHashAlgorithm { hash: name }
        }
        crate::offcrypto::CryptoError::InvalidParameter(reason) => OffCryptoError::InvalidAttribute {
            element: "encryptedKey".to_string(),
            attr: "saltValue".to_string(),
            reason: reason.to_string(),
        },
    })?;

    let mut buf = ciphertext.to_vec();
    decrypt_aes_cbc_no_padding_in_place(&key, &iv, &mut buf).map_err(|err| match err {
        AesCbcDecryptError::UnsupportedKeyLength(key_len) => OffCryptoError::InvalidAttribute {
            element: "encryptedKey".to_string(),
            attr: "keyBits".to_string(),
            reason: format!("derived key length {key_len} is not a supported AES key size"),
        },
        AesCbcDecryptError::InvalidIvLength(iv_len) => OffCryptoError::InvalidAttribute {
            element: "encryptedKey".to_string(),
            attr: "blockSize".to_string(),
            reason: format!("derived IV length {iv_len} does not match AES block size"),
        },
        AesCbcDecryptError::InvalidCiphertextLength(ciphertext_len) => {
            OffCryptoError::CiphertextNotBlockAligned {
                field: "ciphertext",
                len: ciphertext_len,
            }
        }
    })?;

    Ok(buf)
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

/// Decrypt the password key-encryptor values and validate the password via verifier hashes.
pub fn decrypt_agile_keys(info: &AgileEncryptionInfo, password: &str) -> Result<AgileDecryptedKeys> {
    let key_encryptor = &info.password_key_encryptor;

    let password_hash = hash_password(
        password,
        &key_encryptor.salt_value,
        key_encryptor.spin_count,
        key_encryptor.hash_algorithm,
    )
    .map_err(|e| match e {
        crate::offcrypto::CryptoError::UnsupportedHashAlgorithm(name) => {
            OffCryptoError::UnsupportedHashAlgorithm { hash: name }
        }
        crate::offcrypto::CryptoError::InvalidParameter(reason) => OffCryptoError::InvalidAttribute {
            element: "encryptedKey".to_string(),
            attr: "saltValue".to_string(),
            reason: reason.to_string(),
        },
    })?;

    // Decrypt verifierHashInput and verifierHashValue for password verification.
    validate_ciphertext_block_aligned(
        "encryptedVerifierHashInput",
        &key_encryptor.encrypted_verifier_hash_input,
    )?;
    let verifier_hash_input = decrypt_key_encryptor_blob(
        &password_hash,
        key_encryptor,
        &VERIFIER_HASH_INPUT_BLOCK,
        &key_encryptor.encrypted_verifier_hash_input,
    )?;
    let verifier_hash_input = verifier_hash_input
        .get(..AES_BLOCK_SIZE)
        .ok_or_else(|| OffCryptoError::WrongPassword)?
        .to_vec();

    validate_ciphertext_block_aligned(
        "encryptedVerifierHashValue",
        &key_encryptor.encrypted_verifier_hash_value,
    )?;
    let verifier_hash_value = decrypt_key_encryptor_blob(
        &password_hash,
        key_encryptor,
        &VERIFIER_HASH_VALUE_BLOCK,
        &key_encryptor.encrypted_verifier_hash_value,
    )?;
    let verifier_hash_value = verifier_hash_value
        .get(..key_encryptor.hash_size as usize)
        .ok_or_else(|| OffCryptoError::WrongPassword)?
        .to_vec();

    let computed = hash_bytes(key_encryptor.hash_algorithm, &verifier_hash_input);
    if computed.get(..verifier_hash_value.len()) != Some(verifier_hash_value.as_slice()) {
        return Err(OffCryptoError::WrongPassword);
    }

    // Decrypt the package key (`keyValue`).
    validate_ciphertext_block_aligned("encryptedKeyValue", &key_encryptor.encrypted_key_value)?;
    let key_value = decrypt_key_encryptor_blob(
        &password_hash,
        key_encryptor,
        &KEY_VALUE_BLOCK,
        &key_encryptor.encrypted_key_value,
    )?;
    let package_key_len = (info.key_data.key_bits / 8) as usize;
    let package_key = key_value
        .get(..package_key_len)
        .ok_or_else(|| OffCryptoError::WrongPassword)?
        .to_vec();

    // Decrypt HMAC key/value (if present).
    let (hmac_key, hmac_value) = match info.data_integrity.as_ref() {
        Some(di) => {
            validate_ciphertext_block_aligned(
                "dataIntegrity.encryptedHmacKey",
                &di.encrypted_hmac_key,
            )?;
            let raw_key =
                decrypt_key_encryptor_blob(&password_hash, key_encryptor, &HMAC_KEY_BLOCK, &di.encrypted_hmac_key)?;
            validate_ciphertext_block_aligned(
                "dataIntegrity.encryptedHmacValue",
                &di.encrypted_hmac_value,
            )?;
            let raw_val = decrypt_key_encryptor_blob(
                &password_hash,
                key_encryptor,
                &HMAC_VALUE_BLOCK,
                &di.encrypted_hmac_value,
            )?;
            let hmac_len = info.key_data.hash_size as usize;
            (
                Some(raw_key.get(..hmac_len).unwrap_or(&raw_key).to_vec()),
                Some(raw_val.get(..hmac_len).unwrap_or(&raw_val).to_vec()),
            )
        }
        None => (None, None),
    };

    Ok(AgileDecryptedKeys {
        package_key,
        hmac_key,
        hmac_value,
    })
}

fn derive_segment_iv(
    salt: &[u8],
    segment_index: u32,
    hash_alg: HashAlgorithm,
    iv_len: usize,
) -> [u8; AES_BLOCK_SIZE] {
    debug_assert_eq!(
        iv_len, AES_BLOCK_SIZE,
        "Agile keyData.blockSize must match AES block size"
    );

    let idx = segment_index.to_le_bytes();
    let mut iv = [0u8; AES_BLOCK_SIZE];
    match hash_alg {
        HashAlgorithm::Sha1 => {
            let mut h = sha1::Sha1::new();
            h.update(salt);
            h.update(&idx);
            iv.copy_from_slice(&h.finalize()[..AES_BLOCK_SIZE]);
        }
        HashAlgorithm::Sha256 => {
            let mut h = sha2::Sha256::new();
            h.update(salt);
            h.update(&idx);
            iv.copy_from_slice(&h.finalize()[..AES_BLOCK_SIZE]);
        }
        HashAlgorithm::Sha384 => {
            let mut h = sha2::Sha384::new();
            h.update(salt);
            h.update(&idx);
            iv.copy_from_slice(&h.finalize()[..AES_BLOCK_SIZE]);
        }
        HashAlgorithm::Sha512 => {
            let mut h = sha2::Sha512::new();
            h.update(salt);
            h.update(&idx);
            iv.copy_from_slice(&h.finalize()[..AES_BLOCK_SIZE]);
        }
    }
    iv
}

/// Decrypt an MS-OFFCRYPTO Agile `EncryptedPackage` stream given the decrypted package key.
pub fn decrypt_agile_encrypted_package_stream_with_key(
    encrypted_package_stream: &[u8],
    key_data: &AgileKeyData,
    package_key: &[u8],
) -> Result<Vec<u8>> {
    if encrypted_package_stream.len() < 8 {
        return Err(OffCryptoError::EncryptedPackageTooShort {
            len: encrypted_package_stream.len(),
        });
    }

    let mut size_bytes = [0u8; 8];
    size_bytes.copy_from_slice(&encrypted_package_stream[..8]);
    const SEGMENT_LEN: usize = 0x1000;
    let declared_len_u64 = u64::from_le_bytes(size_bytes);

    let ciphertext = &encrypted_package_stream[8..];
    if ciphertext.len() % AES_BLOCK_SIZE != 0 {
        return Err(OffCryptoError::CiphertextNotBlockAligned {
            field: "EncryptedPackage",
            len: ciphertext.len(),
        });
    }

    // --- Guardrails for malicious `declared_len` ---
    //
    // `EncryptedPackage` stores the unencrypted package size (`declared_len`) separately from the
    // ciphertext bytes. A corrupt/malicious size can otherwise induce large allocations (OOM) or
    // panics in `Vec::with_capacity` on 64-bit targets.
    let plausible_max = (ciphertext.len() as u64).saturating_add(SEGMENT_LEN as u64);
    if declared_len_u64 > plausible_max {
        return Err(OffCryptoError::InvalidAttribute {
            element: "EncryptedPackage".to_string(),
            attr: "originalSize".to_string(),
            reason: format!(
                "orig_size {declared_len_u64} is implausibly large for ciphertext length {}",
                ciphertext.len()
            ),
        });
    }

    // Guardrail: ciphertext must be long enough to possibly contain `declared_len` bytes
    // (accounting for AES block padding).
    let expected_min_ciphertext_len = declared_len_u64
        .checked_add((AES_BLOCK_SIZE - 1) as u64)
        .and_then(|v| v.checked_div(AES_BLOCK_SIZE as u64))
        .and_then(|blocks| blocks.checked_mul(AES_BLOCK_SIZE as u64))
        .ok_or_else(|| OffCryptoError::InvalidAttribute {
            element: "EncryptedPackage".to_string(),
            attr: "originalSize".to_string(),
            reason: format!(
                "orig_size {declared_len_u64} is implausibly large for ciphertext length {}",
                ciphertext.len()
            ),
        })?;
    if (ciphertext.len() as u64) < expected_min_ciphertext_len {
        return Err(OffCryptoError::InvalidAttribute {
            element: "EncryptedPackage".to_string(),
            attr: "originalSize".to_string(),
            reason: format!(
                "ciphertext length {} is too short for declared orig_size {declared_len_u64}",
                ciphertext.len()
            ),
        });
    }

    let declared_len: usize = declared_len_u64
        .try_into()
        .map_err(|_| OffCryptoError::InvalidAttribute {
            element: "EncryptedPackage".to_string(),
            attr: "originalSize".to_string(),
            reason: format!("orig_size {declared_len_u64} does not fit into usize"),
        })?;

    // Decrypt segment-by-segment until we have produced `declared_len` bytes.
    let mut out = Vec::with_capacity(declared_len);
    let mut offset = 0usize;
    let mut segment_index: u32 = 0;
    while offset < ciphertext.len() && out.len() < declared_len {
        let remaining = ciphertext.len() - offset;
        let seg_len = remaining.min(SEGMENT_LEN);
        if seg_len % AES_BLOCK_SIZE != 0 {
            return Err(OffCryptoError::CiphertextNotBlockAligned {
                field: "EncryptedPackage",
                len: seg_len,
            });
        }

        let iv = derive_segment_iv(
            &key_data.salt_value,
            segment_index,
            key_data.hash_algorithm,
            key_data.block_size as usize,
        );
        let mut decrypted = ciphertext[offset..offset + seg_len].to_vec();
        decrypt_aes_cbc_no_padding_in_place(package_key, &iv, &mut decrypted).map_err(|err| match err {
            AesCbcDecryptError::UnsupportedKeyLength(key_len) => OffCryptoError::InvalidAttribute {
                element: "keyData".to_string(),
                attr: "keyBits".to_string(),
                reason: format!("derived key length {key_len} is not a supported AES key size"),
            },
            AesCbcDecryptError::InvalidIvLength(iv_len) => OffCryptoError::InvalidAttribute {
                element: "keyData".to_string(),
                attr: "blockSize".to_string(),
                reason: format!("derived IV length {iv_len} does not match AES block size"),
            },
            AesCbcDecryptError::InvalidCiphertextLength(ciphertext_len) => {
                OffCryptoError::CiphertextNotBlockAligned {
                    field: "EncryptedPackage",
                    len: ciphertext_len,
                }
            }
        })?;

        let remaining_needed = declared_len - out.len();
        if decrypted.len() > remaining_needed {
            out.extend_from_slice(&decrypted[..remaining_needed]);
            break;
        }

        out.extend_from_slice(&decrypted);
        offset += seg_len;
        segment_index = segment_index.saturating_add(1);
    }

    if out.len() < declared_len {
        return Err(OffCryptoError::DecryptedLengthShorterThanHeader {
            declared_len,
            available_len: out.len(),
        });
    }

    out.truncate(declared_len);
    Ok(out)
}

/// High-level helper: parse an Agile `EncryptionInfo` stream, verify `password`, and decrypt
/// `EncryptedPackage`.
pub fn decrypt_agile_encrypted_package_stream(
    encryption_info_stream: &[u8],
    encrypted_package_stream: &[u8],
    password: &str,
) -> Result<Vec<u8>> {
    let info = parse_agile_encryption_info_stream(encryption_info_stream)?;
    let keys = decrypt_agile_keys(&info, password)?;
    decrypt_agile_encrypted_package_stream_with_key(encrypted_package_stream, &info.key_data, &keys.package_key)
}

#[cfg(test)]
mod tests {
    use super::*;
    use aes::{Aes128, Aes192, Aes256};
    use cbc::cipher::block_padding::NoPadding;
    use cbc::cipher::{BlockEncryptMut, KeyIvInit};

    fn wrap_xml_in_encryption_info_stream(xml: &str) -> Vec<u8> {
        let mut encryption_info_stream = Vec::new();
        encryption_info_stream.extend_from_slice(&4u16.to_le_bytes()); // major
        encryption_info_stream.extend_from_slice(&4u16.to_le_bytes()); // minor
        encryption_info_stream.extend_from_slice(&0u32.to_le_bytes()); // flags
        encryption_info_stream.extend_from_slice(xml.as_bytes());
        encryption_info_stream
    }

    fn minimal_encryption_info_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
    xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="32"
           cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA256"
           saltValue="AAECAwQFBgcICQoLDA0ODw=="/>
  <dataIntegrity encryptedHmacKey="EBESEw==" encryptedHmacValue="qrvM"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey saltSize="16" blockSize="16" keyBits="256" hashSize="32"
                      spinCount="100000" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA256"
                      saltValue="AQIDBA=="
                      encryptedVerifierHashInput="CQoLDA=="
                      encryptedVerifierHashValue="DQ4PEA=="
                      encryptedKeyValue="BQYHCA=="/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>
"#
    }

    fn wrap_payload_in_encryption_info_stream(payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&4u16.to_le_bytes()); // major
        out.extend_from_slice(&4u16.to_le_bytes()); // minor
        out.extend_from_slice(&0u32.to_le_bytes()); // flags
        out.extend_from_slice(payload);
        out
    }

    fn parse_stream_payload(payload: &[u8]) -> AgileEncryptionInfo {
        let stream = wrap_payload_in_encryption_info_stream(payload);
        parse_agile_encryption_info_stream(&stream).expect("parse agile encryption info")
    }

    #[test]
    fn parses_agile_encryption_info_with_utf8_bom_and_trailing_nuls() {
        let xml = minimal_encryption_info_xml();
        let expected = parse_stream_payload(xml.as_bytes());

        let mut payload = Vec::new();
        payload.extend_from_slice(&[0xEF, 0xBB, 0xBF]); // UTF-8 BOM
        payload.extend_from_slice(xml.as_bytes());
        payload.extend_from_slice(&[0, 0, 0]);

        let parsed = parse_stream_payload(&payload);
        assert_eq!(parsed, expected);
    }

    #[test]
    fn parses_agile_encryption_info_with_utf16le_xml() {
        let xml = minimal_encryption_info_xml();
        let expected = parse_stream_payload(xml.as_bytes());

        let mut payload = Vec::new();
        // No BOM: rely on NUL-density heuristic.
        for unit in xml.encode_utf16() {
            payload.extend_from_slice(&unit.to_le_bytes());
        }
        // UTF-16 NUL terminator.
        payload.extend_from_slice(&[0x00, 0x00]);

        let parsed = parse_stream_payload(&payload);
        assert_eq!(parsed, expected);
    }

    #[test]
    fn parses_agile_encryption_info_with_length_prefix_and_trailing_garbage() {
        let xml = minimal_encryption_info_xml();
        let expected = parse_stream_payload(xml.as_bytes());

        let xml_bytes = xml.as_bytes();
        let mut payload = Vec::new();
        payload.extend_from_slice(&(xml_bytes.len() as u32).to_le_bytes());
        payload.extend_from_slice(xml_bytes);
        payload.extend_from_slice(b"GARBAGE");

        let parsed = parse_stream_payload(&payload);
        assert_eq!(parsed, expected);
    }

    #[test]
    fn parses_agile_encryption_info_with_leading_bytes_before_xml() {
        let xml = minimal_encryption_info_xml();
        let expected = parse_stream_payload(xml.as_bytes());

        let mut payload = Vec::new();
        payload.extend_from_slice(b"JUNK");
        payload.extend_from_slice(xml.as_bytes());

        let parsed = parse_stream_payload(&payload);
        assert_eq!(parsed, expected);
    }

    #[test]
    fn rejects_cfb_cipher_chaining_in_key_data() {
        let xml = format!(
            r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltValue="AA==" hashAlgorithm="SHA1" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCFB"
                       keyBits="128" blockSize="16" />
              <keyEncryptors>
                <keyEncryptor uri="{OOXML_PASSWORD_KEY_ENCRYPTOR_URI}">
                  <p:encryptedKey saltValue="AA==" spinCount="1" hashAlgorithm="SHA1" hashSize="20"
                                  cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  keyBits="128" blockSize="16"
                                  encryptedVerifierHashInput="AA=="
                                  encryptedVerifierHashValue="AA=="
                                  encryptedKeyValue="AA=="/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#
        );

        let stream = wrap_xml_in_encryption_info_stream(&xml);
        let err = parse_agile_encryption_info_stream(&stream).expect_err("expected error");
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
        let xml = format!(
            r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltValue="AA==" hashAlgorithm="SHA1" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       keyBits="128" blockSize="16" />
              <keyEncryptors>
                <keyEncryptor uri="{OOXML_PASSWORD_KEY_ENCRYPTOR_URI}">
                  <p:encryptedKey saltValue="AA==" spinCount="1" hashAlgorithm="SHA1" hashSize="20"
                                  cipherAlgorithm="AES" cipherChaining="ChainingModeCFB"
                                  keyBits="128" blockSize="16"
                                  encryptedVerifierHashInput="AA=="
                                  encryptedVerifierHashValue="AA=="
                                  encryptedKeyValue="AA=="/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#
        );

        let stream = wrap_xml_in_encryption_info_stream(&xml);
        let err = parse_agile_encryption_info_stream(&stream).expect_err("expected error");
        assert!(
            matches!(err, OffCryptoError::UnsupportedCipherChaining { ref chaining } if chaining == "ChainingModeCFB"),
            "unexpected error: {err:?}"
        );
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

    fn encrypt_aes_cbc_no_padding(key: &[u8], iv: &[u8], plaintext: &[u8]) -> Vec<u8> {
        assert!(
            plaintext.len() % AES_BLOCK_SIZE == 0,
            "plaintext must be block-aligned"
        );
        let mut buf = plaintext.to_vec();
        let len = buf.len();
        match key.len() {
            16 => {
                cbc::Encryptor::<Aes128>::new_from_slices(key, iv)
                    .unwrap()
                    .encrypt_padded_mut::<NoPadding>(&mut buf, len)
                    .unwrap();
            }
            24 => {
                cbc::Encryptor::<Aes192>::new_from_slices(key, iv)
                    .unwrap()
                    .encrypt_padded_mut::<NoPadding>(&mut buf, len)
                    .unwrap();
            }
            32 => {
                cbc::Encryptor::<Aes256>::new_from_slices(key, iv)
                    .unwrap()
                    .encrypt_padded_mut::<NoPadding>(&mut buf, len)
                    .unwrap();
            }
            _ => panic!("unsupported AES key length"),
        }
        buf
    }

    #[test]
    fn agile_roundtrip_decrypts_key_blobs_and_package_segments() {
        // Synthetic Agile Encryption descriptor (not a real OOXML ZIP) to validate end-to-end
        // decryption logic and ensure AES-CBC no-padding is used consistently.
        let password = "password";
        let wrong_password = "not-the-password";

        // keyData (package encryption parameters).
        let key_data_salt = (0u8..=15).collect::<Vec<_>>();
        let key_data_key_bits = 128u32;
        let key_data_block_size = 16u32;
        let key_data_hash_alg = HashAlgorithm::Sha1;
        let key_data_hash_size = 20u32;

        // password key encryptor parameters.
        let ke_salt = (16u8..=31).collect::<Vec<_>>();
        let ke_spin = 10u32;
        let ke_key_bits = 128u32;
        let ke_block_size = 16u32;
        let ke_hash_alg = HashAlgorithm::Sha1;
        let ke_hash_size = 20u32;

        // Generate a deterministic package key and plaintext.
        let package_key = b"0123456789ABCDEF".to_vec(); // 16 bytes
        let plaintext = (0..5000u32).map(|i| (i % 251) as u8).collect::<Vec<_>>();

        // --- Encrypt EncryptedPackage stream (segment-wise) -----------------------------------
        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());

        let padded_plaintext = zero_pad(plaintext.clone());
        for (i, chunk) in padded_plaintext.chunks(0x1000).enumerate() {
            let iv = derive_segment_iv(&key_data_salt, i as u32, key_data_hash_alg, AES_BLOCK_SIZE);
            let ct = encrypt_aes_cbc_no_padding(&package_key, &iv, chunk);
            encrypted_package.extend_from_slice(&ct);
        }

        // --- Encrypt password key-encryptor blobs ---------------------------------------------
        let pw_hash = hash_password(password, &ke_salt, ke_spin, ke_hash_alg).unwrap();

        let verifier_hash_input = b"abcdefghijklmnop".to_vec(); // 16 bytes
        let verifier_hash_value = hash_bytes(ke_hash_alg, &verifier_hash_input); // 20 bytes for SHA1

        fn encrypt_ke_blob(
            pw_hash: &[u8],
            ke_salt: &[u8],
            ke_key_bits: u32,
            ke_block_size: u32,
            ke_hash_alg: HashAlgorithm,
            block_key: &[u8],
            plaintext: &[u8],
        ) -> Vec<u8> {
            let key_len = (ke_key_bits / 8) as usize;
            let iv_len = ke_block_size as usize;
            let key = derive_key(pw_hash, block_key, key_len, ke_hash_alg).unwrap();
            let iv = derive_iv(ke_salt, block_key, iv_len, ke_hash_alg).unwrap();
            let padded = zero_pad(plaintext.to_vec());
            encrypt_aes_cbc_no_padding(&key, &iv, &padded)
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

        // dataIntegrity blobs (we don't currently verify HMAC, but decrypting them exercises the same primitive).
        let hmac_key_plain = (100u8..120).collect::<Vec<_>>(); // 20 bytes
        let hmac_value_plain = (200u8..220).collect::<Vec<_>>(); // 20 bytes
        let encrypted_hmac_key = encrypt_ke_blob(
            &pw_hash,
            &ke_salt,
            ke_key_bits,
            ke_block_size,
            ke_hash_alg,
            &HMAC_KEY_BLOCK,
            &hmac_key_plain,
        );
        let encrypted_hmac_value = encrypt_ke_blob(
            &pw_hash,
            &ke_salt,
            ke_key_bits,
            ke_block_size,
            ke_hash_alg,
            &HMAC_VALUE_BLOCK,
            &hmac_value_plain,
        );

        // Build the EncryptionInfo XML.
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="{key_data_block_size}" keyBits="{key_data_key_bits}" hashSize="{key_data_hash_size}"
           cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
           saltValue="{key_data_salt_b64}"/>
  <dataIntegrity encryptedHmacKey="{ehk_b64}" encryptedHmacValue="{ehv_b64}"/>
  <keyEncryptors>
    <keyEncryptor uri="{OOXML_PASSWORD_KEY_ENCRYPTOR_URI}">
      <p:encryptedKey saltSize="16" blockSize="{ke_block_size}" keyBits="{ke_key_bits}" hashSize="{ke_hash_size}"
                      spinCount="{ke_spin}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                      saltValue="{ke_salt_b64}"
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
            ehk_b64 = BASE64.encode(&encrypted_hmac_key),
            ehv_b64 = BASE64.encode(&encrypted_hmac_value),
        );

        let mut encryption_info_stream = Vec::new();
        encryption_info_stream.extend_from_slice(&4u16.to_le_bytes()); // major
        encryption_info_stream.extend_from_slice(&4u16.to_le_bytes()); // minor
        encryption_info_stream.extend_from_slice(&0u32.to_le_bytes()); // flags
        encryption_info_stream.extend_from_slice(xml.as_bytes());

        // Parse + decrypt.
        let parsed = parse_agile_encryption_info_stream(&encryption_info_stream).expect("parse");
        assert_eq!(parsed.key_data.block_size, key_data_block_size);

        let keys = decrypt_agile_keys(&parsed, password).expect("decrypt keys");
        assert_eq!(keys.package_key, package_key);
        assert_eq!(keys.hmac_key.as_deref(), Some(hmac_key_plain.as_slice()));
        assert_eq!(keys.hmac_value.as_deref(), Some(hmac_value_plain.as_slice()));

        let decrypted =
            decrypt_agile_encrypted_package_stream(&encryption_info_stream, &encrypted_package, password)
                .expect("decrypt package");
        assert_eq!(decrypted, plaintext);

        let err = decrypt_agile_encrypted_package_stream(
            &encryption_info_stream,
            &encrypted_package,
            wrong_password,
        )
        .expect_err("wrong password should fail");
        assert!(
            matches!(err, OffCryptoError::WrongPassword),
            "expected WrongPassword, got {err:?}"
        );
    }

    fn build_encryption_info_stream(xml: &str) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&4u16.to_le_bytes()); // major
        out.extend_from_slice(&4u16.to_le_bytes()); // minor
        out.extend_from_slice(&0u32.to_le_bytes()); // flags
        out.extend_from_slice(xml.as_bytes());
        out
    }

    #[test]
    fn selects_password_key_encryptor_when_multiple_present() {
        let xml = format!(
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"
                xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
              <keyData saltSize="16" blockSize="16" keyBits="128" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                       saltValue=""/>
              <keyEncryptors>
                <keyEncryptor uri="{OOXML_CERTIFICATE_KEY_ENCRYPTOR_URI}">
                  <c:encryptedKey/>
                </keyEncryptor>
                <keyEncryptor uri="{OOXML_PASSWORD_KEY_ENCRYPTOR_URI}">
                  <p:encryptedKey saltSize="16" blockSize="16" keyBits="128" hashSize="20"
                                  spinCount="99" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  hashAlgorithm="SHA1" saltValue=""
                                  encryptedVerifierHashInput="" encryptedVerifierHashValue=""
                                  encryptedKeyValue=""/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>"#
        );

        let stream = build_encryption_info_stream(&xml);
        let info = parse_agile_encryption_info_stream(&stream).expect("parse should succeed");
        assert_eq!(info.password_key_encryptor.spin_count, 99);
        assert!(info.warnings.is_empty());
    }

    #[test]
    fn errors_when_password_key_encryptor_missing() {
        let xml = format!(
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
              <keyData saltSize="16" blockSize="16" keyBits="128" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                       saltValue=""/>
              <keyEncryptors>
                <keyEncryptor uri="{OOXML_CERTIFICATE_KEY_ENCRYPTOR_URI}">
                  <c:encryptedKey/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>"#
        );

        let stream = build_encryption_info_stream(&xml);
        let err = parse_agile_encryption_info_stream(&stream).expect_err("expected error");
        match err {
            OffCryptoError::UnsupportedKeyEncryptor { available_uris, .. } => {
                assert!(
                    available_uris
                        .iter()
                        .any(|u| u == OOXML_CERTIFICATE_KEY_ENCRYPTOR_URI),
                    "expected certificate URI to be listed, got {available_uris:?}"
                );
            }
            other => panic!("expected UnsupportedKeyEncryptor, got {other:?}"),
        }
    }

    #[test]
    fn warns_on_multiple_password_key_encryptors() {
        let xml = format!(
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltSize="16" blockSize="16" keyBits="128" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                       saltValue=""/>
              <keyEncryptors>
                <keyEncryptor uri="{OOXML_PASSWORD_KEY_ENCRYPTOR_URI}">
                  <p:encryptedKey saltSize="16" blockSize="16" keyBits="128" hashSize="20"
                                  spinCount="1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  hashAlgorithm="SHA1" saltValue=""
                                  encryptedVerifierHashInput="" encryptedVerifierHashValue=""
                                  encryptedKeyValue=""/>
                </keyEncryptor>
                <keyEncryptor uri="{OOXML_PASSWORD_KEY_ENCRYPTOR_URI}">
                  <p:encryptedKey saltSize="16" blockSize="16" keyBits="128" hashSize="20"
                                  spinCount="2" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  hashAlgorithm="SHA1" saltValue=""
                                  encryptedVerifierHashInput="" encryptedVerifierHashValue=""
                                  encryptedKeyValue=""/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>"#
        );

        let stream = build_encryption_info_stream(&xml);
        let info = parse_agile_encryption_info_stream(&stream).expect("parse should succeed");
        assert_eq!(info.password_key_encryptor.spin_count, 1);
        assert_eq!(
            info.warnings,
            vec![AgileEncryptionInfoWarning::MultiplePasswordKeyEncryptors { count: 2 }]
        );
    }

    #[test]
    fn decrypt_agile_encrypted_package_rejects_implausible_orig_size_without_panic() {
        let key_data = AgileKeyData {
            salt_value: vec![0u8; 16],
            hash_algorithm: HashAlgorithm::Sha1,
            cipher_algorithm: "AES".to_string(),
            cipher_chaining: "ChainingModeCBC".to_string(),
            key_bits: 128,
            block_size: 16,
            hash_size: 20,
        };
        let package_key = [0u8; 16];

        let mut stream = Vec::new();
        stream.extend_from_slice(&u64::MAX.to_le_bytes());
        stream.extend_from_slice(&[0u8; AES_BLOCK_SIZE]); // 1 AES block of ciphertext

        let err = decrypt_agile_encrypted_package_stream_with_key(&stream, &key_data, &package_key)
            .expect_err("expected error");

        assert!(
            matches!(
                err,
                OffCryptoError::InvalidAttribute { ref element, ref attr, .. }
                    if element == "EncryptedPackage" && attr == "originalSize"
            ),
            "unexpected error: {err:?}"
        );
    }

    #[test]
    fn decrypt_agile_encrypted_package_rejects_orig_size_near_u64_max_without_overflow() {
        let key_data = AgileKeyData {
            salt_value: vec![0u8; 16],
            hash_algorithm: HashAlgorithm::Sha1,
            cipher_algorithm: "AES".to_string(),
            cipher_chaining: "ChainingModeCBC".to_string(),
            key_bits: 128,
            block_size: 16,
            hash_size: 20,
        };
        let package_key = [0u8; 16];

        let mut stream = Vec::new();
        stream.extend_from_slice(&(u64::MAX - 4094).to_le_bytes());
        stream.extend_from_slice(&[0u8; AES_BLOCK_SIZE]);

        let err = decrypt_agile_encrypted_package_stream_with_key(&stream, &key_data, &package_key)
            .expect_err("expected error");

        assert!(
            matches!(
                err,
                OffCryptoError::InvalidAttribute { ref element, ref attr, .. }
                    if element == "EncryptedPackage" && attr == "originalSize"
            ),
            "unexpected error: {err:?}"
        );
    }
}
