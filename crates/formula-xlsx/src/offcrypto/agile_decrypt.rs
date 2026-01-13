//! MS-OFFCRYPTO Agile decryption for OOXML `EncryptedPackage`.

use digest::Digest as _;
use hmac::{Hmac, Mac};

use super::aes_cbc::{decrypt_aes_cbc_no_padding, AES_BLOCK_SIZE};
use super::crypto::{
    derive_iv, derive_key, hash_password, segment_block_key, HashAlgorithm, HMAC_KEY_BLOCK,
    HMAC_VALUE_BLOCK, KEY_VALUE_BLOCK, VERIFIER_HASH_INPUT_BLOCK, VERIFIER_HASH_VALUE_BLOCK,
};
use super::encryption_info::{
    decode_base64_field_limited, decode_encryption_info_xml_text, extract_encryption_info_xml,
    ParseOptions,
};
use super::error::{OffCryptoError, Result};

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
    data_integrity: DataIntegrity,
    password_key: PasswordKeyEncryptor,
}

#[derive(Debug, Clone, Copy)]
enum PasswordKeyIvDerivation {
    /// Use the password key encryptor `saltValue` directly as the AES-CBC IV (truncated to
    /// `blockSize`). This matches the behavior of several implementations and is what we
    /// historically supported.
    SaltValue,
    /// Derive the IV per block key: `IV = Truncate(Hash(saltValue || blockKey), blockSize)`.
    ///
    /// Some producers appear to use this scheme for the password-key-encryptor blobs
    /// (`encryptedVerifierHashInput`, `encryptedVerifierHashValue`, `encryptedKeyValue`).
    Derived,
}

fn decrypt_agile_package_key_from_password(
    info: &AgileEncryptionInfo,
    password_hash: &[u8],
    key_encrypt_key_len: usize,
    package_key_len: usize,
    iv_derivation: PasswordKeyIvDerivation,
) -> Result<Vec<u8>> {
    let password_key = &info.password_key;

    let password_iv_for = |block_key: &[u8]| -> Result<Vec<u8>> {
        match iv_derivation {
            PasswordKeyIvDerivation::SaltValue => password_key
                .salt_value
                .get(..password_key.block_size)
                .ok_or_else(|| OffCryptoError::InvalidAttribute {
                    element: "p:encryptedKey".to_string(),
                    attr: "saltValue".to_string(),
                    reason: "saltValue shorter than blockSize".to_string(),
                })
                .map(|iv| iv.to_vec()),
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
        let k = derive_key_or_err(
            password_hash,
            &VERIFIER_HASH_INPUT_BLOCK,
            key_encrypt_key_len,
            password_key.hash_algorithm,
        )?;
        let iv = password_iv_for(&VERIFIER_HASH_INPUT_BLOCK)?;
        let decrypted = decrypt_aes_cbc_no_padding(&k, &iv, &password_key.encrypted_verifier_hash_input)
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
        let k = derive_key_or_err(
            password_hash,
            &VERIFIER_HASH_VALUE_BLOCK,
            key_encrypt_key_len,
            password_key.hash_algorithm,
        )?;
        let iv = password_iv_for(&VERIFIER_HASH_VALUE_BLOCK)?;
        let decrypted = decrypt_aes_cbc_no_padding(&k, &iv, &password_key.encrypted_verifier_hash_value)
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
        let k = derive_key_or_err(
            password_hash,
            &KEY_VALUE_BLOCK,
            key_encrypt_key_len,
            password_key.hash_algorithm,
        )?;
        let iv = password_iv_for(&KEY_VALUE_BLOCK)?;
        let decrypted =
            decrypt_aes_cbc_no_padding(&k, &iv, &password_key.encrypted_key_value).map_err(|e| {
                OffCryptoError::InvalidAttribute {
                    element: "p:encryptedKey".to_string(),
                    attr: "encryptedKeyValue".to_string(),
                    reason: e.to_string(),
                }
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

/// Decrypt an MS-OFFCRYPTO Agile `EncryptedPackage` stream (OOXML password protection).
///
/// Inputs are the raw bytes of the CFB streams:
/// - `EncryptionInfo`
/// - `EncryptedPackage`
///
/// Returns the decrypted OOXML package bytes (a ZIP file) after validating `dataIntegrity`.
pub fn decrypt_agile_encrypted_package(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> Result<Vec<u8>> {
    let info = parse_agile_encryption_info(encryption_info)?;

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
    validate_ciphertext_block_aligned(
        "dataIntegrity.encryptedHmacKey",
        &info.data_integrity.encrypted_hmac_key,
    )?;
    validate_ciphertext_block_aligned(
        "dataIntegrity.encryptedHmacValue",
        &info.data_integrity.encrypted_hmac_value,
    )?;

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

    // Some producers appear to vary how the AES-CBC IV is derived for the password-key-encryptor
    // fields. Be conservative and try both strategies, treating the verifier-hash mismatch as a
    // signal to fall back to the alternative IV derivation.
    let key_value = match decrypt_agile_package_key_from_password(
        &info,
        &password_hash,
        key_encrypt_key_len,
        package_key_len,
        PasswordKeyIvDerivation::SaltValue,
    ) {
        Ok(key) => key,
        Err(OffCryptoError::WrongPassword) => decrypt_agile_package_key_from_password(
            &info,
            &password_hash,
            key_encrypt_key_len,
            package_key_len,
            PasswordKeyIvDerivation::Derived,
        )?,
        Err(other) => return Err(other),
    };

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

    // 3) Validate dataIntegrity HMAC.
    let hmac_key = {
        let iv = derive_iv_or_err(
            &info.key_data.salt_value,
            &HMAC_KEY_BLOCK,
            info.key_data.block_size,
            info.key_data.hash_algorithm,
        )?;
        let decrypted =
            decrypt_aes_cbc_no_padding(&key_value, &iv, &info.data_integrity.encrypted_hmac_key)
                .map_err(|e| OffCryptoError::InvalidAttribute {
                    element: "dataIntegrity".to_string(),
                    attr: "encryptedHmacKey".to_string(),
                    reason: e.to_string(),
                })?;
        decrypted
            .get(..info.key_data.hash_size)
            .ok_or_else(|| OffCryptoError::InvalidAttribute {
                element: "dataIntegrity".to_string(),
                attr: "encryptedHmacKey".to_string(),
                reason: "decrypted HMAC key shorter than hashSize".to_string(),
            })?
            .to_vec()
    };

    let expected_hmac = {
        let iv = derive_iv_or_err(
            &info.key_data.salt_value,
            &HMAC_VALUE_BLOCK,
            info.key_data.block_size,
            info.key_data.hash_algorithm,
        )?;
        let decrypted =
            decrypt_aes_cbc_no_padding(&key_value, &iv, &info.data_integrity.encrypted_hmac_value)
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
        let actual_hmac_plaintext = compute_hmac(info.key_data.hash_algorithm, &hmac_key, &plaintext)?;
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

fn parse_encrypted_package_stream(encrypted_package: &[u8]) -> Result<(usize, &[u8])> {
    if encrypted_package.len() < 8 {
        return Err(OffCryptoError::EncryptedPackageTooShort {
            len: encrypted_package.len(),
        });
    }

    let len_bytes: [u8; 8] = encrypted_package[..8]
        .try_into()
        .expect("slice length already checked");
    let declared_len_u64 = u64::from_le_bytes(len_bytes);
    let declared_len =
        usize::try_from(declared_len_u64).map_err(|_| OffCryptoError::InvalidAttribute {
            element: "EncryptedPackage".to_string(),
            attr: "original_package_size".to_string(),
            reason: format!("declared size {declared_len_u64} does not fit in usize"),
        })?;

    Ok((declared_len, &encrypted_package[8..]))
}

fn parse_agile_encryption_info(encryption_info: &[u8]) -> Result<AgileEncryptionInfo> {
    let opts = ParseOptions::default();
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

    let xml_bytes = extract_encryption_info_xml(encryption_info, &opts)?;
    let xml = decode_encryption_info_xml_text(xml_bytes)?;
    let doc = roxmltree::Document::parse(xml.as_ref())?;

    let key_data_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "keyData")
        .ok_or_else(|| OffCryptoError::MissingRequiredElement {
            element: "keyData".to_string(),
        })?;
    let data_integrity_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "dataIntegrity")
        .ok_or_else(|| OffCryptoError::MissingRequiredElement {
            element: "dataIntegrity".to_string(),
        })?;

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
        let uri = key_encryptor.attribute("uri").ok_or_else(|| OffCryptoError::MissingRequiredAttribute {
            element: "keyEncryptor".to_string(),
            attr: "uri".to_string(),
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

    // `password_encryptor_count > 1` is unusual but legal; decryption remains deterministic by
    // selecting the first password entry. We currently don't surface warnings from this helper.
    let _ = password_encryptor_count;
    let encrypted_key_node = key_encryptor_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "encryptedKey")
        .ok_or_else(|| OffCryptoError::MissingRequiredElement {
            element: "encryptedKey".to_string(),
        })?;

    let key_data = parse_key_data(key_data_node, &opts)?;
    let data_integrity = parse_data_integrity(data_integrity_node, &opts)?;
    let password_key = parse_password_key_encryptor(encrypted_key_node, &opts)?;

    Ok(AgileEncryptionInfo {
        key_data,
        data_integrity,
        password_key,
    })
}

fn parse_key_data(node: roxmltree::Node<'_, '_>, opts: &ParseOptions) -> Result<KeyData> {
    validate_cipher_settings(node)?;

    Ok(KeyData {
        salt_value: parse_base64_attr(node, "saltValue", opts)?,
        hash_algorithm: parse_hash_algorithm(node, "hashAlgorithm")?,
        block_size: parse_usize_attr(node, "blockSize")?,
        key_bits: parse_usize_attr(node, "keyBits")?,
        hash_size: parse_usize_attr(node, "hashSize")?,
    })
}

fn parse_data_integrity(node: roxmltree::Node<'_, '_>, opts: &ParseOptions) -> Result<DataIntegrity> {
    Ok(DataIntegrity {
        encrypted_hmac_key: parse_base64_attr(node, "encryptedHmacKey", opts)?,
        encrypted_hmac_value: parse_base64_attr(node, "encryptedHmacValue", opts)?,
    })
}

fn parse_password_key_encryptor(
    node: roxmltree::Node<'_, '_>,
    opts: &ParseOptions,
) -> Result<PasswordKeyEncryptor> {
    validate_cipher_settings(node)?;

    Ok(PasswordKeyEncryptor {
        salt_value: parse_base64_attr(node, "saltValue", opts)?,
        hash_algorithm: parse_hash_algorithm(node, "hashAlgorithm")?,
        spin_count: parse_u32_attr(node, "spinCount")?,
        block_size: parse_usize_attr(node, "blockSize")?,
        key_bits: parse_usize_attr(node, "keyBits")?,
        hash_size: parse_usize_attr(node, "hashSize")?,
        encrypted_verifier_hash_input: parse_base64_attr(node, "encryptedVerifierHashInput", opts)?,
        encrypted_verifier_hash_value: parse_base64_attr(node, "encryptedVerifierHashValue", opts)?,
        encrypted_key_value: parse_base64_attr(node, "encryptedKeyValue", opts)?,
    })
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

    fn wrap_encryption_info(xml: &str) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&4u16.to_le_bytes()); // major
        out.extend_from_slice(&4u16.to_le_bytes()); // minor
        out.extend_from_slice(&0u32.to_le_bytes()); // flags
        out.extend_from_slice(xml.as_bytes());
        out
    }

    #[test]
    fn rejects_cfb_cipher_chaining_in_key_data() {
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltValue="AA==" hashAlgorithm="SHA1" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCFB"
                       keyBits="128" blockSize="16" />
              <dataIntegrity encryptedHmacKey="AA==" encryptedHmacValue="AA==" />
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey saltValue="AA==" spinCount="1" hashAlgorithm="SHA1" hashSize="20"
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
              <keyData saltValue="AA==" hashAlgorithm="SHA1" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       keyBits="128" blockSize="16" />
              <dataIntegrity encryptedHmacKey="AA==" encryptedHmacValue="AA==" />
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey saltValue="AA==" spinCount="1" hashAlgorithm="SHA1" hashSize="20"
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
    Ok(key_bits / 8)
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
    if a.len() != b.len() {
        return false;
    }
    let mut diff = 0u8;
    for (&x, &y) in a.iter().zip(b.iter()) {
        diff |= x ^ y;
    }
    diff == 0
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
              <keyData saltValue="" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       keyBits="128" blockSize="16" hashSize="20"/>
              <dataIntegrity encryptedHmacKey="" encryptedHmacValue=""/>
              <keyEncryptors>
                <keyEncryptor uri="{KEY_ENCRYPTOR_URI_CERTIFICATE}">
                  <c:encryptedKey/>
                </keyEncryptor>
                <keyEncryptor uri="{KEY_ENCRYPTOR_URI_PASSWORD}">
                  <p:encryptedKey saltValue="" hashAlgorithm="SHA1" spinCount="1" cipherAlgorithm="AES"
                                  cipherChaining="ChainingModeCBC" keyBits="128" blockSize="16" hashSize="20"
                                  encryptedVerifierHashInput="" encryptedVerifierHashValue="" encryptedKeyValue=""/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>"#
        );

        let stream = build_encryption_info_stream(&xml);
        let info = parse_agile_encryption_info(&stream).expect("parse should succeed");
        assert_eq!(info.password_key.spin_count, 1);
    }

    #[test]
    fn errors_when_password_key_encryptor_is_missing() {
        let xml = format!(
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
              <keyData saltValue="" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       keyBits="128" blockSize="16" hashSize="20"/>
              <dataIntegrity encryptedHmacKey="" encryptedHmacValue=""/>
              <keyEncryptors>
                <keyEncryptor uri="{KEY_ENCRYPTOR_URI_CERTIFICATE}">
                  <c:encryptedKey/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>"#
        );

        let stream = build_encryption_info_stream(&xml);
        let err = parse_agile_encryption_info(&stream).expect_err("expected error");
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
