//! MS-OFFCRYPTO Agile decryption for OOXML `EncryptedPackage`.

use base64::engine::general_purpose::{
    STANDARD as BASE64_STANDARD, STANDARD_NO_PAD as BASE64_STANDARD_NO_PAD,
};
use base64::Engine as _;
use digest::Digest as _;
use hmac::{Hmac, Mac};

use super::aes_cbc::decrypt_aes_cbc_no_padding;
use super::crypto::{
    derive_iv, derive_key, hash_password, segment_block_key, HashAlgorithm, HMAC_KEY_BLOCK,
    HMAC_VALUE_BLOCK, KEY_VALUE_BLOCK, VERIFIER_HASH_INPUT_BLOCK, VERIFIER_HASH_VALUE_BLOCK,
};
use super::error::{OffCryptoError, Result};

const SEGMENT_SIZE: usize = 0x1000;

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

    // The IV for the password key encryptor fields is the saltValue itself (truncated to blockSize).
    let verifier_iv = info
        .password_key
        .salt_value
        .get(..info.password_key.block_size)
        .ok_or_else(|| OffCryptoError::InvalidAttribute {
            element: "p:encryptedKey".to_string(),
            attr: "saltValue".to_string(),
            reason: "saltValue shorter than blockSize".to_string(),
        })?;

    let verifier_input = {
        let k = derive_key_or_err(
            &password_hash,
            &VERIFIER_HASH_INPUT_BLOCK,
            key_encrypt_key_len,
            info.password_key.hash_algorithm,
        )?;
        let decrypted = decrypt_aes_cbc_no_padding(
            &k,
            verifier_iv,
            &info.password_key.encrypted_verifier_hash_input,
        )
        .map_err(|e| OffCryptoError::InvalidAttribute {
            element: "p:encryptedKey".to_string(),
            attr: "encryptedVerifierHashInput".to_string(),
            reason: e.to_string(),
        })?;
        decrypted
            .get(..info.password_key.block_size)
            .ok_or_else(|| OffCryptoError::InvalidAttribute {
                element: "p:encryptedKey".to_string(),
                attr: "encryptedVerifierHashInput".to_string(),
                reason: "decrypted verifierHashInput shorter than blockSize".to_string(),
            })?
            .to_vec()
    };

    let verifier_hash = {
        let k = derive_key_or_err(
            &password_hash,
            &VERIFIER_HASH_VALUE_BLOCK,
            key_encrypt_key_len,
            info.password_key.hash_algorithm,
        )?;
        let decrypted = decrypt_aes_cbc_no_padding(
            &k,
            verifier_iv,
            &info.password_key.encrypted_verifier_hash_value,
        )
        .map_err(|e| OffCryptoError::InvalidAttribute {
            element: "p:encryptedKey".to_string(),
            attr: "encryptedVerifierHashValue".to_string(),
            reason: e.to_string(),
        })?;
        decrypted
            .get(..info.password_key.hash_size)
            .ok_or_else(|| OffCryptoError::InvalidAttribute {
                element: "p:encryptedKey".to_string(),
                attr: "encryptedVerifierHashValue".to_string(),
                reason: "decrypted verifierHashValue shorter than hashSize".to_string(),
            })?
            .to_vec()
    };

    let computed_verifier_hash_full = hash_bytes(info.password_key.hash_algorithm, &verifier_input);
    let computed_verifier_hash = computed_verifier_hash_full
        .get(..info.password_key.hash_size)
        .ok_or_else(|| OffCryptoError::InvalidAttribute {
            element: "p:encryptedKey".to_string(),
            attr: "hashAlgorithm".to_string(),
            reason: "hash output shorter than hashSize".to_string(),
        })?;
    if !ct_eq(computed_verifier_hash, &verifier_hash) {
        return Err(OffCryptoError::WrongPassword);
    }

    let key_value = {
        let k = derive_key_or_err(
            &password_hash,
            &KEY_VALUE_BLOCK,
            key_encrypt_key_len,
            info.password_key.hash_algorithm,
        )?;
        let decrypted =
            decrypt_aes_cbc_no_padding(&k, verifier_iv, &info.password_key.encrypted_key_value)
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

    // 2) Decrypt EncryptedPackage stream to plaintext ZIP bytes.
    let (declared_len, ciphertext) = parse_encrypted_package_stream(encrypted_package)?;

    if ciphertext.len() % info.key_data.block_size != 0 {
        return Err(OffCryptoError::CiphertextNotBlockAligned {
            ciphertext_len: ciphertext.len(),
            block_size: info.key_data.block_size,
        });
    }

    let mut plaintext = Vec::with_capacity(ciphertext.len());
    for (idx, chunk) in ciphertext.chunks(SEGMENT_SIZE).enumerate() {
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

    // MS-OFFCRYPTO "dataIntegrity" is computed over the full EncryptedPackage stream bytes
    // (length prefix + ciphertext). This matches the reference implementation used by Excel and
    // the `ms-offcrypto-writer` crate.
    let actual_hmac = compute_hmac(info.key_data.hash_algorithm, &hmac_key, encrypted_package)?;
    let actual_hmac = actual_hmac.get(..info.key_data.hash_size).ok_or_else(|| {
        OffCryptoError::InvalidAttribute {
            element: "dataIntegrity".to_string(),
            attr: "hashAlgorithm".to_string(),
            reason: "HMAC output shorter than hashSize".to_string(),
        }
    })?;

    if !ct_eq(actual_hmac, &expected_hmac) {
        return Err(OffCryptoError::IntegrityMismatch);
    }

    Ok(plaintext)
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
    let xml_start = encryption_info
        .iter()
        .position(|b| *b == b'<')
        .ok_or_else(|| OffCryptoError::MissingRequiredElement {
            element: "encryption".to_string(),
        })?;

    if encryption_info.len() < 4 {
        return Err(OffCryptoError::UnsupportedEncryptionVersion { major: 0, minor: 0 });
    }
    let major = u16::from_le_bytes([encryption_info[0], encryption_info[1]]);
    let minor = u16::from_le_bytes([encryption_info[2], encryption_info[3]]);
    if major != 4 || minor != 4 {
        return Err(OffCryptoError::UnsupportedEncryptionVersion { major, minor });
    }

    let xml_bytes = encryption_info.get(xml_start..).unwrap_or_default();
    let xml = std::str::from_utf8(xml_bytes)?;
    let doc = roxmltree::Document::parse(xml)?;

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

    let key_encryptor_node = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "keyEncryptor"
                && n.attribute("uri")
                    .is_some_and(|u| u.to_ascii_lowercase().contains("password"))
        })
        .ok_or_else(|| OffCryptoError::MissingRequiredElement {
            element: "keyEncryptor (password)".to_string(),
        })?;
    let encrypted_key_node = key_encryptor_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "encryptedKey")
        .ok_or_else(|| OffCryptoError::MissingRequiredElement {
            element: "encryptedKey".to_string(),
        })?;

    let key_data = parse_key_data(key_data_node)?;
    let data_integrity = parse_data_integrity(data_integrity_node)?;
    let password_key = parse_password_key_encryptor(encrypted_key_node)?;

    Ok(AgileEncryptionInfo {
        key_data,
        data_integrity,
        password_key,
    })
}

fn parse_key_data(node: roxmltree::Node<'_, '_>) -> Result<KeyData> {
    validate_cipher_settings(node)?;

    Ok(KeyData {
        salt_value: parse_base64_attr(node, "saltValue")?,
        hash_algorithm: parse_hash_algorithm(node, "hashAlgorithm")?,
        block_size: parse_usize_attr(node, "blockSize")?,
        key_bits: parse_usize_attr(node, "keyBits")?,
        hash_size: parse_usize_attr(node, "hashSize")?,
    })
}

fn parse_data_integrity(node: roxmltree::Node<'_, '_>) -> Result<DataIntegrity> {
    Ok(DataIntegrity {
        encrypted_hmac_key: parse_base64_attr(node, "encryptedHmacKey")?,
        encrypted_hmac_value: parse_base64_attr(node, "encryptedHmacValue")?,
    })
}

fn parse_password_key_encryptor(node: roxmltree::Node<'_, '_>) -> Result<PasswordKeyEncryptor> {
    validate_cipher_settings(node)?;

    Ok(PasswordKeyEncryptor {
        salt_value: parse_base64_attr(node, "saltValue")?,
        hash_algorithm: parse_hash_algorithm(node, "hashAlgorithm")?,
        spin_count: parse_u32_attr(node, "spinCount")?,
        block_size: parse_usize_attr(node, "blockSize")?,
        key_bits: parse_usize_attr(node, "keyBits")?,
        hash_size: parse_usize_attr(node, "hashSize")?,
        encrypted_verifier_hash_input: parse_base64_attr(node, "encryptedVerifierHashInput")?,
        encrypted_verifier_hash_value: parse_base64_attr(node, "encryptedVerifierHashValue")?,
        encrypted_key_value: parse_base64_attr(node, "encryptedKeyValue")?,
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

fn parse_base64_attr(node: roxmltree::Node<'_, '_>, attr: &str) -> Result<Vec<u8>> {
    let val = required_attr(node, attr)?;
    decode_b64_attr(val, node.tag_name().name(), attr)
}

fn decode_b64_attr(value: &str, element: &str, attr: &str) -> Result<Vec<u8>> {
    let bytes = value.as_bytes();

    // Most inputs are already compact; only allocate if we see whitespace.
    let mut cleaned: Option<Vec<u8>> = None;
    for (idx, &b) in bytes.iter().enumerate() {
        if matches!(b, b'\r' | b'\n' | b'\t' | b' ') {
            let mut out = Vec::with_capacity(bytes.len());
            out.extend_from_slice(&bytes[..idx]);
            for &b2 in &bytes[idx..] {
                if !matches!(b2, b'\r' | b'\n' | b'\t' | b' ') {
                    out.push(b2);
                }
            }
            cleaned = Some(out);
            break;
        }
    }

    let input = cleaned.as_deref().unwrap_or(bytes);
    let decoded = BASE64_STANDARD
        .decode(input)
        .or_else(|_| BASE64_STANDARD_NO_PAD.decode(input))
        .map_err(|source| OffCryptoError::Base64Decode {
            element: element.to_string(),
            attr: attr.to_string(),
            source,
        })?;
    Ok(decoded)
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
