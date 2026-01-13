//! Decryption helpers for Office-encrypted OOXML workbooks (OLE `EncryptionInfo` + `EncryptedPackage`).
//!
//! This module is behind the `encrypted-workbooks` feature because password-based decryption is
//! still landing.

use std::io;
use std::io::{Read, Seek};

use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine as _;
use roxmltree::Document;

use crate::encrypted_package_reader::{DecryptedPackageReader, EncryptionMethod};

use formula_xlsx::offcrypto::{
    decrypt_aes_cbc_no_padding_in_place, derive_key, hash_password, CryptoError, HashAlgorithm,
    KEY_VALUE_BLOCK, VERIFIER_HASH_INPUT_BLOCK, VERIFIER_HASH_VALUE_BLOCK,
};

#[derive(Debug, thiserror::Error)]
pub(crate) enum DecryptError {
    #[error("unsupported EncryptionInfo version {major}.{minor}")]
    UnsupportedVersion { major: u16, minor: u16 },
    #[error("invalid EncryptionInfo: {0}")]
    InvalidInfo(String),
    #[error("invalid password")]
    InvalidPassword,
    #[error(transparent)]
    Io(#[from] io::Error),
}

pub(crate) fn decrypted_package_reader<R: Read + Seek>(
    ciphertext_reader: R,
    plaintext_len: u64,
    encryption_info: &[u8],
    password: &str,
) -> Result<DecryptedPackageReader<R>, DecryptError> {
    if encryption_info.len() < 4 {
        return Err(DecryptError::InvalidInfo(
            "EncryptionInfo truncated (missing version header)".to_string(),
        ));
    }

    let major = u16::from_le_bytes([encryption_info[0], encryption_info[1]]);
    let minor = u16::from_le_bytes([encryption_info[2], encryption_info[3]]);

    match (major, minor) {
        (4, 4) => decrypted_package_reader_agile(
            ciphertext_reader,
            plaintext_len,
            encryption_info,
            password,
        ),
        (3, 2) => decrypted_package_reader_standard(
            ciphertext_reader,
            plaintext_len,
            encryption_info,
            password,
        ),
        _ => Err(DecryptError::UnsupportedVersion { major, minor }),
    }
}

fn decrypted_package_reader_standard<R: Read + Seek>(
    ciphertext_reader: R,
    plaintext_len: u64,
    encryption_info: &[u8],
    password: &str,
) -> Result<DecryptedPackageReader<R>, DecryptError> {
    use formula_offcrypto::{
        parse_encryption_info, standard_derive_key, standard_verify_key, EncryptionInfo,
        OffcryptoError, StandardEncryptionInfo,
    };

    let info = match parse_encryption_info(encryption_info) {
        Ok(EncryptionInfo::Standard {
            header, verifier, ..
        }) => StandardEncryptionInfo { header, verifier },
        Ok(EncryptionInfo::Unsupported { version }) => {
            return Err(DecryptError::UnsupportedVersion {
                major: version.major,
                minor: version.minor,
            })
        }
        Ok(EncryptionInfo::Agile { .. }) => {
            return Err(DecryptError::InvalidInfo(
                "expected Standard EncryptionInfo, got Agile".to_string(),
            ))
        }
        Err(OffcryptoError::UnsupportedVersion { major, minor }) => {
            return Err(DecryptError::UnsupportedVersion { major, minor })
        }
        Err(OffcryptoError::InvalidPassword) => return Err(DecryptError::InvalidPassword),
        Err(err) => {
            return Err(DecryptError::InvalidInfo(format!(
                "failed to parse Standard EncryptionInfo: {err}"
            )))
        }
    };

    let key = standard_derive_key(&info, password).map_err(|err| match err {
        OffcryptoError::InvalidPassword => DecryptError::InvalidPassword,
        other => DecryptError::InvalidInfo(format!("failed to derive Standard key: {other}")),
    })?;

    standard_verify_key(&info, &key).map_err(|err| match err {
        OffcryptoError::InvalidPassword => DecryptError::InvalidPassword,
        other => DecryptError::InvalidInfo(format!("failed to verify Standard key: {other}")),
    })?;

    Ok(DecryptedPackageReader::new(
        ciphertext_reader,
        EncryptionMethod::StandardCryptoApi {
            key,
            salt: info.verifier.salt,
        },
        plaintext_len,
    ))
}

#[derive(Debug, Clone)]
struct AgileKeyData {
    salt_value: Vec<u8>,
    hash_algorithm: HashAlgorithm,
    block_size: usize,
    key_bits: usize,
    hash_size: usize,
}

#[derive(Debug, Clone)]
struct AgilePasswordKeyEncryptor {
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
    key_data: AgileKeyData,
    password_key: AgilePasswordKeyEncryptor,
}

fn decrypted_package_reader_agile<R: Read + Seek>(
    ciphertext_reader: R,
    plaintext_len: u64,
    encryption_info: &[u8],
    password: &str,
) -> Result<DecryptedPackageReader<R>, DecryptError> {
    let xml = crate::extract_agile_encryption_info_xml(encryption_info)
        .map_err(|err| DecryptError::InvalidInfo(err.to_string()))?;
    let info = parse_agile_encryption_info(&xml)?;

    let key = agile_decrypt_package_key(password, &info)?;

    Ok(DecryptedPackageReader::new(
        ciphertext_reader,
        EncryptionMethod::Agile {
            key,
            salt: info.key_data.salt_value,
            hash_alg: info.key_data.hash_algorithm,
            block_size: info.key_data.block_size,
        },
        plaintext_len,
    ))
}

fn parse_agile_encryption_info(xml: &str) -> Result<AgileEncryptionInfo, DecryptError> {
    let doc = Document::parse(xml)
        .map_err(|err| DecryptError::InvalidInfo(format!("EncryptionInfo XML parse: {err}")))?;

    let key_data_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "keyData")
        .ok_or_else(|| DecryptError::InvalidInfo("missing keyData element".into()))?;

    validate_cipher_settings(key_data_node)?;

    let key_data = AgileKeyData {
        salt_value: parse_base64_attr(key_data_node, "saltValue")?,
        hash_algorithm: parse_hash_algorithm(key_data_node, "hashAlgorithm")?,
        block_size: parse_usize_attr(key_data_node, "blockSize")?,
        key_bits: parse_usize_attr(key_data_node, "keyBits")?,
        hash_size: parse_usize_attr(key_data_node, "hashSize")?,
    };

    let key_encryptor_node = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "keyEncryptor"
                && n.attribute("uri")
                    .is_some_and(|u| u.to_ascii_lowercase().contains("password"))
        })
        .ok_or_else(|| DecryptError::InvalidInfo("missing keyEncryptor (password)".into()))?;

    let encrypted_key_node = key_encryptor_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "encryptedKey")
        .ok_or_else(|| DecryptError::InvalidInfo("missing encryptedKey element".into()))?;

    validate_cipher_settings(encrypted_key_node)?;

    let password_key = AgilePasswordKeyEncryptor {
        salt_value: parse_base64_attr(encrypted_key_node, "saltValue")?,
        hash_algorithm: parse_hash_algorithm(encrypted_key_node, "hashAlgorithm")?,
        spin_count: parse_u32_attr(encrypted_key_node, "spinCount")?,
        block_size: parse_usize_attr(encrypted_key_node, "blockSize")?,
        key_bits: parse_usize_attr(encrypted_key_node, "keyBits")?,
        hash_size: parse_usize_attr(encrypted_key_node, "hashSize")?,
        encrypted_verifier_hash_input: parse_base64_attr(
            encrypted_key_node,
            "encryptedVerifierHashInput",
        )?,
        encrypted_verifier_hash_value: parse_base64_attr(
            encrypted_key_node,
            "encryptedVerifierHashValue",
        )?,
        encrypted_key_value: parse_base64_attr(encrypted_key_node, "encryptedKeyValue")?,
    };

    Ok(AgileEncryptionInfo {
        key_data,
        password_key,
    })
}

fn validate_cipher_settings(node: roxmltree::Node<'_, '_>) -> Result<(), DecryptError> {
    let cipher_alg = required_attr(node, "cipherAlgorithm")?;
    if !cipher_alg.eq_ignore_ascii_case("AES") {
        return Err(DecryptError::InvalidInfo(format!(
            "unsupported cipherAlgorithm {cipher_alg}"
        )));
    }
    let chaining = required_attr(node, "cipherChaining")?;
    if !chaining.eq_ignore_ascii_case("ChainingModeCBC") {
        return Err(DecryptError::InvalidInfo(format!(
            "unsupported cipherChaining {chaining}"
        )));
    }
    Ok(())
}

fn agile_decrypt_package_key(
    password: &str,
    info: &AgileEncryptionInfo,
) -> Result<Vec<u8>, DecryptError> {
    let password_key = &info.password_key;

    let password_hash = hash_password(
        password,
        &password_key.salt_value,
        password_key.spin_count,
        password_key.hash_algorithm,
    )
    .map_err(|e| DecryptError::InvalidInfo(format!("hash_password: {e}")))?;

    let key_encrypt_key_len = key_len_bytes(password_key.key_bits, "encryptedKey", "keyBits")?;
    let package_key_len = key_len_bytes(info.key_data.key_bits, "keyData", "keyBits")?;

    // Password key encryptor uses IV = saltValue (truncated to blockSize).
    if password_key.block_size != formula_xlsx::offcrypto::AES_BLOCK_SIZE {
        return Err(DecryptError::InvalidInfo(format!(
            "unsupported encryptedKey.blockSize {} (expected {})",
            password_key.block_size,
            formula_xlsx::offcrypto::AES_BLOCK_SIZE
        )));
    }
    let verifier_iv = password_key
        .salt_value
        .get(..password_key.block_size)
        .ok_or_else(|| {
            DecryptError::InvalidInfo("encryptedKey.saltValue shorter than blockSize".into())
        })?;

    let verifier_input = {
        let k = derive_key(
            &password_hash,
            &VERIFIER_HASH_INPUT_BLOCK,
            key_encrypt_key_len,
            password_key.hash_algorithm,
        )
        .map_err(map_crypto_err("derive_key(verifierHashInput)"))?;
        let mut decrypted = password_key.encrypted_verifier_hash_input.clone();
        decrypt_aes_cbc_no_padding_in_place(&k, verifier_iv, &mut decrypted)
            .map_err(|e| DecryptError::InvalidInfo(format!("decrypt verifierHashInput: {e}")))?;
        decrypted
            .get(..password_key.block_size)
            .ok_or_else(|| {
                DecryptError::InvalidInfo(
                    "decrypted verifierHashInput shorter than blockSize".into(),
                )
            })?
            .to_vec()
    };

    let verifier_hash = {
        let k = derive_key(
            &password_hash,
            &VERIFIER_HASH_VALUE_BLOCK,
            key_encrypt_key_len,
            password_key.hash_algorithm,
        )
        .map_err(map_crypto_err("derive_key(verifierHashValue)"))?;
        let mut decrypted = password_key.encrypted_verifier_hash_value.clone();
        decrypt_aes_cbc_no_padding_in_place(&k, verifier_iv, &mut decrypted)
            .map_err(|e| DecryptError::InvalidInfo(format!("decrypt verifierHashValue: {e}")))?;
        decrypted
            .get(..password_key.hash_size)
            .ok_or_else(|| {
                DecryptError::InvalidInfo(
                    "decrypted verifierHashValue shorter than hashSize".into(),
                )
            })?
            .to_vec()
    };

    let expected_hash_full = hash_bytes(password_key.hash_algorithm, &verifier_input);
    let expected_hash = expected_hash_full
        .get(..password_key.hash_size)
        .ok_or_else(|| DecryptError::InvalidInfo("hash output shorter than hashSize".into()))?;

    if expected_hash != verifier_hash.as_slice() {
        return Err(DecryptError::InvalidPassword);
    }

    let key_value = {
        let k = derive_key(
            &password_hash,
            &KEY_VALUE_BLOCK,
            key_encrypt_key_len,
            password_key.hash_algorithm,
        )
        .map_err(map_crypto_err("derive_key(keyValue)"))?;
        let mut decrypted = password_key.encrypted_key_value.clone();
        decrypt_aes_cbc_no_padding_in_place(&k, verifier_iv, &mut decrypted)
            .map_err(|e| DecryptError::InvalidInfo(format!("decrypt encryptedKeyValue: {e}")))?;
        decrypted
            .get(..package_key_len)
            .ok_or_else(|| {
                DecryptError::InvalidInfo("decrypted keyValue shorter than keyData.keyBits".into())
            })?
            .to_vec()
    };

    Ok(key_value)
}

fn map_crypto_err(ctx: &'static str) -> impl FnOnce(CryptoError) -> DecryptError {
    move |e| DecryptError::InvalidInfo(format!("{ctx}: {e}"))
}

fn key_len_bytes(
    key_bits: usize,
    element: &'static str,
    attr: &'static str,
) -> Result<usize, DecryptError> {
    if key_bits % 8 != 0 {
        return Err(DecryptError::InvalidInfo(format!(
            "{element}.{attr} must be divisible by 8"
        )));
    }
    Ok(key_bits / 8)
}

fn hash_bytes(alg: HashAlgorithm, data: &[u8]) -> Vec<u8> {
    use sha2::Digest as _;

    match alg {
        HashAlgorithm::Sha1 => sha1::Sha1::digest(data).to_vec(),
        HashAlgorithm::Sha256 => sha2::Sha256::digest(data).to_vec(),
        HashAlgorithm::Sha384 => sha2::Sha384::digest(data).to_vec(),
        HashAlgorithm::Sha512 => sha2::Sha512::digest(data).to_vec(),
    }
}

fn required_attr<'a>(node: roxmltree::Node<'a, '_>, attr: &str) -> Result<&'a str, DecryptError> {
    node.attribute(attr).ok_or_else(|| {
        DecryptError::InvalidInfo(format!(
            "missing attribute `{attr}` on element `{}`",
            node.tag_name().name()
        ))
    })
}

fn parse_usize_attr(node: roxmltree::Node<'_, '_>, attr: &str) -> Result<usize, DecryptError> {
    let val = required_attr(node, attr)?;
    val.trim()
        .parse::<usize>()
        .map_err(|err| DecryptError::InvalidInfo(format!("invalid `{attr}` value `{val}`: {err}")))
}

fn parse_u32_attr(node: roxmltree::Node<'_, '_>, attr: &str) -> Result<u32, DecryptError> {
    let val = required_attr(node, attr)?;
    val.trim()
        .parse::<u32>()
        .map_err(|err| DecryptError::InvalidInfo(format!("invalid `{attr}` value `{val}`: {err}")))
}

fn parse_base64_attr(node: roxmltree::Node<'_, '_>, attr: &str) -> Result<Vec<u8>, DecryptError> {
    let val = required_attr(node, attr)?;
    BASE64
        .decode(val.trim())
        .map_err(|err| DecryptError::InvalidInfo(format!("base64 decode `{attr}`: {err}")))
}

fn parse_hash_algorithm(
    node: roxmltree::Node<'_, '_>,
    attr: &str,
) -> Result<HashAlgorithm, DecryptError> {
    let val = required_attr(node, attr)?;
    HashAlgorithm::parse_offcrypto_name(val)
        .map_err(|_| DecryptError::InvalidInfo(format!("unsupported hashAlgorithm `{val}`")))
}
