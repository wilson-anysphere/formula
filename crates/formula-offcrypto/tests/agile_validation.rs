use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;

use formula_offcrypto::{parse_encryption_info, EncryptionInfo, OffcryptoError};

fn b64_bytes(len: usize) -> String {
    STANDARD.encode(vec![0x42u8; len])
}

fn build_agile_xml(
    key_data_salt_len: usize,
    password_salt_len: usize,
    encrypted_verifier_hash_input_len: usize,
    encrypted_verifier_hash_value_len: usize,
    encrypted_key_value_len: usize,
    encrypted_hmac_key_len: usize,
    encrypted_hmac_value_len: usize,
    key_bits: usize,
) -> String {
    let key_data_salt = b64_bytes(key_data_salt_len);
    let password_salt = b64_bytes(password_salt_len);
    let vhi = b64_bytes(encrypted_verifier_hash_input_len);
    let vhv = b64_bytes(encrypted_verifier_hash_value_len);
    let ekv = b64_bytes(encrypted_key_value_len);
    let hmac_key = b64_bytes(encrypted_hmac_key_len);
    let hmac_value = b64_bytes(encrypted_hmac_value_len);

    format!(
        r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
    xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltValue="{key_data_salt}" hashAlgorithm="SHA256" blockSize="16"/>
  <dataIntegrity encryptedHmacKey="{hmac_key}" encryptedHmacValue="{hmac_value}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="100000" saltValue="{password_salt}" hashAlgorithm="SHA512" keyBits="{key_bits}"
        encryptedKeyValue="{ekv}"
        encryptedVerifierHashInput="{vhi}"
        encryptedVerifierHashValue="{vhv}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#
    )
}

fn build_agile_encryption_info_bytes(xml: &str) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&4u16.to_le_bytes());
    bytes.extend_from_slice(&4u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes());
    bytes.extend_from_slice(xml.as_bytes());
    bytes
}

#[test]
fn agile_parse_ok_smoke() {
    let xml = build_agile_xml(16, 16, 32, 32, 32, 32, 32, 256);
    let bytes = build_agile_encryption_info_bytes(&xml);
    let info = parse_encryption_info(&bytes).expect("parse");
    assert!(matches!(info, EncryptionInfo::Agile { .. }));
}

#[test]
fn rejects_key_data_salt_too_short() {
    let xml = build_agile_xml(15, 16, 32, 32, 32, 32, 32, 256);
    let bytes = build_agile_encryption_info_bytes(&xml);
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidFormat { .. }));
}

#[test]
fn rejects_password_salt_too_short() {
    let xml = build_agile_xml(16, 15, 32, 32, 32, 32, 32, 256);
    let bytes = build_agile_encryption_info_bytes(&xml);
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidFormat { .. }));
}

#[test]
fn rejects_encrypted_verifier_hash_input_not_block_aligned() {
    let xml = build_agile_xml(16, 16, 17, 32, 32, 32, 32, 256);
    let bytes = build_agile_encryption_info_bytes(&xml);
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidFormat { .. }));
}

#[test]
fn rejects_encrypted_verifier_hash_input_too_large() {
    let xml = build_agile_xml(16, 16, 80, 32, 32, 32, 32, 256);
    let bytes = build_agile_encryption_info_bytes(&xml);
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidFormat { .. }));
}

#[test]
fn rejects_encrypted_verifier_hash_value_not_block_aligned() {
    let xml = build_agile_xml(16, 16, 32, 17, 32, 32, 32, 256);
    let bytes = build_agile_encryption_info_bytes(&xml);
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidFormat { .. }));
}

#[test]
fn rejects_encrypted_verifier_hash_value_too_large() {
    let xml = build_agile_xml(16, 16, 32, 80, 32, 32, 32, 256);
    let bytes = build_agile_encryption_info_bytes(&xml);
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidFormat { .. }));
}

#[test]
fn rejects_encrypted_key_value_not_block_aligned() {
    let xml = build_agile_xml(16, 16, 32, 32, 17, 32, 32, 256);
    let bytes = build_agile_encryption_info_bytes(&xml);
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidFormat { .. }));
}

#[test]
fn rejects_encrypted_key_value_too_large() {
    let xml = build_agile_xml(16, 16, 32, 32, 80, 32, 32, 256);
    let bytes = build_agile_encryption_info_bytes(&xml);
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidFormat { .. }));
}

#[test]
fn rejects_encrypted_key_value_too_short_for_key_bits() {
    // keyBits=256 implies at least 32 bytes of key material (ciphertext will be block aligned).
    let xml = build_agile_xml(16, 16, 32, 32, 16, 32, 32, 256);
    let bytes = build_agile_encryption_info_bytes(&xml);
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidFormat { .. }));
}

#[test]
fn rejects_encrypted_hmac_key_not_block_aligned() {
    let xml = build_agile_xml(16, 16, 32, 32, 32, 17, 32, 256);
    let bytes = build_agile_encryption_info_bytes(&xml);
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidFormat { .. }));
}

#[test]
fn rejects_encrypted_hmac_value_not_block_aligned() {
    let xml = build_agile_xml(16, 16, 32, 32, 32, 32, 17, 256);
    let bytes = build_agile_encryption_info_bytes(&xml);
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidFormat { .. }));
}
