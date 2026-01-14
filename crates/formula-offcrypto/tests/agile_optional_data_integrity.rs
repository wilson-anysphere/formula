use std::io::{Cursor, Read};
use std::path::PathBuf;

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;

use formula_offcrypto::{
    decrypt_encrypted_package, parse_encryption_info, AgileEncryptionInfo, DecryptOptions,
    EncryptionInfo, HashAlgorithm, OffcryptoError,
};

fn fixture(path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("fixtures")
        .join("encrypted")
        .join("ooxml")
        .join(path)
}

fn read_ole_stream(bytes: &[u8], name: &str) -> Vec<u8> {
    let mut ole = cfb::CompoundFile::open(Cursor::new(bytes)).expect("open fixture cfb");
    let mut stream = ole.open_stream(name).expect("open stream");
    let mut out = Vec::new();
    stream.read_to_end(&mut out).expect("read stream");
    out
}

fn build_agile_xml_without_data_integrity(info: &AgileEncryptionInfo) -> String {
    let key_data_salt_b64 = STANDARD.encode(&info.key_data_salt);
    let password_salt_b64 = STANDARD.encode(&info.password_salt);
    let encrypted_key_value_b64 = STANDARD.encode(&info.encrypted_key_value);
    let encrypted_verifier_hash_input_b64 = STANDARD.encode(&info.encrypted_verifier_hash_input);
    let encrypted_verifier_hash_value_b64 = STANDARD.encode(&info.encrypted_verifier_hash_value);

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
    xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" saltValue="{key_data_salt_b64}" hashAlgorithm="{key_hash}" blockSize="{key_block_size}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" spinCount="{spin_count}" saltValue="{password_salt_b64}" hashAlgorithm="{password_hash}" keyBits="{password_key_bits}"
        encryptedKeyValue="{encrypted_key_value_b64}"
        encryptedVerifierHashInput="{encrypted_verifier_hash_input_b64}"
        encryptedVerifierHashValue="{encrypted_verifier_hash_value_b64}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>
"#,
        key_hash = info.key_data_hash_algorithm.as_ooxml_name(),
        key_block_size = info.key_data_block_size,
        spin_count = info.spin_count,
        password_hash = info.password_hash_algorithm.as_ooxml_name(),
        password_key_bits = info.password_key_bits,
    )
}

fn build_agile_encryption_info_stream(xml_payload: &str) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&4u16.to_le_bytes());
    out.extend_from_slice(&4u16.to_le_bytes());
    out.extend_from_slice(&0u32.to_le_bytes());
    out.extend_from_slice(xml_payload.as_bytes());
    out
}

#[test]
fn parses_agile_encryption_info_without_data_integrity() {
    let key_data_salt: Vec<u8> = (0u8..16).collect();
    let password_salt: Vec<u8> = (1u8..17).collect();
    let encrypted_key_value: Vec<u8> = (0x20u8..0x40).collect(); // 32 bytes (AES block aligned)
    let encrypted_verifier_hash_input: Vec<u8> = (0x30u8..0x50).collect();
    let encrypted_verifier_hash_value: Vec<u8> = (0x40u8..0x60).collect();

    let xml = format!(
        r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
    xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltValue="{key_data_salt_b64}" hashAlgorithm="SHA256" blockSize="16"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="100000" saltValue="{password_salt_b64}" hashAlgorithm="SHA512" keyBits="256"
        encryptedKeyValue="{encrypted_key_value_b64}"
        encryptedVerifierHashInput="{encrypted_verifier_hash_input_b64}"
        encryptedVerifierHashValue="{encrypted_verifier_hash_value_b64}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        key_data_salt_b64 = STANDARD.encode(&key_data_salt),
        password_salt_b64 = STANDARD.encode(&password_salt),
        encrypted_key_value_b64 = STANDARD.encode(&encrypted_key_value),
        encrypted_verifier_hash_input_b64 = STANDARD.encode(&encrypted_verifier_hash_input),
        encrypted_verifier_hash_value_b64 = STANDARD.encode(&encrypted_verifier_hash_value),
    );

    let bytes = build_agile_encryption_info_stream(&xml);
    let info = parse_encryption_info(&bytes).expect("parse");

    let EncryptionInfo::Agile { info, .. } = info else {
        panic!("expected Agile EncryptionInfo");
    };
    assert_eq!(info.key_data_salt, key_data_salt);
    assert_eq!(info.key_data_hash_algorithm, HashAlgorithm::Sha256);
    assert!(info.data_integrity.is_none());
}

#[test]
fn decrypt_encrypted_package_without_data_integrity_succeeds_when_integrity_disabled() {
    let encrypted = std::fs::read(fixture("agile.xlsx")).expect("read encrypted fixture");
    let expected = std::fs::read(fixture("plaintext.xlsx")).expect("read expected decrypted bytes");

    let encryption_info = read_ole_stream(&encrypted, "EncryptionInfo");
    let encrypted_package = read_ole_stream(&encrypted, "EncryptedPackage");

    let info = parse_encryption_info(&encryption_info).expect("parse original EncryptionInfo");
    let EncryptionInfo::Agile { info, .. } = info else {
        panic!("expected Agile EncryptionInfo");
    };

    let xml_without_data_integrity = build_agile_xml_without_data_integrity(&info);
    let mut patched_encryption_info = encryption_info[..8].to_vec();
    patched_encryption_info.extend_from_slice(xml_without_data_integrity.as_bytes());

    let patched_parsed =
        parse_encryption_info(&patched_encryption_info).expect("parse patched EncryptionInfo");
    let EncryptionInfo::Agile { info, .. } = patched_parsed else {
        panic!("expected Agile EncryptionInfo");
    };
    assert!(
        info.data_integrity.is_none(),
        "expected patched EncryptionInfo to omit <dataIntegrity>"
    );

    let decrypted = decrypt_encrypted_package(
        &patched_encryption_info,
        &encrypted_package,
        "password",
        DecryptOptions {
            verify_integrity: false,
            ..Default::default()
        },
    )
    .expect("decrypt");

    assert!(decrypted.starts_with(b"PK"));
    assert_eq!(decrypted, expected);
}

#[test]
fn decrypt_encrypted_package_without_data_integrity_errors_when_integrity_enabled() {
    let encrypted = std::fs::read(fixture("agile.xlsx")).expect("read encrypted fixture");

    let encryption_info = read_ole_stream(&encrypted, "EncryptionInfo");
    let encrypted_package = read_ole_stream(&encrypted, "EncryptedPackage");

    let info = parse_encryption_info(&encryption_info).expect("parse original EncryptionInfo");
    let EncryptionInfo::Agile { info, .. } = info else {
        panic!("expected Agile EncryptionInfo");
    };

    let xml_without_data_integrity = build_agile_xml_without_data_integrity(&info);
    let mut patched_encryption_info = encryption_info[..8].to_vec();
    patched_encryption_info.extend_from_slice(xml_without_data_integrity.as_bytes());

    let err = decrypt_encrypted_package(
        &patched_encryption_info,
        &encrypted_package,
        "password",
        DecryptOptions {
            verify_integrity: true,
            ..Default::default()
        },
    )
    .expect_err("expected integrity verification to require <dataIntegrity>");

    assert_eq!(
        err,
        OffcryptoError::InvalidEncryptionInfo {
            context: "missing <dataIntegrity> element"
        }
    );
}

