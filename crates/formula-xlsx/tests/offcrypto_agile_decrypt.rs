#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Read, Write};

use cfb::CompoundFile;
use ms_offcrypto_writer::Ecma376AgileWriter;
use zip::write::FileOptions;

use formula_xlsx::{decrypt_agile_encrypted_package, OffCryptoError};
use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine as _;

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
    let mut agile =
        Ecma376AgileWriter::create(&mut rand::rng(), password, &mut cursor).expect("create agile");
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

#[test]
fn agile_decrypt_roundtrip() {
    let password = "correct horse battery staple";
    let plain_zip = build_tiny_zip();

    let encrypted_cfb = encrypt_zip_with_password(&plain_zip, password);
    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    let decrypted =
        decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password).unwrap();
    assert_eq!(decrypted, plain_zip);
}

#[test]
fn agile_decrypt_wrong_password_fails() {
    let plain_zip = build_tiny_zip();
    let encrypted_cfb = encrypt_zip_with_password(&plain_zip, "password-1");

    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    let err = decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, "password-2")
        .expect_err("wrong password should fail");
    match err {
        OffCryptoError::WrongPassword | OffCryptoError::IntegrityMismatch => {}
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn agile_decrypt_errors_on_non_block_aligned_verifier_ciphertext() {
    // Build a minimal Agile `EncryptionInfo` stream where one ciphertext blob decodes to a
    // non-multiple-of-16 length. The decrypt implementation should fail early with a structured
    // error pointing at the offending field.
    let invalid = BASE64.encode([0u8; 15]); // 15 % 16 != 0
    let valid = BASE64.encode([0u8; 16]);

    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltValue="{valid}" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
           blockSize="16" keyBits="128" hashSize="20" />
  <dataIntegrity encryptedHmacKey="{valid}" encryptedHmacValue="{valid}" />
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey saltValue="{valid}" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                      spinCount="0" blockSize="16" keyBits="128" hashSize="20"
                      encryptedVerifierHashInput="{invalid}"
                      encryptedVerifierHashValue="{valid}"
                      encryptedKeyValue="{valid}" />
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#
    );

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&0u32.to_le_bytes());
    encryption_info.extend_from_slice(xml.as_bytes());

    let err = decrypt_agile_encrypted_package(&encryption_info, &[], "pw")
        .expect_err("expected CiphertextNotBlockAligned");
    assert!(
        matches!(
            err,
            OffCryptoError::CiphertextNotBlockAligned {
                field: "encryptedVerifierHashInput",
                len: 15
            }
        ),
        "unexpected error: {err:?}"
    );
}
