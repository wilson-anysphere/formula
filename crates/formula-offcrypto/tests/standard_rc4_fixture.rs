use std::path::PathBuf;
use std::io::{Cursor, Read};

use formula_offcrypto::{decrypt_encrypted_package, decrypt_standard_ooxml_from_bytes, DecryptOptions, OffcryptoError};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures")
        .join(rel)
}

fn extract_stream_bytes(cfb_bytes: &[u8], stream_name: &str) -> Vec<u8> {
    let mut ole = cfb::CompoundFile::open(Cursor::new(cfb_bytes)).expect("open cfb");
    let mut stream = ole.open_stream(stream_name).expect("open stream");
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf).expect("read stream");
    buf
}

#[test]
fn decrypts_standard_rc4_fixture_to_plaintext_zip() {
    let encrypted_path = fixture_path("encrypted/ooxml/standard-rc4.xlsx");
    let plaintext_path = fixture_path("encrypted/ooxml/plaintext.xlsx");

    let encrypted = std::fs::read(&encrypted_path).expect("read encrypted fixture");
    let decrypted =
        decrypt_standard_ooxml_from_bytes(encrypted, "password").expect("decrypt standard-rc4.xlsx");

    let plaintext = std::fs::read(&plaintext_path).expect("read plaintext fixture");
    assert_eq!(decrypted, plaintext);
}

#[test]
fn decrypt_encrypted_package_decrypts_standard_rc4_fixture_streams() {
    let encrypted_path = fixture_path("encrypted/ooxml/standard-rc4.xlsx");
    let plaintext_path = fixture_path("encrypted/ooxml/plaintext.xlsx");

    let encrypted = std::fs::read(&encrypted_path).expect("read encrypted fixture");
    let encryption_info = extract_stream_bytes(&encrypted, "EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted, "EncryptedPackage");

    let decrypted = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "password",
        DecryptOptions::default(),
    )
    .expect("decrypt standard-rc4 EncryptedPackage streams");

    let plaintext = std::fs::read(&plaintext_path).expect("read plaintext fixture");
    assert_eq!(decrypted, plaintext);
}

#[test]
fn decrypt_encrypted_package_standard_rc4_rejects_size_mismatch_before_password_check() {
    let encrypted_path = fixture_path("encrypted/ooxml/standard-rc4.xlsx");
    let encrypted = std::fs::read(&encrypted_path).expect("read encrypted fixture");
    let encryption_info = extract_stream_bytes(&encrypted, "EncryptionInfo");

    // total_size=32, but ciphertext is only 16 bytes. This should error structurally even if the
    // password is wrong.
    let mut encrypted_package = 32u64.to_le_bytes().to_vec();
    encrypted_package.extend_from_slice(&[0u8; 16]);

    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "wrong-password",
        DecryptOptions::default(),
    )
    .unwrap_err();

    assert_eq!(
        err,
        OffcryptoError::EncryptedPackageSizeMismatch {
            total_size: 32,
            ciphertext_len: 16
        }
    );
}
