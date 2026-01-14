use std::io::{Cursor, Read, Write};
use std::path::PathBuf;

use formula_offcrypto::{decrypt_agile_ooxml_from_bytes, EncryptionType, OffcryptoError};

fn fixture(path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("fixtures")
        .join("encrypted")
        .join("ooxml")
        .join(path)
}

#[test]
fn decrypts_agile_fixture_xlsx() {
    let encrypted = std::fs::read(fixture("agile.xlsx")).expect("read encrypted fixture");
    let expected = std::fs::read(fixture("plaintext.xlsx")).expect("read expected decrypted bytes");

    let decrypted = decrypt_agile_ooxml_from_bytes(encrypted, "password").expect("decrypt fixture");
    assert!(decrypted.starts_with(b"PK"));
    assert_eq!(decrypted, expected);
}

#[test]
fn decrypts_agile_fixture_xlsm() {
    let encrypted = std::fs::read(fixture("agile-basic.xlsm")).expect("read encrypted fixture");
    let expected =
        std::fs::read(fixture("plaintext-basic.xlsm")).expect("read expected decrypted bytes");

    let decrypted = decrypt_agile_ooxml_from_bytes(encrypted, "password").expect("decrypt fixture");
    assert!(decrypted.starts_with(b"PK"));
    assert_eq!(decrypted, expected);
}

#[test]
fn decrypts_agile_fixture_empty_password_xlsx() {
    let encrypted =
        std::fs::read(fixture("agile-empty-password.xlsx")).expect("read encrypted fixture");
    let expected = std::fs::read(fixture("plaintext.xlsx")).expect("read expected decrypted bytes");

    let decrypted = decrypt_agile_ooxml_from_bytes(encrypted, "").expect("decrypt fixture");
    assert!(decrypted.starts_with(b"PK"));
    assert_eq!(decrypted, expected);
}

#[test]
fn agile_wrong_password_returns_invalid_password() {
    let encrypted = std::fs::read(fixture("agile.xlsx")).expect("read encrypted fixture");

    let err = decrypt_agile_ooxml_from_bytes(encrypted, "not-the-password")
        .expect_err("expected wrong password to error");
    assert!(
        matches!(&err, formula_offcrypto::OffcryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}

#[test]
fn supports_case_insensitive_stream_names() {
    let encrypted = std::fs::read(fixture("agile.xlsx")).expect("read encrypted fixture");
    let expected = std::fs::read(fixture("plaintext.xlsx")).expect("read expected decrypted bytes");

    let mut ole_fixture = cfb::CompoundFile::open(Cursor::new(encrypted)).expect("open fixture cfb");
    let mut encryption_info = Vec::new();
    ole_fixture
        .open_stream("EncryptionInfo")
        .expect("open EncryptionInfo stream")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo stream");
    let mut encrypted_package = Vec::new();
    ole_fixture
        .open_stream("EncryptedPackage")
        .expect("open EncryptedPackage stream")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage stream");

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("encryptioninfo")
        .expect("create encryptioninfo stream")
        .write_all(&encryption_info)
        .expect("write encryptioninfo stream");
    ole.create_stream("encryptedpackage")
        .expect("create encryptedpackage stream")
        .write_all(&encrypted_package)
        .expect("write encryptedpackage stream");

    let decrypted =
        decrypt_agile_ooxml_from_bytes(ole.into_inner().into_inner(), "password").expect("decrypt");
    assert_eq!(decrypted, expected);
}

#[test]
fn rejects_standard_fixture() {
    let encrypted = std::fs::read(fixture("standard.xlsx")).expect("read encrypted fixture");

    let err = decrypt_agile_ooxml_from_bytes(encrypted, "password")
        .expect_err("expected Agile decryptor to reject Standard encryption");
    assert!(
        matches!(
            &err,
            OffcryptoError::UnsupportedEncryption {
                encryption_type: EncryptionType::Standard,
                ..
            }
        ),
        "expected UnsupportedEncryption(Standard), got {err:?}"
    );
}

#[test]
fn missing_encryptioninfo_stream_returns_error() {
    let cursor = Cursor::new(Vec::new());
    let ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    let bytes = ole.into_inner().into_inner();

    let err = decrypt_agile_ooxml_from_bytes(bytes, "pw").unwrap_err();
    assert!(
        matches!(
            &err,
            OffcryptoError::InvalidStructure(msg) if msg.contains("missing `EncryptionInfo` stream")
        ),
        "expected InvalidStructure missing EncryptionInfo, got {err:?}"
    );
}

#[test]
fn invalid_ole_container_returns_error() {
    let err = decrypt_agile_ooxml_from_bytes(vec![0u8; 32], "pw").unwrap_err();
    assert!(
        matches!(
            &err,
            OffcryptoError::InvalidStructure(msg) if msg.contains("failed to open OLE compound file")
        ),
        "expected InvalidStructure for invalid OLE container, got {err:?}"
    );
}

#[test]
fn missing_encryptedpackage_stream_returns_error_without_verifying_password() {
    // This should fail with a structural error before performing the expensive Agile password KDF.
    let encrypted = std::fs::read(fixture("agile.xlsx")).expect("read encrypted fixture");
    let mut ole_fixture = cfb::CompoundFile::open(Cursor::new(encrypted)).expect("open fixture cfb");
    let mut encryption_info = Vec::new();
    ole_fixture
        .open_stream("EncryptionInfo")
        .expect("open EncryptionInfo stream")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo stream");

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo stream")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo stream");

    // Use an incorrect password; we should still get a missing EncryptedPackage error.
    let err = decrypt_agile_ooxml_from_bytes(ole.into_inner().into_inner(), "not-the-password")
        .unwrap_err();
    assert!(
        matches!(
            &err,
            OffcryptoError::InvalidStructure(msg) if msg.contains("missing `EncryptedPackage` stream")
        ),
        "expected InvalidStructure missing EncryptedPackage, got {err:?}"
    );
}
