use std::io::{Cursor, Read, Write};
use std::path::PathBuf;

use formula_offcrypto::decrypt_agile_ooxml_from_bytes;

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
