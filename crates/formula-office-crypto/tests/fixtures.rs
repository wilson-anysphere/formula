use std::path::PathBuf;

use formula_office_crypto::{decrypt_encrypted_package_ole, OfficeCryptoError};

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(name)
}

fn read_fixture(name: &str) -> Vec<u8> {
    std::fs::read(fixture_path(name)).unwrap_or_else(|err| panic!("read fixture {name}: {err}"))
}

#[test]
fn decrypts_standard_fixture_matches_plaintext() {
    let plaintext = read_fixture("plaintext.xlsx");
    let standard = read_fixture("standard.xlsx");

    let decrypted = decrypt_encrypted_package_ole(&standard, "password").expect("decrypt standard");
    assert_eq!(decrypted, plaintext);
    assert!(decrypted.starts_with(b"PK"));
}

#[test]
fn decrypts_agile_fixture_matches_plaintext() {
    let plaintext = read_fixture("plaintext.xlsx");
    let agile = read_fixture("agile.xlsx");

    let decrypted = decrypt_encrypted_package_ole(&agile, "password").expect("decrypt agile");
    assert_eq!(decrypted, plaintext);
    assert!(decrypted.starts_with(b"PK"));
}

#[test]
fn standard_wrong_password_returns_invalid_password() {
    let standard = read_fixture("standard.xlsx");

    let err =
        decrypt_encrypted_package_ole(&standard, "wrong").expect_err("wrong password should fail");
    assert!(matches!(err, OfficeCryptoError::InvalidPassword));
}
