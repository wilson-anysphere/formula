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
        matches!(err, formula_offcrypto::OffcryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}

