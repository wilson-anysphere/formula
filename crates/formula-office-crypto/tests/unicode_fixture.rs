use formula_office_crypto::{decrypt_encrypted_package_ole, OfficeCryptoError};

const AGILE_UNICODE_FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/agile-unicode.xlsx"
));

#[test]
fn decrypts_agile_unicode_fixture() {
    let decrypted =
        decrypt_encrypted_package_ole(AGILE_UNICODE_FIXTURE, "pässwörd").expect("decrypt");
    assert!(
        decrypted.starts_with(b"PK"),
        "decrypted payload should start with ZIP magic"
    );
}

#[test]
fn agile_unicode_wrong_password_returns_invalid_password() {
    let err =
        decrypt_encrypted_package_ole(AGILE_UNICODE_FIXTURE, "wrong-password").expect_err("wrong");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}

