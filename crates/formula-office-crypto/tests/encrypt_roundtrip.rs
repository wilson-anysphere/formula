use formula_office_crypto::{
    decrypt_encrypted_package_ole, encrypt_package_to_ole, EncryptOptions, OfficeCryptoError,
};

fn basic_xlsx_fixture_bytes() -> Vec<u8> {
    let path = std::path::Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));
    std::fs::read(path).expect("read basic.xlsx fixture")
}

#[test]
fn agile_encrypt_decrypt_round_trip() {
    let zip = basic_xlsx_fixture_bytes();
    let password = "correct horse battery staple";

    let ole = encrypt_package_to_ole(&zip, password, EncryptOptions::default()).expect("encrypt");
    let decrypted = decrypt_encrypted_package_ole(&ole, password).expect("decrypt");
    assert_eq!(decrypted, zip);
}

#[test]
fn wrong_password_fails() {
    let zip = basic_xlsx_fixture_bytes();
    let ole = encrypt_package_to_ole(&zip, "password", EncryptOptions::default()).expect("encrypt");

    let err =
        decrypt_encrypted_package_ole(&ole, "not-the-password").expect_err("expected failure");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}
