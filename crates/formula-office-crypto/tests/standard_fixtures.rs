use std::path::Path;

use formula_office_crypto::{decrypt_encrypted_package_ole, OfficeCryptoError};

#[test]
fn decrypts_real_standard_docx_fixture() {
    let encrypted_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-offcrypto/tests/fixtures/inputs/ecma376standard_password.docx");
    let expected_plain_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-offcrypto/tests/fixtures/outputs/ecma376standard_password_plain.docx");

    let encrypted = std::fs::read(&encrypted_path).expect("read encrypted fixture");
    let expected_plain = std::fs::read(&expected_plain_path).expect("read plaintext fixture");

    let decrypted =
        decrypt_encrypted_package_ole(&encrypted, "Password1234_").expect("decrypt fixture");
    assert_eq!(decrypted, expected_plain);

    let err =
        decrypt_encrypted_package_ole(&encrypted, "wrong-password").expect_err("wrong password");
    assert!(matches!(err, OfficeCryptoError::InvalidPassword));
}

#[test]
fn decrypts_standard_xlsx_fixture_from_fixtures_dir() {
    let encrypted_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml/standard.xlsx");
    let expected_plain_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml/plaintext.xlsx");

    let encrypted = std::fs::read(&encrypted_path).expect("read encrypted fixture");
    let expected_plain = std::fs::read(&expected_plain_path).expect("read plaintext fixture");

    let decrypted = decrypt_encrypted_package_ole(&encrypted, "password").expect("decrypt fixture");
    assert_eq!(decrypted, expected_plain);

    let err =
        decrypt_encrypted_package_ole(&encrypted, "wrong-password").expect_err("wrong password");
    assert!(matches!(err, OfficeCryptoError::InvalidPassword));
}

