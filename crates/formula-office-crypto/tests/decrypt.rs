use std::io::Cursor;

use formula_office_crypto::{decrypt_encrypted_package, OfficeCryptoError};

const AGILE_FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/agile-large.xlsx"
));
const AGILE_PLAINTEXT: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/plaintext-large.xlsx"
));
const STANDARD_FIXTURE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/standard.xlsx"
));
const STANDARD_PLAINTEXT: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../fixtures/encrypted/ooxml/plaintext.xlsx"
));

fn assert_decrypted_zip_contains_workbook(decrypted: &[u8]) {
    assert!(
        decrypted.starts_with(b"PK"),
        "decrypted payload should start with ZIP magic"
    );

    let archive = zip::ZipArchive::new(Cursor::new(decrypted)).expect("open zip archive");
    let mut has_workbook = false;
    for name in archive.file_names() {
        if name.eq_ignore_ascii_case("xl/workbook.xml")
            || name.eq_ignore_ascii_case("xl/workbook.bin")
        {
            has_workbook = true;
            break;
        }
    }
    assert!(has_workbook, "expected decrypted ZIP to contain xl/workbook.*");
}

#[test]
fn decrypts_agile_encrypted_package() {
    let decrypted = decrypt_encrypted_package(AGILE_FIXTURE, "password").expect("decrypt agile");
    assert_eq!(decrypted.as_slice(), AGILE_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn decrypts_standard_encrypted_package() {
    let decrypted =
        decrypt_encrypted_package(STANDARD_FIXTURE, "password").expect("decrypt standard");
    assert_eq!(decrypted.as_slice(), STANDARD_PLAINTEXT);
    assert_decrypted_zip_contains_workbook(&decrypted);
}

#[test]
fn wrong_password_returns_invalid_password() {
    let err = decrypt_encrypted_package(AGILE_FIXTURE, "wrong").expect_err("expected error");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );

    let err = decrypt_encrypted_package(STANDARD_FIXTURE, "wrong").expect_err("expected error");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}
