use formula_office_crypto::{encrypt_package_to_ole, EncryptOptions, EncryptionScheme, HashAlgorithm};
use formula_xlsb::{OpenOptions, XlsbWorkbook};

#[test]
fn opens_encrypted_xlsb_from_bytes() {
    let plain_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let plain_zip = std::fs::read(plain_path).expect("read plain xlsb fixture");

    let password = "password";
    let encrypted = encrypt_package_to_ole(
        &plain_zip,
        password,
        EncryptOptions {
            scheme: EncryptionScheme::Agile,
            key_bits: 256,
            hash_algorithm: HashAlgorithm::Sha256,
            // Keep the test fast; real files typically use much larger values (e.g. 100,000).
            spin_count: 2_048,
        },
    )
    .expect("encrypt wrapper");

    let wb =
        XlsbWorkbook::open_from_bytes_with_password(&encrypted, password, OpenOptions::default())
            .expect("open encrypted xlsb from bytes");
    assert_eq!(wb.sheet_metas().len(), 1);
}

#[test]
fn opens_encrypted_xlsb_from_path() {
    let plain_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let plain_zip = std::fs::read(plain_path).expect("read plain xlsb fixture");

    let password = "password";
    let encrypted = encrypt_package_to_ole(
        &plain_zip,
        password,
        EncryptOptions {
            scheme: EncryptionScheme::Agile,
            key_bits: 256,
            hash_algorithm: HashAlgorithm::Sha256,
            spin_count: 2_048,
        },
    )
    .expect("encrypt wrapper");

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&path, encrypted).expect("write encrypted wrapper");

    let wb = XlsbWorkbook::open_with_password(&path, password).expect("open encrypted xlsb");
    assert_eq!(wb.sheet_metas().len(), 1);
}

#[test]
fn wrong_password_is_invalid_password() {
    let plain_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let plain_zip = std::fs::read(plain_path).expect("read plain xlsb fixture");

    let encrypted = encrypt_package_to_ole(
        &plain_zip,
        "correct",
        EncryptOptions {
            scheme: EncryptionScheme::Agile,
            key_bits: 256,
            hash_algorithm: HashAlgorithm::Sha256,
            spin_count: 2_048,
        },
    )
    .expect("encrypt wrapper");

    let err = XlsbWorkbook::open_from_bytes_with_password(&encrypted, "wrong", OpenOptions::default())
        .expect_err("expected wrong password to error");
    assert!(
        matches!(err, formula_xlsb::Error::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}

#[test]
fn unicode_and_whitespace_passwords_are_not_trimmed_or_normalized() {
    let plain_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let plain_zip = std::fs::read(plain_path).expect("read plain xlsb fixture");

    // Trailing whitespace is significant and the emoji exercises non-BMP UTF-16 surrogate pairs.
    let password = "pässwörd🔒 ";
    let encrypted = encrypt_package_to_ole(
        &plain_zip,
        password,
        EncryptOptions {
            scheme: EncryptionScheme::Agile,
            key_bits: 256,
            hash_algorithm: HashAlgorithm::Sha256,
            spin_count: 2_048,
        },
    )
    .expect("encrypt wrapper");

    let wb = XlsbWorkbook::open_from_bytes_with_password(&encrypted, password, OpenOptions::default())
        .expect("open encrypted xlsb with exact password");
    assert_eq!(wb.sheet_metas().len(), 1);

    let trimmed = password.trim();
    let err = XlsbWorkbook::open_from_bytes_with_password(&encrypted, trimmed, OpenOptions::default())
        .expect_err("expected trimmed password to fail");
    assert!(
        matches!(err, formula_xlsb::Error::InvalidPassword),
        "expected InvalidPassword for trimmed password, got {err:?}"
    );

    // NFC vs NFD should differ: visually-similar strings must not be normalized.
    let nfd = "pa\u{0308}sswo\u{0308}rd🔒 ";
    let err = XlsbWorkbook::open_from_bytes_with_password(&encrypted, nfd, OpenOptions::default())
        .expect_err("expected NFD password to fail");
    assert!(
        matches!(err, formula_xlsb::Error::InvalidPassword),
        "expected InvalidPassword for NFD password, got {err:?}"
    );
}
