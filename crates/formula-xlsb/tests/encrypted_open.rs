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

