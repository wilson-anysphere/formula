use formula_office_crypto::{
    decrypt_encrypted_package_ole, encrypt_package_to_ole, EncryptOptions, EncryptionScheme,
    HashAlgorithm, OfficeCryptoError,
};
use std::io::{Cursor, Read, Write};

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

    let ole = encrypt_package_to_ole(
        &zip,
        password,
        EncryptOptions {
            spin_count: 10_000,
            ..Default::default()
        },
    )
    .expect("encrypt");
    let decrypted = decrypt_encrypted_package_ole(&ole, password).expect("decrypt");
    assert_eq!(decrypted, zip);
}

#[test]
fn wrong_password_fails() {
    let zip = basic_xlsx_fixture_bytes();
    let ole = encrypt_package_to_ole(
        &zip,
        "password",
        EncryptOptions {
            spin_count: 10_000,
            ..Default::default()
        },
    )
    .expect("encrypt");

    let err =
        decrypt_encrypted_package_ole(&ole, "not-the-password").expect_err("expected failure");
    assert!(
        matches!(err, OfficeCryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}

#[test]
fn standard_encrypt_decrypt_round_trip() {
    let zip = basic_xlsx_fixture_bytes();
    let password = "swordfish";

    let ole = encrypt_package_to_ole(
        &zip,
        password,
        EncryptOptions {
            scheme: EncryptionScheme::Standard,
            key_bits: 128,
            hash_algorithm: HashAlgorithm::Sha1,
            // Standard uses a fixed 50k spin count internally (CryptoAPI), but keep the option
            // explicit so callers don't accidentally rely on Agile defaults.
            spin_count: 50_000,
        },
    )
    .expect("encrypt");
    let decrypted = decrypt_encrypted_package_ole(&ole, password).expect("decrypt");
    assert_eq!(decrypted, zip);
}

#[test]
fn tampered_ciphertext_fails_integrity_check() {
    let zip = basic_xlsx_fixture_bytes();
    let password = "correct horse battery staple";

    let ole = encrypt_package_to_ole(
        &zip,
        password,
        EncryptOptions {
            spin_count: 10_000,
            ..Default::default()
        },
    )
    .expect("encrypt");

    // Extract the streams, flip a byte in the EncryptedPackage ciphertext, and re-wrap into a new
    // OLE container.
    let cursor = Cursor::new(&ole);
    let mut ole_in = cfb::CompoundFile::open(cursor).expect("open cfb");

    let mut encryption_info = Vec::new();
    ole_in
        .open_stream("EncryptionInfo")
        .expect("open EncryptionInfo")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let mut encrypted_package = Vec::new();
    ole_in
        .open_stream("EncryptedPackage")
        .expect("open EncryptedPackage")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    assert!(
        encrypted_package.len() > 8,
        "EncryptedPackage should contain a size prefix and ciphertext"
    );
    encrypted_package[8] ^= 0x01; // Flip a byte in the ciphertext (not the length prefix).

    let cursor_out = Cursor::new(Vec::new());
    let mut ole_out = cfb::CompoundFile::create(cursor_out).expect("create cfb");
    ole_out
        .create_stream("EncryptionInfo")
        .expect("create EncryptionInfo")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo");
    ole_out
        .create_stream("EncryptedPackage")
        .expect("create EncryptedPackage")
        .write_all(&encrypted_package)
        .expect("write EncryptedPackage");

    let tampered = ole_out.into_inner().into_inner();

    let err = decrypt_encrypted_package_ole(&tampered, password).expect_err("expected failure");
    assert!(
        matches!(err, OfficeCryptoError::IntegrityCheckFailed),
        "expected IntegrityCheckFailed, got {err:?}"
    );
}
