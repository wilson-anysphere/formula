use std::path::Path;
use std::io::Cursor;

use formula_offcrypto::{parse_encryption_info, EncryptionInfo};
use formula_io::offcrypto::{
    decrypt_encrypted_package_standard_aes_to_writer, decrypt_standard_encrypted_package_stream,
};

const KEY: [u8; 16] = [
    0x0b, 0x9d, 0x2f, 0xab, 0xa2, 0xd8, 0xe6, 0xe7, 0xac, 0x2e, 0xc2, 0xc5, 0xa1, 0xfc,
    0xc4, 0xa1,
];

const SALT: [u8; 16] = [
    0x91, 0x33, 0xca, 0x74, 0x07, 0xdd, 0x5a, 0x2d, 0x04, 0x55, 0x34, 0x91, 0x79, 0xe3,
    0x2a, 0xe9,
];

fn fixture_path(name: &str) -> std::path::PathBuf {
    Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/offcrypto/"
    ))
    .join(name)
}

#[test]
fn decrypts_real_excel_standard_aes_encryptedpackage_fixture() {
    let encryption_info =
        std::fs::read(fixture_path("standard_aes_encryptioninfo.bin")).expect("read EncryptionInfo");
    let parsed =
        parse_encryption_info(&encryption_info).expect("parse EncryptionInfo (Standard CryptoAPI)");
    let verifier_salt = match parsed {
        EncryptionInfo::Standard { verifier, .. } => verifier.salt,
        other => panic!("expected Standard EncryptionInfo, got {other:?}"),
    };
    assert_eq!(
        verifier_salt.as_slice(),
        SALT.as_ref(),
        "fixture salt constant should match EncryptionVerifier.salt"
    );

    let encrypted_package =
        std::fs::read(fixture_path("standard_aes_encryptedpackage.bin")).expect("read fixture");
    let expected =
        std::fs::read(fixture_path("standard_aes_plain.zip")).expect("read expected plaintext zip");

    // Standard/CryptoAPI AES `EncryptedPackage` uses AES-ECB (no IV). The salt is only used for
    // key derivation / password verification, not for package decryption.
    let decrypted =
        decrypt_standard_encrypted_package_stream(&encrypted_package, &KEY, &SALT).expect("decrypt");

    assert_eq!(decrypted.len(), expected.len());
    assert_eq!(decrypted, expected);
    assert!(decrypted.starts_with(b"PK\x03\x04"), "expected ZIP magic");

    // Also validate the streaming decryptor (which exercises 0x1000 segment sizing + orig_size
    // truncation rules without allocating a ciphertext-sized buffer).
    let mut streamed = Vec::new();
    let written = decrypt_encrypted_package_standard_aes_to_writer(
        Cursor::new(encrypted_package.as_slice()),
        &KEY,
        &SALT,
        &mut streamed,
    )
    .expect("stream decrypt");
    assert_eq!(written as usize, expected.len());
    assert_eq!(streamed, expected);

    // Regression: some producers treat the 8-byte size prefix as (u32 size, u32 reserved).
    let mut mutated = encrypted_package.clone();
    mutated[4..8].copy_from_slice(&1u32.to_le_bytes());

    let decrypted = decrypt_standard_encrypted_package_stream(&mutated, &KEY, &SALT).expect("decrypt");
    assert_eq!(decrypted, expected);

    let mut streamed = Vec::new();
    let written = decrypt_encrypted_package_standard_aes_to_writer(
        Cursor::new(mutated.as_slice()),
        &KEY,
        &SALT,
        &mut streamed,
    )
    .expect("stream decrypt");
    assert_eq!(written as usize, expected.len());
    assert_eq!(streamed, expected);
}
