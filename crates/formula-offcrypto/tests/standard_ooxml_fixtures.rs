#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Read};
use std::path::PathBuf;

use formula_offcrypto::{decrypt_ooxml_standard, OffcryptoError};

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(name)
}

fn read_ole_stream(ole_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(ole_bytes);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open OLE fixture");
    let mut out = Vec::new();
    ole.open_stream(name)
        .or_else(|_| ole.open_stream(format!("/{name}")))
        .expect("open stream")
        .read_to_end(&mut out)
        .expect("read stream");
    out
}

#[test]
fn decrypts_all_standard_ooxml_fixtures() {
    let fixtures = [
        ("standard.xlsx", "plaintext.xlsx", "password"),
        ("standard-4.2.xlsx", "plaintext.xlsx", "password"),
        ("standard-rc4.xlsx", "plaintext.xlsx", "password"),
        ("standard-unicode.xlsx", "plaintext.xlsx", "pässwörd🔒"),
        ("standard-large.xlsx", "plaintext-large.xlsx", "password"),
        ("standard-basic.xlsm", "plaintext-basic.xlsm", "password"),
    ];

    for (encrypted_name, plaintext_name, password) in fixtures {
        let encrypted =
            std::fs::read(fixture_path(encrypted_name)).expect("read encrypted fixture");
        let expected = std::fs::read(fixture_path(plaintext_name)).expect("read plaintext fixture");

        let encryption_info = read_ole_stream(&encrypted, "EncryptionInfo");
        let encrypted_package = read_ole_stream(&encrypted, "EncryptedPackage");

        let decrypted = decrypt_ooxml_standard(&encryption_info, &encrypted_package, password)
            .unwrap_or_else(|err| panic!("failed to decrypt {encrypted_name}: {err:?}"));
        assert_eq!(
            decrypted, expected,
            "fixture {encrypted_name} did not decrypt to {plaintext_name}"
        );
    }
}

#[test]
fn wrong_password_returns_invalid_password_for_all_standard_ooxml_fixtures() {
    let fixtures = [
        "standard.xlsx",
        "standard-4.2.xlsx",
        "standard-rc4.xlsx",
        "standard-unicode.xlsx",
        "standard-large.xlsx",
        "standard-basic.xlsm",
    ];

    for encrypted_name in fixtures {
        let encrypted =
            std::fs::read(fixture_path(encrypted_name)).expect("read encrypted fixture");
        let encryption_info = read_ole_stream(&encrypted, "EncryptionInfo");
        let encrypted_package = read_ole_stream(&encrypted, "EncryptedPackage");

        let err = decrypt_ooxml_standard(&encryption_info, &encrypted_package, "wrong-password")
            .unwrap_err();
        assert_eq!(err, OffcryptoError::InvalidPassword);
    }
}
