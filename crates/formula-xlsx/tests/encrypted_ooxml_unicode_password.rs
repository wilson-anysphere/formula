#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Read, Seek};
use std::path::{Path, PathBuf};

use formula_xlsx::offcrypto::{decrypt_ooxml_encrypted_package, OffCryptoError};

fn fixture_path(rel: &str) -> PathBuf {
    Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/../../fixtures/encrypted/ooxml/")).join(rel)
}

fn read_stream<R: Read + Seek + std::io::Write>(ole: &mut cfb::CompoundFile<R>, name: &str) -> Vec<u8> {
    let mut stream = ole
        .open_stream(name)
        .or_else(|_| ole.open_stream(&format!("/{name}")))
        .unwrap_or_else(|err| panic!("open {name} stream: {err}"));

    let mut buf = Vec::new();
    stream
        .read_to_end(&mut buf)
        .unwrap_or_else(|err| panic!("read {name} stream: {err}"));
    buf
}

#[test]
fn decrypts_agile_unicode_password_fixture() {
    let encrypted =
        std::fs::read(fixture_path("agile-unicode.xlsx")).expect("read agile-unicode.xlsx fixture");
    let expected =
        std::fs::read(fixture_path("plaintext.xlsx")).expect("read plaintext.xlsx fixture");

    let mut ole = cfb::CompoundFile::open(Cursor::new(encrypted)).expect("open OLE container");
    let encryption_info = read_stream(&mut ole, "EncryptionInfo");
    let encrypted_package = read_stream(&mut ole, "EncryptedPackage");

    let decrypted =
        decrypt_ooxml_encrypted_package(&encryption_info, &encrypted_package, "pässwörd")
            .expect("decrypt agile-unicode fixture");
    assert_eq!(decrypted, expected);

    let err = decrypt_ooxml_encrypted_package(&encryption_info, &encrypted_package, "wrong-password")
        .expect_err("expected wrong password to fail");
    assert!(
        matches!(err, OffCryptoError::WrongPassword | OffCryptoError::IntegrityMismatch),
        "unexpected error: {err:?}"
    );
}

