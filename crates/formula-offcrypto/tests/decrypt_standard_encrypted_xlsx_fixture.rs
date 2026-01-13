#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Read};
use std::path::Path;

use formula_offcrypto::{
    decrypt_standard_only, parse_encryption_info, EncryptionInfo,
};

#[test]
fn decrypts_standard_encrypted_xlsx_fixture() {
    let encrypted_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/encrypted/standard_password.xlsx"
    ));
    let plain_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/encrypted/standard_password_plain.xlsx"
    ));

    let encrypted_bytes = std::fs::read(encrypted_path).expect("read encrypted xlsx fixture");
    let plain_bytes = std::fs::read(plain_path).expect("read plaintext xlsx fixture");

    let cursor = Cursor::new(encrypted_bytes);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open OLE/CFB container");

    let mut encryption_info_bytes = Vec::new();
    ole.open_stream("EncryptionInfo")
        .expect("EncryptionInfo stream")
        .read_to_end(&mut encryption_info_bytes)
        .expect("read EncryptionInfo stream");

    let mut encrypted_package_bytes = Vec::new();
    ole.open_stream("EncryptedPackage")
        .expect("EncryptedPackage stream")
        .read_to_end(&mut encrypted_package_bytes)
        .expect("read EncryptedPackage stream");

    let info = parse_encryption_info(&encryption_info_bytes).expect("parse EncryptionInfo");
    assert!(
        matches!(info, EncryptionInfo::Standard { .. }),
        "fixture should use Standard encryption"
    );

    let decrypted =
        decrypt_standard_only(&encryption_info_bytes, &encrypted_package_bytes, "Password1234_")
            .expect("decrypt encrypted package");

    assert!(
        decrypted.starts_with(b"PK"),
        "decrypted bytes should be a ZIP/OPC package"
    );
    assert_eq!(
        decrypted, plain_bytes,
        "decrypted bytes should exactly match plaintext fixture"
    );
}
