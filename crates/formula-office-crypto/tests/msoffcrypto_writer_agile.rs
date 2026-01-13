#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use ms_offcrypto_writer::Ecma376AgileWriter;

use formula_office_crypto::{decrypt_encrypted_package_ole, is_encrypted_ooxml_ole, OfficeCryptoError};

#[test]
fn agile_decrypt_msoffcrypto_writer_roundtrip() {
    let password = "correct horse battery staple";
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));

    // Encrypt the ZIP bytes into an OLE `EncryptionInfo` + `EncryptedPackage` wrapper, matching
    // Excel/MS-OFFCRYPTO behavior.
    let mut rng = rand09::rng();
    let cursor = Cursor::new(Vec::<u8>::new());
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).expect("create agile writer");
    writer.write_all(plaintext).expect("write plaintext");
    let cursor = writer.into_inner().expect("finalize agile writer");
    let ole_bytes = cursor.into_inner();

    assert!(
        is_encrypted_ooxml_ole(&ole_bytes),
        "expected ms-offcrypto-writer to emit an OLE EncryptedPackage wrapper"
    );

    let decrypted = decrypt_encrypted_package_ole(&ole_bytes, password).expect("decrypt");
    assert_eq!(decrypted, plaintext);

    let err = decrypt_encrypted_package_ole(&ole_bytes, "wrong-password").expect_err("wrong pw");
    assert!(matches!(err, OfficeCryptoError::InvalidPassword));
}
