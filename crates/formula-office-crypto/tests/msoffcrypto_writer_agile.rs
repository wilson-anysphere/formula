#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Read, Write};

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

#[test]
fn agile_tampered_ciphertext_fails_integrity_check() {
    let password = "correct horse battery staple";
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));

    // Encrypt a workbook using `ms-offcrypto-writer` (Excel-compatible Agile 4.4).
    let mut rng = rand09::rng();
    let cursor = Cursor::new(Vec::<u8>::new());
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).expect("create agile writer");
    writer.write_all(plaintext).expect("write plaintext");
    let cursor = writer.into_inner().expect("finalize agile writer");
    let ole_bytes = cursor.into_inner();

    // Extract streams and tamper the `EncryptedPackage` ciphertext.
    let mut ole = cfb::CompoundFile::open(Cursor::new(&ole_bytes)).expect("open cfb");
    let mut encryption_info = Vec::new();
    ole.open_stream("EncryptionInfo")
        .expect("open EncryptionInfo")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");
    let mut encrypted_package = Vec::new();
    ole.open_stream("EncryptedPackage")
        .expect("open EncryptedPackage")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    assert!(
        encrypted_package.len() > 8,
        "EncryptedPackage stream unexpectedly short"
    );
    encrypted_package[8] ^= 0x55; // flip a ciphertext byte (after size header)

    // Rebuild a minimal OLE container with the tampered stream.
    let cursor = Cursor::new(Vec::<u8>::new());
    let mut out = cfb::CompoundFile::create(cursor).expect("create cfb");
    out.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo");
    out.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage")
        .write_all(&encrypted_package)
        .expect("write EncryptedPackage");
    let tampered_ole_bytes = out.into_inner().into_inner();

    assert!(is_encrypted_ooxml_ole(&tampered_ole_bytes));

    let err =
        decrypt_encrypted_package_ole(&tampered_ole_bytes, password).expect_err("expected failure");
    assert!(
        matches!(err, OfficeCryptoError::IntegrityCheckFailed),
        "expected IntegrityCheckFailed, got {err:?}"
    );
}

#[test]
fn agile_tampered_size_header_fails_integrity_check() {
    let password = "correct horse battery staple";
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));

    // Encrypt a workbook using `ms-offcrypto-writer` (Excel-compatible Agile 4.4).
    let mut rng = rand09::rng();
    let cursor = Cursor::new(Vec::<u8>::new());
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).expect("create agile writer");
    writer.write_all(plaintext).expect("write plaintext");
    let cursor = writer.into_inner().expect("finalize agile writer");
    let ole_bytes = cursor.into_inner();

    // Extract streams and tamper the 8-byte plaintext size prefix in `EncryptedPackage`.
    let mut ole = cfb::CompoundFile::open(Cursor::new(&ole_bytes)).expect("open cfb");
    let mut encryption_info = Vec::new();
    ole.open_stream("EncryptionInfo")
        .expect("open EncryptionInfo")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");
    let mut encrypted_package = Vec::new();
    ole.open_stream("EncryptedPackage")
        .expect("open EncryptedPackage")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    assert!(
        encrypted_package.len() >= 8,
        "EncryptedPackage stream unexpectedly short"
    );
    let original_size = u64::from_le_bytes(
        encrypted_package[..8]
            .try_into()
            .expect("EncryptedPackage header is 8 bytes"),
    );
    assert!(original_size > 0, "unexpected empty EncryptedPackage payload");
    let tampered_size = original_size - 1;
    encrypted_package[..8].copy_from_slice(&tampered_size.to_le_bytes());

    // Rebuild a minimal OLE container with the tampered stream.
    let cursor = Cursor::new(Vec::<u8>::new());
    let mut out = cfb::CompoundFile::create(cursor).expect("create cfb");
    out.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo");
    out.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage")
        .write_all(&encrypted_package)
        .expect("write EncryptedPackage");
    let tampered_ole_bytes = out.into_inner().into_inner();

    assert!(is_encrypted_ooxml_ole(&tampered_ole_bytes));

    let err =
        decrypt_encrypted_package_ole(&tampered_ole_bytes, password).expect_err("expected failure");
    assert!(
        matches!(err, OfficeCryptoError::IntegrityCheckFailed),
        "expected IntegrityCheckFailed, got {err:?}"
    );
}

#[test]
fn agile_appended_ciphertext_fails_integrity_check() {
    let password = "correct horse battery staple";
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));

    // Encrypt a workbook using `ms-offcrypto-writer` (Excel-compatible Agile 4.4).
    let mut rng = rand09::rng();
    let cursor = Cursor::new(Vec::<u8>::new());
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).expect("create agile writer");
    writer.write_all(plaintext).expect("write plaintext");
    let cursor = writer.into_inner().expect("finalize agile writer");
    let ole_bytes = cursor.into_inner();

    // Extract streams and append an extra AES block to the `EncryptedPackage` ciphertext.
    let mut ole = cfb::CompoundFile::open(Cursor::new(&ole_bytes)).expect("open cfb");
    let mut encryption_info = Vec::new();
    ole.open_stream("EncryptionInfo")
        .expect("open EncryptionInfo")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");
    let mut encrypted_package = Vec::new();
    ole.open_stream("EncryptedPackage")
        .expect("open EncryptedPackage")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    encrypted_package.extend_from_slice(&[0xA5u8; 16]);

    // Rebuild a minimal OLE container with the tampered stream.
    let cursor = Cursor::new(Vec::<u8>::new());
    let mut out = cfb::CompoundFile::create(cursor).expect("create cfb");
    out.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo");
    out.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage")
        .write_all(&encrypted_package)
        .expect("write EncryptedPackage");
    let tampered_ole_bytes = out.into_inner().into_inner();

    assert!(is_encrypted_ooxml_ole(&tampered_ole_bytes));

    let err =
        decrypt_encrypted_package_ole(&tampered_ole_bytes, password).expect_err("expected failure");
    assert!(
        matches!(err, OfficeCryptoError::IntegrityCheckFailed),
        "expected IntegrityCheckFailed, got {err:?}"
    );
}
