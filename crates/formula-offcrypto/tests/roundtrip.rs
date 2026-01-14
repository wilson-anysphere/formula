use std::io::{Cursor, Read, Write};

use formula_offcrypto::{
    agile_decrypt_package, agile_secret_key, agile_verify_password, decrypt_standard_ooxml_from_bytes,
    parse_encryption_info, EncryptionInfo,
};

mod support;

fn build_test_zip() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options)
        .expect("start [Content_Types].xml");
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#)
        .expect("write [Content_Types].xml");

    zip.start_file("xl/workbook.xml", options)
        .expect("start xl/workbook.xml");
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#)
        .expect("write xl/workbook.xml");

    zip.finish().expect("finish zip").into_inner()
}

fn assert_zip_contains_workbook_xml(bytes: &[u8]) {
    let cursor = Cursor::new(bytes);
    let zip = zip::ZipArchive::new(cursor).expect("zip archive");
    let mut found = false;
    for name in zip.file_names() {
        if name.eq_ignore_ascii_case("xl/workbook.xml") {
            found = true;
            break;
        }
    }
    assert!(found, "zip should contain xl/workbook.xml");
}

fn read_ole_stream(ole_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(ole_bytes);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open ole");
    let mut out = Vec::new();
    ole.open_stream(name)
        .expect("open stream")
        .read_to_end(&mut out)
        .expect("read stream");
    out
}

#[test]
fn roundtrip_standard_encryption() {
    let password = "Password";
    let plaintext = build_test_zip();

    let (encryption_info, encrypted_package) = support::encrypt_standard(&plaintext, password);
    let ole_bytes = support::wrap_in_ole_cfb(&encryption_info, &encrypted_package);

    let decrypted = decrypt_standard_ooxml_from_bytes(ole_bytes, password).expect("decrypt");

    assert_eq!(decrypted, plaintext);
    assert_zip_contains_workbook_xml(&decrypted);
}

#[test]
fn roundtrip_agile_encryption() {
    let password = "Password";
    let plaintext = build_test_zip();

    let (encryption_info, encrypted_package) = support::encrypt_agile(&plaintext, password);
    let ole_bytes = support::wrap_in_ole_cfb(&encryption_info, &encrypted_package);

    let encryption_info = read_ole_stream(&ole_bytes, "EncryptionInfo");
    let encrypted_package = read_ole_stream(&ole_bytes, "EncryptedPackage");

    let parsed = parse_encryption_info(&encryption_info).expect("parse EncryptionInfo");
    let EncryptionInfo::Agile { info, .. } = parsed else {
        panic!("expected Agile EncryptionInfo");
    };

    agile_verify_password(&info, password).expect("verify password");
    let secret_key = agile_secret_key(&info, password).expect("derive agile secret key");

    assert_eq!(
        info.key_data_block_size, 16,
        "expected test helper to use 16-byte block size"
    );

    let decrypted = agile_decrypt_package(&info, &secret_key, &encrypted_package)
        .expect("decrypt EncryptedPackage");

    assert_eq!(decrypted, plaintext);
    assert_zip_contains_workbook_xml(&decrypted);
}
