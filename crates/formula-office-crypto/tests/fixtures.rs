use std::io::{Cursor, Read, Write};
use std::path::PathBuf;

use cfb::CompoundFile;
use formula_office_crypto::{decrypt_encrypted_package_ole, OfficeCryptoError};

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(name)
}

fn read_fixture(name: &str) -> Vec<u8> {
    std::fs::read(fixture_path(name)).unwrap_or_else(|err| panic!("read fixture {name}: {err}"))
}

fn read_ole_stream(bytes: &[u8], name: &str) -> Vec<u8> {
    let mut ole = CompoundFile::open(Cursor::new(bytes)).expect("open OLE");
    let mut stream = ole
        .open_stream(name)
        .or_else(|_| ole.open_stream(&format!("/{name}")))
        .unwrap_or_else(|err| panic!("open OLE stream {name}: {err}"));
    let mut out = Vec::new();
    stream
        .read_to_end(&mut out)
        .unwrap_or_else(|err| panic!("read OLE stream {name}: {err}"));
    out
}

#[test]
fn decrypts_standard_fixture_matches_plaintext() {
    let plaintext = read_fixture("plaintext.xlsx");
    let standard = read_fixture("standard.xlsx");

    let decrypted = decrypt_encrypted_package_ole(&standard, "password").expect("decrypt standard");
    assert_eq!(decrypted, plaintext);
    assert!(decrypted.starts_with(b"PK"));
}

#[test]
fn decrypts_standard_rc4_fixture_matches_plaintext() {
    let plaintext = read_fixture("plaintext.xlsx");
    let standard_rc4 = read_fixture("standard-rc4.xlsx");

    let decrypted =
        decrypt_encrypted_package_ole(&standard_rc4, "password").expect("decrypt standard rc4");
    assert_eq!(decrypted, plaintext);
    assert!(decrypted.starts_with(b"PK"));

    let err = decrypt_encrypted_package_ole(&standard_rc4, "wrong")
        .expect_err("wrong password should fail");
    assert!(matches!(err, OfficeCryptoError::InvalidPassword));
}

#[test]
fn decrypts_agile_fixture_matches_plaintext() {
    let plaintext = read_fixture("plaintext.xlsx");
    let agile = read_fixture("agile.xlsx");

    let decrypted = decrypt_encrypted_package_ole(&agile, "password").expect("decrypt agile");
    assert_eq!(decrypted, plaintext);
    assert!(decrypted.starts_with(b"PK"));
}

#[test]
fn decrypts_agile_basic_xlsm_fixture_matches_plaintext() {
    let plaintext = read_fixture("plaintext-basic.xlsm");
    let agile = read_fixture("agile-basic.xlsm");

    let decrypted = decrypt_encrypted_package_ole(&agile, "password").expect("decrypt agile xlsm");
    assert_eq!(decrypted, plaintext);
    assert!(decrypted.starts_with(b"PK"));
}

#[test]
fn decrypts_agile_fixture_without_data_integrity() {
    // Some non-Excel producers omit the `<dataIntegrity>` element. Formula should still be able to
    // decrypt, but must skip integrity verification.
    let plaintext = read_fixture("plaintext.xlsx");
    let agile = read_fixture("agile.xlsx");

    let mut encryption_info = read_ole_stream(&agile, "EncryptionInfo");
    let encrypted_package = read_ole_stream(&agile, "EncryptedPackage");

    let xml_start = encryption_info
        .iter()
        .position(|b| *b == b'<')
        .expect("EncryptionInfo must contain XML");
    let mut header = encryption_info[..xml_start].to_vec();
    let xml =
        std::str::from_utf8(&encryption_info[xml_start..]).expect("EncryptionInfo XML is UTF-8");

    let start = xml
        .find("<dataIntegrity")
        .expect("expected <dataIntegrity> element");
    let end = if let Some(end_rel) = xml[start..].find("/>") {
        start + end_rel + 2
    } else if let Some(end_rel) = xml[start..].find("</dataIntegrity>") {
        start + end_rel + "</dataIntegrity>".len()
    } else {
        panic!("expected </dataIntegrity> or />");
    };

    let mut patched_xml = String::new();
    patched_xml.push_str(&xml[..start]);
    patched_xml.push_str(&xml[end..]);

    // If this `EncryptionInfo` stream includes a 4-byte XML length prefix (common in Office
    // output), update it so the parser continues to treat the header as `... flags + xml_len`.
    if xml_start == 12 {
        let len = patched_xml.as_bytes().len() as u32;
        header[8..12].copy_from_slice(&len.to_le_bytes());
    }

    encryption_info = header.into_iter().chain(patched_xml.into_bytes()).collect();

    let cursor = Cursor::new(Vec::new());
    let mut ole = CompoundFile::create(cursor).expect("create OLE");
    ole.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo");
    ole.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage")
        .write_all(&encrypted_package)
        .expect("write EncryptedPackage");
    let patched_ole = ole.into_inner().into_inner();

    let decrypted = decrypt_encrypted_package_ole(&patched_ole, "password")
        .expect("decrypt without dataIntegrity");
    assert_eq!(decrypted, plaintext);
    assert!(decrypted.starts_with(b"PK"));
}

#[test]
fn standard_wrong_password_returns_invalid_password() {
    let standard = read_fixture("standard.xlsx");

    let err =
        decrypt_encrypted_package_ole(&standard, "wrong").expect_err("wrong password should fail");
    assert!(matches!(err, OfficeCryptoError::InvalidPassword));
}
