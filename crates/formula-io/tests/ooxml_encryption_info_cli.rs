use std::io::{Cursor, Write as _};
use std::process::Command;

fn ooxml_encryption_info_bin() -> &'static str {
    // Cargo sets `CARGO_BIN_EXE_<name>` for integration tests. Binary names may contain `-`,
    // but some environments/tools normalize them to `_`, so accept either.
    option_env!("CARGO_BIN_EXE_ooxml-encryption-info")
        .or(option_env!("CARGO_BIN_EXE_ooxml_encryption_info"))
        .expect("ooxml-encryption-info binary should be built for integration tests")
}

fn make_ooxml_encrypted_container(major: u16, minor: u16, flags: u32, payload: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut stream = ole
            .create_stream("EncryptionInfo")
            .expect("create EncryptionInfo stream");
        stream.write_all(&major.to_le_bytes()).unwrap();
        stream.write_all(&minor.to_le_bytes()).unwrap();
        stream.write_all(&flags.to_le_bytes()).unwrap();
        stream.write_all(payload).unwrap();
    }

    // The payload is irrelevant for detection; it just needs to exist.
    ole.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage stream");

    ole.into_inner().into_inner()
}

#[test]
fn cli_errors_on_non_ole_input() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("not_ole.bin");
    std::fs::write(&path, b"hello").expect("write fixture");

    let out = Command::new(ooxml_encryption_info_bin())
        .arg(&path)
        .output()
        .expect("run cli");
    assert!(
        !out.status.success(),
        "expected non-zero exit status, got {:?}",
        out.status.code()
    );
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert!(
        stderr.to_lowercase().contains("not an ole"),
        "expected a short OLE/CFB error message, got: {stderr}"
    );
}

#[test]
fn cli_errors_on_ole_without_encryption_streams() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("empty.cfb");

    let cursor = Cursor::new(Vec::new());
    let ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    let bytes = ole.into_inner().into_inner();
    std::fs::write(&path, bytes).expect("write fixture");

    let out = Command::new(ooxml_encryption_info_bin())
        .arg(&path)
        .output()
        .expect("run cli");
    assert!(
        !out.status.success(),
        "expected non-zero exit status, got {:?}",
        out.status.code()
    );
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert!(
        stderr
            .to_lowercase()
            .contains("not an ooxml encrypted container"),
        "expected encryption stream error, got: {stderr}"
    );
}

#[test]
fn cli_prints_standard_version() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("standard.xlsx");

    let bytes = make_ooxml_encrypted_container(3, 2, 0xAABBCCDD, b"");
    std::fs::write(&path, bytes).expect("write fixture");

    let out = Command::new(ooxml_encryption_info_bin())
        .arg(&path)
        .output()
        .expect("run cli");
    assert!(
        out.status.success(),
        "expected success exit status, got {:?}",
        out.status.code()
    );
    let stdout = String::from_utf8_lossy(&out.stdout);
    assert_eq!(
        stdout.trim_end(),
        "Standard (3.2) flags=0xaabbccdd",
        "unexpected stdout: {stdout}"
    );
}

#[test]
fn cli_prints_agile_version_and_root_tag() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("agile.xlsx");

    let xml = br#"<?xml version="1.0" encoding="UTF-8"?><encryption></encryption>"#;
    let bytes = make_ooxml_encrypted_container(4, 4, 0, xml);
    std::fs::write(&path, bytes).expect("write fixture");

    let out = Command::new(ooxml_encryption_info_bin())
        .arg(&path)
        .output()
        .expect("run cli");
    assert!(
        out.status.success(),
        "expected success exit status, got {:?}",
        out.status.code()
    );
    let stdout = String::from_utf8_lossy(&out.stdout);
    assert!(
        stdout.contains("Agile (4.4)"),
        "expected Agile version, got: {stdout}"
    );
    assert!(
        stdout.contains("xml_root=encryption"),
        "expected xml_root detection, got: {stdout}"
    );
}
