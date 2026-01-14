use std::io::{Cursor, Write as _};
use std::path::PathBuf;
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

fn make_ooxml_encrypted_container_with_leading_slash_paths(
    major: u16,
    minor: u16,
    flags: u32,
    payload: &[u8],
) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut stream = ole
            .create_stream("/EncryptionInfo")
            .expect("create /EncryptionInfo stream");
        stream.write_all(&major.to_le_bytes()).unwrap();
        stream.write_all(&minor.to_le_bytes()).unwrap();
        stream.write_all(&flags.to_le_bytes()).unwrap();
        stream.write_all(payload).unwrap();
    }

    ole.create_stream("/EncryptedPackage")
        .expect("create /EncryptedPackage stream");

    ole.into_inner().into_inner()
}

fn make_standard_encryption_header_payload(
    header_flags: u32,
    alg_id: u32,
    alg_id_hash: u32,
    key_size: u32,
    provider_type: u32,
) -> Vec<u8> {
    let header_size = 8u32 * 4;
    let mut payload = Vec::new();
    payload.extend_from_slice(&header_size.to_le_bytes());
    // EncryptionHeader (first 8 DWORDs).
    payload.extend_from_slice(&header_flags.to_le_bytes()); // flags
    payload.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    payload.extend_from_slice(&alg_id.to_le_bytes()); // algId
    payload.extend_from_slice(&alg_id_hash.to_le_bytes()); // algIdHash
    payload.extend_from_slice(&key_size.to_le_bytes()); // keySize (bits)
    payload.extend_from_slice(&provider_type.to_le_bytes()); // providerType
    payload.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    payload.extend_from_slice(&0u32.to_le_bytes()); // reserved2
    payload
}

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(rel)
}

#[test]
fn cli_errors_on_missing_args() {
    let out = Command::new(ooxml_encryption_info_bin())
        .output()
        .expect("run cli");
    assert!(
        !out.status.success(),
        "expected non-zero exit status, got {:?}",
        out.status.code()
    );
    assert_eq!(
        out.status.code(),
        Some(2),
        "expected usage exit code 2, got {:?}",
        out.status.code()
    );
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert!(
        stderr.to_lowercase().contains("usage:"),
        "expected usage message, got: {stderr}"
    );
}

#[test]
fn cli_errors_on_extra_args() {
    let out = Command::new(ooxml_encryption_info_bin())
        .arg("a.xlsx")
        .arg("b.xlsx")
        .output()
        .expect("run cli");
    assert!(
        !out.status.success(),
        "expected non-zero exit status, got {:?}",
        out.status.code()
    );
    assert_eq!(
        out.status.code(),
        Some(2),
        "expected usage exit code 2, got {:?}",
        out.status.code()
    );
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert!(
        stderr.to_lowercase().contains("usage:"),
        "expected usage message, got: {stderr}"
    );
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

    // Minimal Standard payload: headerSize + 8 DWORDs of EncryptionHeader fixed fields.
    let mut payload = Vec::new();
    payload.extend_from_slice(&(32u32).to_le_bytes()); // headerSize
    payload.extend_from_slice(&(0x0000_0024u32).to_le_bytes()); // EncryptionHeader.flags (fCryptoAPI | fAES)
    payload.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    payload.extend_from_slice(&0x0000_660Eu32.to_le_bytes()); // algId (AES-128)
    payload.extend_from_slice(&0x0000_8004u32.to_le_bytes()); // algIdHash (SHA1)
    payload.extend_from_slice(&128u32.to_le_bytes()); // keySize (bits)
    payload.extend_from_slice(&0u32.to_le_bytes()); // providerType
    payload.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    payload.extend_from_slice(&0u32.to_le_bytes()); // reserved2

    let bytes = make_ooxml_encrypted_container(3, 2, 0xAABBCCDD, &payload);
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
        "Standard (3.2) flags=0xaabbccdd hdr_flags=0x00000024 fCryptoAPI=1 fAES=1 algId=0x0000660e algIdHash=0x00008004 keySize=128",
        "unexpected stdout: {stdout}"
    );
}

#[test]
fn cli_prints_standard_version_major_2() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("standard.xlsx");

    let bytes = make_ooxml_encrypted_container(2, 2, 0xAABBCCDD, b"");
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
        "Standard (2.2) flags=0xaabbccdd",
        "unexpected stdout: {stdout}"
    );
}

#[test]
fn cli_prints_standard_version_major_4() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("standard.xlsx");

    let bytes = make_ooxml_encrypted_container(4, 2, 0xAABBCCDD, b"");
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
        "Standard (4.2) flags=0xaabbccdd",
        "unexpected stdout: {stdout}"
    );
}

#[test]
fn cli_prints_standard_encryption_header_in_verbose_mode() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("standard.xlsx");

    let payload = make_standard_encryption_header_payload(
        0x0000_0004 | 0x0000_0020, // fCryptoAPI | fAES
        0x0000_660E,               // CALG_AES_128
        0x0000_8004,               // CALG_SHA1
        128,
        0x0000_0018, // PROV_RSA_AES
    );
    let bytes = make_ooxml_encrypted_container(3, 2, 0, &payload);
    std::fs::write(&path, bytes).expect("write fixture");

    let out = Command::new(ooxml_encryption_info_bin())
        .arg("--verbose")
        .arg(&path)
        .output()
        .expect("run cli");
    assert!(
        out.status.success(),
        "expected success exit status, got {:?}",
        out.status.code()
    );
    let stdout = String::from_utf8_lossy(&out.stdout);
    let lines: Vec<_> = stdout.lines().collect();
    assert_eq!(
        lines,
        vec![
            "Standard (3.2) flags=0x00000000 hdr_flags=0x00000024 fCryptoAPI=1 fAES=1 algId=0x0000660e algIdHash=0x00008004 keySize=128",
            "EncryptionHeader.flags=0x00000024 fCryptoAPI=true fAES=true",
            "EncryptionHeader.algId=0x0000660e",
            "EncryptionHeader.algIdHash=0x00008004",
            "EncryptionHeader.keySize=128",
            "EncryptionHeader.providerType=0x00000018",
        ],
        "unexpected stdout: {stdout}"
    );
}

#[test]
fn cli_errors_on_truncated_standard_stream_in_verbose_mode() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("standard.xlsx");

    // Standard version header but no payload after the 8-byte version info.
    let bytes = make_ooxml_encrypted_container(3, 2, 0, b"");
    std::fs::write(&path, bytes).expect("write fixture");

    let out = Command::new(ooxml_encryption_info_bin())
        .arg("-v")
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
        stderr.to_lowercase().contains("truncated"),
        "expected truncated error message, got: {stderr}"
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

#[test]
fn cli_prints_agile_version_for_utf16le_xml() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("agile_utf16.xlsx");

    let xml = r#"<?xml version="1.0" encoding="UTF-16"?><encryption></encryption>"#;
    let mut xml_bytes = Vec::new();
    // UTF-16LE BOM.
    xml_bytes.extend_from_slice(&[0xFF, 0xFE]);
    for u in xml.encode_utf16() {
        xml_bytes.extend_from_slice(&u.to_le_bytes());
    }

    let bytes = make_ooxml_encrypted_container(4, 4, 0, &xml_bytes);
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
        "expected xml_root detection for UTF-16 XML, got: {stdout}"
    );
}

#[test]
fn cli_accepts_leading_slash_stream_paths() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("agile_slash.xlsx");

    let xml = br#"<?xml version="1.0" encoding="UTF-8"?><encryption></encryption>"#;
    let bytes = make_ooxml_encrypted_container_with_leading_slash_paths(4, 4, 0, xml);
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
}

#[test]
fn cli_prints_extensible_version() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("extensible.xlsx");

    let bytes = make_ooxml_encrypted_container(3, 3, 0x11223344, b"");
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
        "Extensible (3.3) flags=0x11223344",
        "unexpected stdout: {stdout}"
    );
}

#[test]
fn cli_prints_unknown_version() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("unknown.xlsx");

    let bytes = make_ooxml_encrypted_container(1, 1, 0, b"");
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
        "Unknown (1.1) flags=0x00000000",
        "unexpected stdout: {stdout}"
    );
}

#[test]
fn cli_reports_expected_versions_for_repo_fixtures() {
    let agile = fixture_path("agile.xlsx");
    let standard = fixture_path("standard.xlsx");

    let out = Command::new(ooxml_encryption_info_bin())
        .arg(&agile)
        .output()
        .expect("run cli on agile fixture");
    assert!(
        out.status.success(),
        "expected success exit status for agile fixture, got {:?}",
        out.status.code()
    );
    let stdout = String::from_utf8_lossy(&out.stdout);
    let stdout = stdout.trim_end();
    assert!(
        stdout.starts_with("Agile (4.4) flags=0x00000040"),
        "unexpected stdout for agile fixture: {stdout}"
    );
    if stdout.contains("xml_root=") {
        assert!(
            stdout.contains("xml_root=encryption"),
            "unexpected xml root tag for agile fixture: {stdout}"
        );
    }

    let out = Command::new(ooxml_encryption_info_bin())
        .arg(&standard)
        .output()
        .expect("run cli on standard fixture");
    assert!(
        out.status.success(),
        "expected success exit status for standard fixture, got {:?}",
        out.status.code()
    );
    let stdout = String::from_utf8_lossy(&out.stdout);
    let stdout = stdout.trim_end();
    assert_eq!(
        stdout,
        "Standard (3.2) flags=0x00000024 hdr_flags=0x00000024 fCryptoAPI=1 fAES=1 algId=0x0000660e algIdHash=0x00008004 keySize=128",
        "unexpected stdout for standard fixture: {stdout}"
    );
}
