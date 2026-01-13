use std::io::Read as _;
use std::path::{Path, PathBuf};

/// Encrypted OOXML fixtures live at `fixtures/encrypted/ooxml/`.
fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(rel)
}

fn open_stream_case_tolerant<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> std::io::Result<cfb::Stream<R>> {
    ole.open_stream(name)
        .or_else(|_| ole.open_stream(format!("/{name}")))
}

fn read_encryption_info_header(path: &Path) -> (u16, u16, u32) {
    let file = std::fs::File::open(path).expect("open fixture file");
    let mut ole = cfb::CompoundFile::open(file).expect("open cfb (OLE) container");

    // Assert encryption streams exist.
    open_stream_case_tolerant(&mut ole, "EncryptionInfo").expect("EncryptionInfo stream missing");
    open_stream_case_tolerant(&mut ole, "EncryptedPackage")
        .expect("EncryptedPackage stream missing");

    let mut stream =
        open_stream_case_tolerant(&mut ole, "EncryptionInfo").expect("open EncryptionInfo stream");
    let mut header = [0u8; 8];
    stream
        .read_exact(&mut header)
        .expect("read EncryptionInfo header");

    let major = u16::from_le_bytes([header[0], header[1]]);
    let minor = u16::from_le_bytes([header[2], header[3]]);
    let flags = u32::from_le_bytes([header[4], header[5], header[6], header[7]]);
    (major, minor, flags)
}

#[test]
fn encrypted_ooxml_fixtures_have_expected_encryption_info_versions() {
    for name in [
        "agile.xlsx",
        "agile-empty-password.xlsx",
        "agile-unicode.xlsx",
        "agile-large.xlsx",
        "agile-basic.xlsm",
    ] {
        let path = fixture_path(name);
        let (major, minor, _flags) = read_encryption_info_header(&path);
        assert_eq!(
            (major, minor),
            (4, 4),
            "Agile-encrypted OOXML should have EncryptionInfo version 4.4 ({name})"
        );
    }

    for name in [
        "standard.xlsx",
        "standard-large.xlsx",
        "standard-basic.xlsm",
    ] {
        let path = fixture_path(name);
        let (major, minor, _flags) = read_encryption_info_header(&path);
        assert_eq!(
            (major, minor),
            (3, 2),
            "Standard-encrypted OOXML should have EncryptionInfo version 3.2 ({name})"
        );
    }
}
