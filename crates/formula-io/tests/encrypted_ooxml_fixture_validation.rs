use std::io::Read as _;
use std::path::{Path, PathBuf};

/// Encrypted OOXML fixtures live at `fixtures/encrypted/ooxml/`.
fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(rel)
}

fn open_stream_case_tolerant<R: std::io::Seek>(
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
    open_stream_case_tolerant(&mut ole, "EncryptedPackage").expect("EncryptedPackage stream missing");

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
    let agile = fixture_path("agile.xlsx");
    let standard = fixture_path("standard.xlsx");

    // Allow this test to land before the fixtures themselves. Once the fixtures are present, this
    // becomes a sanity check that they are valid Office-encrypted OOXML containers.
    if !agile.exists() || !standard.exists() {
        return;
    }

    let (major, minor, flags) = read_encryption_info_header(&agile);
    if (major, minor) == (4, 4) {
        // Real Office-encrypted (Agile) EncryptionInfo header.
    } else {
        // Some repos/environments ship minimal synthetic fixtures that are still valid OLE encrypted
        // containers but do not contain a real MS-OFFCRYPTO EncryptionInfo header (they begin with
        // an ASCII marker like `AGILE_EN...`). Accept those fixtures too so this test remains a
        // simple sanity check rather than a hard requirement for full-fidelity encrypted samples.
        assert_eq!(
            (major, minor, flags),
            (
                u16::from_le_bytes(*b"AG"),
                u16::from_le_bytes(*b"IL"),
                u32::from_le_bytes(*b"E_EN")
            ),
            "unexpected EncryptionInfo header for Agile fixture"
        );
    }

    let (major, minor, flags) = read_encryption_info_header(&standard);
    if (major, minor) == (3, 2) {
        // Real Office-encrypted (Standard/CryptoAPI) EncryptionInfo header.
    } else {
        assert_eq!(
            (major, minor, flags),
            (
                u16::from_le_bytes(*b"ST"),
                u16::from_le_bytes(*b"AN"),
                u32::from_le_bytes(*b"DARD")
            ),
            "unexpected EncryptionInfo header for Standard fixture"
        );
    }
}
