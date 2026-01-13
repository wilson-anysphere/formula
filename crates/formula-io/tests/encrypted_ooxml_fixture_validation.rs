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

fn contains_bytes(haystack: &[u8], needle: &[u8]) -> bool {
    if needle.is_empty() {
        return true;
    }
    haystack
        .windows(needle.len())
        .any(|window| window == needle)
}

fn contains_utf8_or_utf16_bytes(haystack: &[u8], needle: &str) -> bool {
    if contains_bytes(haystack, needle.as_bytes()) {
        return true;
    }

    let utf16le: Vec<u8> = needle
        .encode_utf16()
        .flat_map(|ch| ch.to_le_bytes())
        .collect();
    if contains_bytes(haystack, &utf16le) {
        return true;
    }

    let utf16be: Vec<u8> = needle
        .encode_utf16()
        .flat_map(|ch| ch.to_be_bytes())
        .collect();
    contains_bytes(haystack, &utf16be)
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

fn read_stream_bytes(path: &Path, name: &str) -> Vec<u8> {
    let file = std::fs::File::open(path).expect("open fixture file");
    let mut ole = cfb::CompoundFile::open(file).expect("open cfb (OLE) container");
    let mut stream = open_stream_case_tolerant(&mut ole, name).unwrap_or_else(|_| {
        panic!("{name} stream missing in {}", path.display());
    });
    let mut out = Vec::new();
    stream
        .read_to_end(&mut out)
        .expect("read stream bytes");
    out
}

#[test]
fn encrypted_ooxml_fixtures_have_expected_encryption_info_versions() {
    let agile_fixtures = [
        ("agile.xlsx", "plaintext.xlsx"),
        ("agile-empty-password.xlsx", "plaintext.xlsx"),
        ("agile-unicode.xlsx", "plaintext.xlsx"),
        ("agile-large.xlsx", "plaintext-large.xlsx"),
        ("agile-basic.xlsm", "plaintext-basic.xlsm"),
    ];
    let standard_fixtures = [
        ("standard.xlsx", "plaintext.xlsx"),
        ("standard-rc4.xlsx", "plaintext.xlsx"),
        ("standard-large.xlsx", "plaintext-large.xlsx"),
        ("standard-basic.xlsm", "plaintext-basic.xlsm"),
    ];

    for (name, _) in agile_fixtures {
        let path = fixture_path(name);
        let (major, minor, _flags) = read_encryption_info_header(&path);
        assert_eq!(
            (major, minor),
            (4, 4),
            "Agile-encrypted OOXML should have EncryptionInfo version 4.4 ({name})"
        );
    }

    for (name, _) in standard_fixtures {
        let path = fixture_path(name);
        let (major, minor, _flags) = read_encryption_info_header(&path);
        assert!(
            minor == 2 && matches!(major, 2 | 3 | 4),
            "Standard-encrypted OOXML should have EncryptionInfo version *.2 with major=2/3/4 ({name}); got {major}.{minor}"
        );
    }

    // --- Additional fixture sanity checks (real-world structure). ---

    // 1) EncryptedPackage begins with an 8-byte little-endian u64 package size prefix matching the
    // corresponding plaintext workbook byte length.
    for (name, plaintext_name) in agile_fixtures.iter().chain(standard_fixtures.iter()) {
        let encrypted_path = fixture_path(name);
        let plaintext_path = fixture_path(plaintext_name);
        let plaintext_len = std::fs::metadata(&plaintext_path)
            .unwrap_or_else(|_| panic!("stat {plaintext_name}"))
            .len();

        let encrypted_package = read_stream_bytes(&encrypted_path, "EncryptedPackage");
        assert!(
            encrypted_package.len() >= 8,
            "{name} EncryptedPackage stream is too short ({} bytes)",
            encrypted_package.len()
        );
        let declared_plaintext_len = u64::from_le_bytes(
            encrypted_package[0..8]
                .try_into()
                .expect("slice length checked"),
        );
        assert_eq!(
            declared_plaintext_len, plaintext_len,
            "{name} EncryptedPackage plaintext size prefix mismatch (expected {plaintext_len}, got {declared_plaintext_len})"
        );
    }

    // 2) Agile EncryptionInfo should contain the `<encryption` root tag bytes (best-effort search).
    for (name, _) in agile_fixtures {
        let path = fixture_path(name);
        let info = read_stream_bytes(&path, "EncryptionInfo");
        assert!(
            info.len() >= 8,
            "{name} EncryptionInfo stream is too short ({} bytes)",
            info.len()
        );
        assert!(
            contains_utf8_or_utf16_bytes(&info[8..], "<encryption"),
            "expected {name} EncryptionInfo to contain `<encryption` (UTF-8/UTF-16) after the 8-byte version header"
        );
    }

    // 3) Standard EncryptionInfo is binary and should not look like Agile XML near the start
    // (heuristic).
    for (name, _) in standard_fixtures {
        let path = fixture_path(name);
        let info = read_stream_bytes(&path, "EncryptionInfo");
        assert!(
            info.len() >= 8,
            "{name} EncryptionInfo stream is too short ({} bytes)",
            info.len()
        );
        let preview_len = info.len().saturating_sub(8).min(512);
        assert!(
            !contains_utf8_or_utf16_bytes(&info[8..8 + preview_len], "<encryption"),
            "expected {name} EncryptionInfo not to contain `<encryption` near the start"
        );
    }
}
