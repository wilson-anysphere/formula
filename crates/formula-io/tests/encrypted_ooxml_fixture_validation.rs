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

const AGILE_FIXTURES: &[(&str, &str)] = &[
    ("agile.xlsx", "plaintext.xlsx"),
    ("agile-empty-password.xlsx", "plaintext.xlsx"),
    ("agile-unicode.xlsx", "plaintext.xlsx"),
    ("agile-unicode-excel.xlsx", "plaintext-excel.xlsx"),
    ("agile-large.xlsx", "plaintext-large.xlsx"),
    ("agile-basic.xlsm", "plaintext-basic.xlsm"),
];

const STANDARD_FIXTURES: &[(&str, &str)] = &[
    ("standard.xlsx", "plaintext.xlsx"),
    ("standard-4.2.xlsx", "plaintext.xlsx"),
    ("standard-rc4.xlsx", "plaintext.xlsx"),
    ("standard-unicode.xlsx", "plaintext.xlsx"),
    ("standard-large.xlsx", "plaintext-large.xlsx"),
    ("standard-basic.xlsm", "plaintext-basic.xlsm"),
];

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
    let mut stream = open_stream_case_tolerant(&mut ole, name)
        .unwrap_or_else(|_| panic!("{name} stream missing in {}", path.display()));
    let mut out = Vec::new();
    stream.read_to_end(&mut out).expect("read stream bytes");
    out
}

fn assert_agile_encryption_info_contains_encryption_xml(path: &Path) {
    let bytes = read_stream_bytes(path, "EncryptionInfo");
    assert!(
        bytes.len() >= 8,
        "EncryptionInfo stream too short in {} (len={})",
        path.display(),
        bytes.len()
    );

    // bytes[0..4] == version 4.4
    let major = u16::from_le_bytes([bytes[0], bytes[1]]);
    let minor = u16::from_le_bytes([bytes[2], bytes[3]]);
    assert_eq!(
        (major, minor),
        (4, 4),
        "Agile-encrypted OOXML should have EncryptionInfo version 4.4 ({})",
        path.display()
    );

    // bytes[8..] should contain `<encryption` (allowing length prefixes, BOM, UTF-16, etc).
    assert!(
        contains_utf8_or_utf16_bytes(&bytes[8..], "<encryption"),
        "expected EncryptionInfo to contain `<encryption` (UTF-8/UTF-16) after the 8-byte version header ({})",
        path.display()
    );

    // Best-effort: if we can find the UTF-8 `<encryption` root, parse the XML and validate the
    // expected namespace root element.
    let after_header = &bytes[8..];
    let marker = b"<encryption";
    if let Some(marker_pos) = after_header
        .windows(marker.len())
        .position(|window| window == marker)
    {
        let mut xml_bytes: &[u8] = &after_header[marker_pos..];

        // Trim trailing whitespace/NULs.
        while let Some((&b, rest)) = xml_bytes.split_last() {
            if b == 0 || b.is_ascii_whitespace() {
                xml_bytes = rest;
                continue;
            }
            break;
        }

        if let Ok(xml_str) = std::str::from_utf8(xml_bytes) {
            let doc = roxmltree::Document::parse(xml_str).unwrap_or_else(|err| {
                panic!(
                    "Agile EncryptionInfo XML should parse as XML ({}): {err}",
                    path.display()
                )
            });

            let root = doc.root_element();
            assert_eq!(
                (root.tag_name().namespace(), root.tag_name().name()),
                (
                    Some("http://schemas.microsoft.com/office/2006/encryption"),
                    "encryption"
                ),
                "Agile EncryptionInfo XML should have the expected `<encryption>` root element ({})",
                path.display()
            );
        }
    }
}

#[test]
fn encrypted_ooxml_fixtures_have_expected_encryption_info_versions() {
    for (name, _) in AGILE_FIXTURES {
        let path = fixture_path(name);
        let (major, minor, _flags) = read_encryption_info_header(&path);
        assert_eq!(
            (major, minor),
            (4, 4),
            "Agile-encrypted OOXML should have EncryptionInfo version 4.4 ({name})"
        );
    }

    for (name, _) in STANDARD_FIXTURES {
        let path = fixture_path(name);
        let (major, minor, _flags) = read_encryption_info_header(&path);
        assert!(
            minor == 2 && matches!(major, 2 | 3 | 4),
            "Standard-encrypted OOXML should have EncryptionInfo version *.2 with major=2/3/4 ({name}); got {major}.{minor}"
        );
    }
}

#[test]
fn encrypted_ooxml_fixtures_have_encryptedpackage_size_prefix_matching_plaintext_fixture_size() {
    for (name, plaintext_name) in AGILE_FIXTURES.iter().chain(STANDARD_FIXTURES.iter()) {
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
}

#[test]
fn agile_encrypted_ooxml_fixtures_contain_encryptioninfo_xml_descriptor() {
    for (name, _) in AGILE_FIXTURES {
        let path = fixture_path(name);
        assert_agile_encryption_info_contains_encryption_xml(&path);
    }
}

#[test]
fn standard_encrypted_ooxml_fixtures_do_not_contain_agile_xml_near_start() {
    for (name, _) in STANDARD_FIXTURES {
        let path = fixture_path(name);
        let info = read_stream_bytes(&path, "EncryptionInfo");
        assert!(
            info.len() >= 8,
            "{name} EncryptionInfo stream is too short ({} bytes)",
            info.len()
        );

        // Standard EncryptionInfo is binary; it should not look like Agile XML near the start.
        let preview_len = info.len().saturating_sub(8).min(512);
        assert!(
            !contains_utf8_or_utf16_bytes(&info[8..8 + preview_len], "<encryption"),
            "expected {name} EncryptionInfo not to contain `<encryption` near the start"
        );
    }
}
