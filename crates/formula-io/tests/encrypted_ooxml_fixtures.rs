use std::io::Read;
use std::path::Path;

use formula_io::{detect_workbook_format, Error};

fn assert_encrypted_ooxml_bytes_detected(bytes: &[u8], stem: &str) {
    let tmp = tempfile::tempdir().expect("tempdir");

    // Test both correct and incorrect extensions to ensure content sniffing detects encryption
    // before attempting to open as legacy BIFF.
    for ext in ["xlsx", "xls", "xlsb"] {
        let path = tmp.path().join(format!("{stem}.{ext}"));
        std::fs::write(&path, bytes).expect("write encrypted fixture");

        let err = detect_workbook_format(&path).expect_err("expected encrypted workbook to error");
        if cfg!(feature = "encrypted-workbooks") {
            assert!(
                matches!(err, Error::PasswordRequired { .. }),
                "expected Error::PasswordRequired, got {err:?}"
            );
        } else {
            assert!(
                matches!(err, Error::UnsupportedEncryption { .. }),
                "expected Error::UnsupportedEncryption, got {err:?}"
            );
        }

        let msg = err.to_string().to_lowercase();
        assert!(
            msg.contains("encrypted") || msg.contains("password"),
            "expected error message to mention encryption/password protection, got: {msg}"
        );
    }
}

#[test]
fn detects_encrypted_ooxml_agile_fixture() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/encrypted/ooxml/agile.xlsx"
    ));

    let bytes = std::fs::read(fixture_path).expect("read agile encrypted fixture");
    assert_encrypted_ooxml_bytes_detected(&bytes, "agile");
}

#[test]
fn detects_encrypted_ooxml_standard_fixtures() {
    for (fixture, stem) in [("standard.xlsx", "standard"), ("standard-4.2.xlsx", "standard-4.2")] {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../fixtures/encrypted/ooxml/"
        ))
        .join(fixture);

        let bytes =
            std::fs::read(&fixture_path).unwrap_or_else(|err| panic!("read {fixture}: {err}"));
        assert_encrypted_ooxml_bytes_detected(&bytes, stem);
    }
}

#[test]
fn detects_encrypted_ooxml_standard_unicode_fixture() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/encrypted/ooxml/standard-unicode.xlsx"
    ));

    let bytes = std::fs::read(fixture_path).expect("read standard-unicode encrypted fixture");
    assert_encrypted_ooxml_bytes_detected(&bytes, "standard-unicode");
}

#[test]
fn detects_encrypted_ooxml_agile_empty_password_fixture() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/encrypted/ooxml/agile-empty-password.xlsx"
    ));

    let bytes = std::fs::read(fixture_path).expect("read agile-empty-password encrypted fixture");
    assert_encrypted_ooxml_bytes_detected(&bytes, "agile-empty-password");
}

#[test]
fn detects_encrypted_ooxml_agile_unicode_fixture() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/encrypted/ooxml/agile-unicode.xlsx"
    ));

    let bytes = std::fs::read(fixture_path).expect("read agile-unicode encrypted fixture");
    assert_encrypted_ooxml_bytes_detected(&bytes, "agile-unicode");
}

#[test]
fn detects_encrypted_ooxml_agile_unicode_excel_fixture() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/encrypted/ooxml/agile-unicode-excel.xlsx"
    ));

    let bytes = std::fs::read(fixture_path).expect("read agile-unicode-excel encrypted fixture");
    assert_encrypted_ooxml_bytes_detected(&bytes, "agile-unicode-excel");
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct StandardEncryptionInfoParams {
    version_major: u16,
    version_minor: u16,
    version_flags: u32,
    alg_id: u32,
    alg_id_hash: u32,
    key_size_bits: u32,
    provider_type: u32,
    csp_name: Option<String>,
    salt_size: u32,
}

fn read_u16_le(bytes: &[u8], pos: &mut usize, context: &'static str) -> Result<u16, String> {
    let end = pos.saturating_add(2);
    let slice = bytes
        .get(*pos..end)
        .ok_or_else(|| format!("EncryptionInfo truncated while reading {context}"))?;
    *pos = end;
    Ok(u16::from_le_bytes([slice[0], slice[1]]))
}

fn read_u32_le(bytes: &[u8], pos: &mut usize, context: &'static str) -> Result<u32, String> {
    let end = pos.saturating_add(4);
    let slice = bytes
        .get(*pos..end)
        .ok_or_else(|| format!("EncryptionInfo truncated while reading {context}"))?;
    *pos = end;
    Ok(u32::from_le_bytes([slice[0], slice[1], slice[2], slice[3]]))
}

fn parse_utf16le_z(bytes: &[u8], context: &'static str) -> Result<Option<String>, String> {
    if bytes.is_empty() {
        return Ok(None);
    }
    if bytes.len() % 2 != 0 {
        return Err(format!("{context} is not valid UTF-16LE (odd byte length)"));
    }
    let mut code_units: Vec<u16> = Vec::with_capacity(bytes.len() / 2);
    for chunk in bytes.chunks_exact(2) {
        code_units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }
    let end = code_units
        .iter()
        .position(|u| *u == 0)
        .unwrap_or(code_units.len());
    let s = String::from_utf16(&code_units[..end])
        .map_err(|_| format!("{context} is not valid UTF-16LE"))?;
    if s.is_empty() {
        Ok(None)
    } else {
        Ok(Some(s))
    }
}

fn parse_standard_encryption_info(
    encryption_info: &[u8],
) -> Result<StandardEncryptionInfoParams, String> {
    let mut pos = 0usize;

    let major = read_u16_le(encryption_info, &mut pos, "EncryptionVersionInfo.major")?;
    let minor = read_u16_le(encryption_info, &mut pos, "EncryptionVersionInfo.minor")?;
    let version_flags = read_u32_le(encryption_info, &mut pos, "EncryptionVersionInfo.flags")?;

    if minor != 2 || !matches!(major, 2 | 3 | 4) {
        return Err(format!(
            "expected Standard EncryptionInfo version *.2 with major=2/3/4, got {major}.{minor}"
        ));
    }

    let header_size =
        read_u32_le(encryption_info, &mut pos, "EncryptionInfo.header_size")? as usize;
    let header_bytes = encryption_info
        .get(pos..pos + header_size)
        .ok_or_else(|| "EncryptionInfo truncated while reading EncryptionHeader".to_string())?;
    pos += header_size;

    if header_bytes.len() < 8 * 4 {
        return Err(format!(
            "EncryptionHeader is too short: expected at least 32 bytes, got {}",
            header_bytes.len()
        ));
    }

    let mut hpos = 0usize;
    let _header_flags = read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.flags")?;
    let _size_extra = read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.sizeExtra")?;
    let alg_id = read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.algId")?;
    let alg_id_hash = read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.algIdHash")?;
    let key_size_bits = read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.keySize")?;
    let provider_type = read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.providerType")?;
    let _reserved1 = read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.reserved1")?;
    let _reserved2 = read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.reserved2")?;

    let csp_name = parse_utf16le_z(&header_bytes[hpos..], "EncryptionHeader.CSPName")?;

    // EncryptionVerifier occupies the remaining bytes after the header.
    let salt_size = read_u32_le(encryption_info, &mut pos, "EncryptionVerifier.saltSize")?;
    let salt_size_usize = salt_size as usize;
    if encryption_info.len() < pos + salt_size_usize {
        return Err(format!(
            "EncryptionVerifier.saltSize={salt_size} does not fit into remaining bytes"
        ));
    }

    Ok(StandardEncryptionInfoParams {
        version_major: major,
        version_minor: minor,
        version_flags,
        alg_id,
        alg_id_hash,
        key_size_bits,
        provider_type,
        csp_name,
        salt_size,
    })
}

#[test]
fn standard_fixtures_encryption_info_parameters_are_pinned() {
    let fixture_dir = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/encrypted/ooxml"
    ));

    let common_expected = StandardEncryptionInfoParams {
        version_major: 0, // filled per fixture below
        version_minor: 2,
        // MS-OFFCRYPTO `EncryptionVersionInfo.flags`: `fCryptoAPI` (0x04) + `fAES` (0x20).
        version_flags: 0x0000_0024,
        alg_id: 0x0000_660E,      // CALG_AES_128
        alg_id_hash: 0x0000_8004, // CALG_SHA1
        key_size_bits: 128,
        provider_type: 24, // PROV_RSA_AES
        csp_name: None,    // set per fixture below
        salt_size: 16,
    };

    for (fixture_name, version_major) in [
        ("standard.xlsx", 3),
        ("standard-basic.xlsm", 3),
        ("standard-large.xlsx", 3),
        // Apache POI Standard/CryptoAPI AES fixtures emit `EncryptionInfo` version 4.2.
        ("standard-4.2.xlsx", 4),
        ("standard-unicode.xlsx", 4),
    ] {
        let fixture_path = fixture_dir.join(fixture_name);

        let file = std::fs::File::open(&fixture_path)
            .unwrap_or_else(|err| panic!("open {fixture_name}: {err}"));
        let mut ole = cfb::CompoundFile::open(file)
            .unwrap_or_else(|err| panic!("parse {fixture_name}: {err}"));
        let mut stream = ole
            .open_stream("EncryptionInfo")
            .or_else(|_| ole.open_stream("/EncryptionInfo"))
            .unwrap_or_else(|err| panic!("open {fixture_name} EncryptionInfo stream: {err}"));
        let mut bytes = Vec::new();
        stream
            .read_to_end(&mut bytes)
            .unwrap_or_else(|err| panic!("read {fixture_name} EncryptionInfo stream: {err}"));

        let parsed = parse_standard_encryption_info(&bytes)
            .unwrap_or_else(|err| panic!("parse {fixture_name} Standard EncryptionInfo: {err}"));

        let expected = StandardEncryptionInfoParams {
            version_major,
            csp_name: Some("Microsoft Enhanced RSA and AES Cryptographic Provider".to_string()),
            ..common_expected.clone()
        };

        assert_eq!(
            parsed, expected,
            "{fixture_name}: Standard encryption fixture parameters drifted.\n\
             If this change is intentional, update:\n\
             - fixtures/encrypted/ooxml/README.md\n\
             - crates/formula-io/tests/encrypted_ooxml_fixtures.rs (this assertion)\n"
        );
    }
}
