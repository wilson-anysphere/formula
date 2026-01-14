use std::io::{Cursor, Read, Write};
use std::path::PathBuf;

use formula_model::{CellRef, CellValue, VerticalAlignment};

const PASSWORD: &str = "password";
const UNICODE_PASSWORD: &str = "pässwörd";
const UNICODE_EMOJI_PASSWORD: &str = "pässwörd🔒";

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted")
        .join("biff8_rc4_standard_pw_open.xls")
}

fn unicode_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted")
        .join("biff8_rc4_standard_unicode_pw_open.xls")
}

fn unicode_emoji_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted")
        .join("biff8_rc4_standard_unicode_emoji_pw_open.xls")
}

fn basic_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("basic.xls")
}

fn read_workbook_stream_from_xls_bytes(data: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(data.to_vec());
    let mut ole = cfb::CompoundFile::open(cursor).expect("open xls cfb");

    for candidate in ["/Workbook", "/Book", "Workbook", "Book"] {
        if let Ok(mut stream) = ole.open_stream(candidate) {
            let mut buf = Vec::new();
            stream.read_to_end(&mut buf).expect("read workbook stream");
            return buf;
        }
    }

    panic!("fixture missing Workbook/Book stream");
}

fn build_xls_from_workbook_stream(workbook_stream: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole =
        cfb::CompoundFile::create_with_version(cfb::Version::V3, cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

fn sum_payload_bytes_after_filepass(workbook_stream: &[u8]) -> usize {
    const RECORD_FILEPASS: u16 = 0x002F;

    // Find FILEPASS record.
    let mut offset = 0usize;
    let mut filepass_end = None::<usize>;
    while offset + 4 <= workbook_stream.len() {
        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len =
            u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;
        let data_start = offset + 4;
        let data_end = data_start + len;
        assert!(data_end <= workbook_stream.len(), "truncated record while scanning");

        if record_id == RECORD_FILEPASS {
            filepass_end = Some(data_end);
            break;
        }

        offset = data_end;
    }
    let mut offset = filepass_end.expect("FILEPASS not found");

    // Sum payload lengths of subsequent records (headers are not encrypted).
    let mut sum = 0usize;
    while offset + 4 <= workbook_stream.len() {
        let len =
            u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;
        let data_start = offset + 4;
        let data_end = data_start + len;
        if data_end > workbook_stream.len() {
            break;
        }
        sum = sum.saturating_add(len);
        offset = data_end;
    }

    sum
}

fn patch_filepass_rc4_standard_minor_version(workbook_stream: &mut [u8], new_minor: u16) {
    const RECORD_FILEPASS: u16 = 0x002F;

    let mut offset = 0usize;
    while offset + 4 <= workbook_stream.len() {
        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len =
            u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;
        let data_start = offset + 4;
        let data_end = data_start + len;
        assert!(data_end <= workbook_stream.len(), "truncated record while scanning");

        if record_id == RECORD_FILEPASS {
            let payload = workbook_stream
                .get_mut(data_start..data_end)
                .expect("FILEPASS payload in range");
            assert!(
                payload.len() >= 6,
                "expected FILEPASS payload to include (type, major, minor)"
            );
            // Standard RC4 FILEPASS layout:
            // u16 wEncryptionType, u16 major, u16 minor, ...
            payload[4..6].copy_from_slice(&new_minor.to_le_bytes());
            return;
        }

        offset = data_end;
    }

    panic!("FILEPASS record not found");
}

#[test]
fn decrypts_rc4_standard_biff8_xls() {
    let bytes = std::fs::read(fixture_path()).expect("read fixture");
    let workbook_stream = read_workbook_stream_from_xls_bytes(&bytes);
    let payload_bytes_after_filepass = sum_payload_bytes_after_filepass(&workbook_stream);
    assert!(
        payload_bytes_after_filepass > 1024,
        "expected encrypted record-data stream to cross 1024-byte boundary; got {payload_bytes_after_filepass} bytes after FILEPASS"
    );

    let result = formula_xls::import_xls_path_with_password(fixture_path(), Some(PASSWORD))
        .expect("expected decrypt + import to succeed");

    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));

    // Ensure workbook-global style metadata *after* FILEPASS was imported.
    let cell = sheet
        .cell(CellRef::from_a1("A1").unwrap())
        .expect("A1 cell exists");
    assert_ne!(cell.style_id, 0, "expected BIFF-derived style id");

    let style = result
        .workbook
        .styles
        .get(cell.style_id)
        .expect("style exists for A1");
    assert_eq!(
        style
            .alignment
            .as_ref()
            .and_then(|alignment| alignment.vertical),
        Some(VerticalAlignment::Top),
        "expected A1 style to preserve vertical alignment from the decrypted XF record"
    );
}

#[test]
fn rc4_standard_wrong_password_errors() {
    let err = formula_xls::import_xls_path_with_password(fixture_path(), Some("wrong password"))
        .expect_err("expected wrong password error");
    assert!(matches!(err, formula_xls::ImportError::InvalidPassword));
}

#[test]
fn decrypts_rc4_standard_biff8_xls_with_unicode_password() {
    let result = formula_xls::import_xls_path_with_password(unicode_fixture_path(), Some(UNICODE_PASSWORD))
        .expect("expected decrypt + import to succeed");
    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));
}

#[test]
fn rc4_standard_unicode_password_wrong_password_errors() {
    let err =
        formula_xls::import_xls_path_with_password(unicode_fixture_path(), Some("wrong password"))
            .expect_err("expected wrong password error");
    assert!(matches!(err, formula_xls::ImportError::InvalidPassword));
}

#[test]
fn rc4_standard_unicode_password_different_normalization_fails() {
    // NFC password is "pässwörd" (U+00E4, U+00F6). NFD decomposes those into combining marks.
    let nfd = "pa\u{0308}sswo\u{0308}rd";
    assert_ne!(
        nfd, UNICODE_PASSWORD,
        "strings should differ before UTF-16 encoding"
    );

    let err = formula_xls::import_xls_path_with_password(unicode_fixture_path(), Some(nfd))
        .expect_err("expected wrong password error");
    assert!(matches!(err, formula_xls::ImportError::InvalidPassword));
}

#[test]
fn decrypts_rc4_standard_biff8_xls_with_unicode_emoji_password() {
    let result = formula_xls::import_xls_path_with_password(
        unicode_emoji_fixture_path(),
        Some(UNICODE_EMOJI_PASSWORD),
    )
    .expect("expected decrypt + import to succeed");
    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));
}

#[test]
fn rc4_standard_unicode_emoji_password_wrong_password_errors() {
    let err = formula_xls::import_xls_path_with_password(
        unicode_emoji_fixture_path(),
        Some("wrong password"),
    )
    .expect_err("expected wrong password error");
    assert!(matches!(err, formula_xls::ImportError::InvalidPassword));
}

#[test]
fn rc4_standard_unsupported_version_errors() {
    // Patch the fixture FILEPASS header to claim an unsupported minor version.
    let bytes = std::fs::read(fixture_path()).expect("read fixture");
    let mut workbook_stream = read_workbook_stream_from_xls_bytes(&bytes);
    patch_filepass_rc4_standard_minor_version(&mut workbook_stream, 3);
    let patched_xls = build_xls_from_workbook_stream(&workbook_stream);

    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&patched_xls).expect("write xls bytes");

    let err = formula_xls::import_xls_path_with_password(tmp.path(), Some(PASSWORD))
        .expect_err("expected unsupported encryption error");
    assert!(matches!(
        err,
        formula_xls::ImportError::UnsupportedEncryption(_)
    ));
}

#[test]
fn import_unencrypted_xls_with_password_is_unaffected() {
    // Supplying a password for a non-encrypted workbook should be ignored and import should still
    // succeed.
    let result = formula_xls::import_xls_path_with_password(basic_fixture_path(), Some(PASSWORD))
        .expect("import xls");

    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1");
    assert_eq!(
        sheet.value_a1("A1").unwrap(),
        CellValue::String("Hello".to_owned())
    );
}
