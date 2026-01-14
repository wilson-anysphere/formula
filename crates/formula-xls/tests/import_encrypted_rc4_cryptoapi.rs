use std::io::{Cursor, Read, Write};
use std::path::PathBuf;

use formula_model::CellValue;

const PASSWORD: &str = "correct horse battery staple";
const UNICODE_PASSWORD: &str = "pässwörd";

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted")
        .join("biff8_rc4_cryptoapi_pw_open.xls")
}

fn unicode_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted")
        .join("biff8_rc4_cryptoapi_unicode_pw_open.xls")
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

fn patch_filepass_cryptoapi_alg_id(workbook_stream: &mut [u8], new_alg_id: u32) {
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
            // FILEPASS payload:
            //   u16 wEncryptionType
            //   u16 wEncryptionSubType
            //   u32 dwEncryptionInfoLen
            //   EncryptionInfo bytes...
            // EncryptionInfo:
            //   u16 MajorVersion
            //   u16 MinorVersion
            //   u32 Flags
            //   u32 HeaderSize
            //   EncryptionHeader (HeaderSize bytes) where AlgID lives at offset 8.
            let payload = workbook_stream
                .get_mut(data_start..data_end)
                .expect("FILEPASS payload in range");
            assert!(
                payload.len() >= 32,
                "expected FILEPASS payload to contain CryptoAPI EncryptionInfo"
            );

            let header_size = u32::from_le_bytes([
                payload[16],
                payload[17],
                payload[18],
                payload[19],
            ]) as usize;
            let header_start = 20usize;
            assert!(
                header_start + header_size <= payload.len(),
                "EncryptionHeader out of range (header_size={header_size}, payload_len={})",
                payload.len()
            );

            let alg_id_off = header_start + 8;
            payload[alg_id_off..alg_id_off + 4].copy_from_slice(&new_alg_id.to_le_bytes());
            return;
        }

        offset = data_end;
    }

    panic!("FILEPASS record not found");
}

fn find_record_offset_from(workbook_stream: &[u8], start: usize, target_id: u16) -> Option<usize> {
    let mut offset = start;
    while offset + 4 <= workbook_stream.len() {
        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len =
            u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;
        let data_end = offset + 4 + len;
        if data_end > workbook_stream.len() {
            return None;
        }
        if record_id == target_id {
            return Some(offset);
        }
        offset = data_end;
    }
    None
}

#[test]
fn decrypts_rc4_cryptoapi_biff8_xls() {
    const RECORD_FILEPASS: u16 = 0x002F;
    const RECORD_WINDOW1: u16 = 0x003D;

    // Sanity-check the fixture structure: the `WINDOW1` record we want to validate is stored *after*
    // the `FILEPASS` record, meaning its payload is encrypted and can only be decoded once BIFF
    // parsing continues past `FILEPASS` on the decrypted stream.
    let bytes = std::fs::read(fixture_path()).expect("read fixture");
    let workbook_stream = read_workbook_stream_from_xls_bytes(&bytes);
    let filepass_offset =
        find_record_offset_from(&workbook_stream, 0, RECORD_FILEPASS).expect("FILEPASS record");
    let filepass_len = u16::from_le_bytes([
        workbook_stream[filepass_offset + 2],
        workbook_stream[filepass_offset + 3],
    ]) as usize;
    let after_filepass = filepass_offset + 4 + filepass_len;
    let window1_offset =
        find_record_offset_from(&workbook_stream, after_filepass, RECORD_WINDOW1)
            .expect("WINDOW1 record after FILEPASS");
    assert!(
        window1_offset > filepass_offset,
        "expected WINDOW1 to appear after FILEPASS in fixture stream"
    );

    let result = formula_xls::import_xls_path_with_password(fixture_path(), PASSWORD)
        .expect("expected decrypt + import to succeed");

    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));

    // Ensure workbook-global metadata *after* FILEPASS was imported.
    //
    // We use the active tab index (`WINDOW1.iTabCur`) as the canary for "did BIFF workbook-globals
    // parsing continue after FILEPASS on the decrypted stream?".
    //
    // Without masking/ignoring `FILEPASS`, the workbook-globals parser stops at that record and
    // never sees `WINDOW1`, leaving `active_sheet_id` unset.
    assert_eq!(
        result.workbook.view.active_sheet_id,
        Some(sheet.id),
        "expected BIFF-derived active sheet id"
    );
}

#[test]
fn rc4_cryptoapi_wrong_password_errors() {
    let err = formula_xls::import_xls_path_with_password(fixture_path(), "wrong password")
        .expect_err("expected wrong password error");
    assert!(matches!(err, formula_xls::ImportError::InvalidPassword));
}

#[test]
fn decrypts_rc4_cryptoapi_biff8_xls_with_unicode_password() {
    let result = formula_xls::import_xls_path_with_password(unicode_fixture_path(), UNICODE_PASSWORD)
        .expect("expected decrypt + import to succeed");
    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));
}

#[test]
fn rc4_cryptoapi_unicode_password_wrong_password_errors() {
    let err = formula_xls::import_xls_path_with_password(unicode_fixture_path(), "wrong password")
        .expect_err("expected wrong password error");
    assert!(matches!(err, formula_xls::ImportError::InvalidPassword));
}

#[test]
fn rc4_cryptoapi_unicode_password_different_normalization_fails() {
    // NFC password is "pässwörd" (U+00E4, U+00F6). NFD decomposes those into combining marks.
    let nfd = "pa\u{0308}sswo\u{0308}rd";
    assert_ne!(
        nfd, UNICODE_PASSWORD,
        "strings should differ before UTF-16 encoding"
    );

    let err = formula_xls::import_xls_path_with_password(unicode_fixture_path(), nfd)
        .expect_err("expected wrong password error");
    assert!(matches!(
        err,
        formula_xls::ImportError::Decrypt(formula_xls::DecryptError::WrongPassword)
    ));
}

#[test]
fn rc4_cryptoapi_unsupported_algorithm_errors() {
    // Patch the fixture FILEPASS header to claim AES-128 instead of RC4.
    const CALG_AES_128: u32 = 0x0000_660E;

    let bytes = std::fs::read(fixture_path()).expect("read fixture");
    let mut workbook_stream = read_workbook_stream_from_xls_bytes(&bytes);
    patch_filepass_cryptoapi_alg_id(&mut workbook_stream, CALG_AES_128);
    let patched_xls = build_xls_from_workbook_stream(&workbook_stream);

    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&patched_xls).expect("write xls bytes");

    let err = formula_xls::import_xls_path_with_password(tmp.path(), PASSWORD)
        .expect_err("expected unsupported encryption error");
    let msg = err.to_string();
    assert!(
        matches!(err, formula_xls::ImportError::UnsupportedEncryption(_)),
        "expected ImportError::UnsupportedEncryption, got {err:?} ({msg})"
    );
    assert!(
        msg.contains("AlgID") || msg.contains("algorithm") || msg.contains("AES"),
        "expected unsupported-encryption error message to mention algorithm; got: {msg}"
    );
}
