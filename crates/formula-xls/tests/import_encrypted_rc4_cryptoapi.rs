use std::io::{Cursor, Read, Write};
use std::path::PathBuf;

use formula_model::{CellRef, CellValue, VerticalAlignment};

const PASSWORD: &str = "correct horse battery staple";

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted_rc4_cryptoapi_biff8.xls")
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

#[test]
fn decrypts_rc4_cryptoapi_biff8_xls() {
    let result = formula_xls::import_xls_path_with_password(fixture_path(), PASSWORD)
        .expect("expected decrypt + import to succeed");

    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));

    // Ensure workbook-global style metadata *after* FILEPASS was imported.
    //
    // In an encrypted workbook, XF/font/format records are located after FILEPASS and their
    // payload bytes are encrypted. After decryption we still retain the FILEPASS record header,
    // so we must mask it to allow BIFF parsing to continue and import the XF table.
    let cell = sheet
        .cell(CellRef::from_a1("A1").unwrap())
        .expect("A1 cell exists");
    assert_ne!(cell.style_id, 0, "expected BIFF-derived style id");

    let style = result
        .workbook
        .styles
        .get(cell.style_id)
        .expect("style exists for A1");
    assert!(
        style.alignment.is_some(),
        "expected at least one XF-derived alignment property to be imported"
    );
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
fn rc4_cryptoapi_wrong_password_errors() {
    let err = formula_xls::import_xls_path_with_password(fixture_path(), "wrong password")
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
    assert!(matches!(
        err,
        formula_xls::ImportError::Decrypt(formula_xls::DecryptError::UnsupportedEncryption)
    ));
}
