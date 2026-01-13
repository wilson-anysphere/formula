use std::io::{Cursor, Write};

use formula_io::{open_workbook_model_with_password, open_workbook_with_password, Error};

fn record(record_id: u16, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(4 + payload.len());
    out.extend_from_slice(&record_id.to_le_bytes());
    out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    out.extend_from_slice(payload);
    out
}

fn encrypted_biff_xls_bytes_filepass() -> Vec<u8> {
    // Minimal BIFF stream:
    // - BOF (BIFF8) with dummy payload
    // - FILEPASS (0x002F) indicates workbook encryption/password protection
    // - EOF
    const RECORD_BOF_BIFF8: u16 = 0x0809;
    const RECORD_FILEPASS: u16 = 0x002F;
    const RECORD_EOF: u16 = 0x000A;

    let workbook_stream = [
        record(RECORD_BOF_BIFF8, &[0u8; 16]),
        record(RECORD_FILEPASS, &[]),
        record(RECORD_EOF, &[]),
    ]
    .concat();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream bytes");
    }
    ole.into_inner().into_inner()
}

#[test]
fn password_open_path_routes_encrypted_biff_xls_to_xls_importer() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let bytes = encrypted_biff_xls_bytes_filepass();

    // Test both correct and incorrect extensions to ensure content sniffing routes through the
    // `.xls` importer rather than failing early in format detection with `Error::EncryptedWorkbook`.
    for filename in ["encrypted.xls", "encrypted.xlsx", "encrypted.xlsb"] {
        let path = tmp.path().join(filename);
        std::fs::write(&path, &bytes).expect("write encrypted xls fixture");

        let err = open_workbook_with_password(&path, Some("password"))
            .expect_err("expected synthetic FILEPASS fixture to fail import");
        assert!(
            !matches!(err, Error::EncryptedWorkbook { .. }),
            "expected password-capable open path to reach `.xls` importer (not fail early with Error::EncryptedWorkbook), got {err:?}"
        );
        assert!(
            matches!(err, Error::OpenXls { .. }),
            "expected `.xls` importer error, got {err:?}"
        );

        let err = open_workbook_model_with_password(&path, Some("password"))
            .expect_err("expected synthetic FILEPASS fixture to fail import");
        assert!(
            !matches!(err, Error::EncryptedWorkbook { .. }),
            "expected password-capable open path to reach `.xls` importer (not fail early with Error::EncryptedWorkbook), got {err:?}"
        );
        assert!(
            matches!(err, Error::OpenXls { .. }),
            "expected `.xls` importer error, got {err:?}"
        );
    }
}
