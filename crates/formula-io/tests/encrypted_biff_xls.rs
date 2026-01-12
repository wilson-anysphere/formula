use std::io::{Cursor, Write};

use formula_io::{open_workbook, open_workbook_model, Error};

fn encrypted_biff_xls_bytes() -> Vec<u8> {
    // Minimal workbook stream: FILEPASS + EOF.
    // - 0x002F FILEPASS indicates BIFF encryption/password protection.
    // - We don't need a valid BOF; the importer detects FILEPASS before parsing anything else.
    let workbook_stream = [
        0x002Fu16.to_le_bytes(),
        0u16.to_le_bytes(),
        0x000Au16.to_le_bytes(),
        0u16.to_le_bytes(),
    ]
    .concat();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

#[test]
fn errors_on_encrypted_biff_xls_filepass() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted_biff.xls");
    std::fs::write(&path, encrypted_biff_xls_bytes()).expect("write encrypted fixture");

    let err = open_workbook(&path).expect_err("expected encrypted workbook to error");
    assert!(
        matches!(err, Error::EncryptedWorkbook { .. }),
        "expected Error::EncryptedWorkbook, got {err:?}"
    );
    let msg = err.to_string().to_lowercase();
    assert!(
        msg.contains("encrypted") || msg.contains("password"),
        "expected error message to mention encryption/password protection, got: {msg}"
    );

    let err = open_workbook_model(&path).expect_err("expected encrypted workbook to error");
    assert!(
        matches!(err, Error::EncryptedWorkbook { .. }),
        "expected Error::EncryptedWorkbook, got {err:?}"
    );
}

