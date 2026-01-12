use std::io::Cursor;

use formula_io::{open_workbook, Error};

#[test]
fn detects_encrypted_ooxml_xlsx_container() {
    let tmp = tempfile::tempdir().expect("tempdir");

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo stream");
    ole.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage stream");
    let bytes = ole.into_inner().into_inner();

    // Test both correct and incorrect extensions to ensure content sniffing detects encryption
    // before attempting to open as legacy BIFF.
    for filename in ["encrypted.xlsx", "encrypted.xls"] {
        let path = tmp.path().join(filename);
        std::fs::write(&path, &bytes).expect("write encrypted fixture");

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
    }
}
