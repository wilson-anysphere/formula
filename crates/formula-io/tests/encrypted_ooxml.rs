use std::io::Cursor;

use formula_io::{
    detect_workbook_encryption, detect_workbook_format, open_workbook, open_workbook_model, Error,
    WorkbookEncryptionKind,
};

fn encrypted_ooxml_bytes() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo stream");
    ole.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage stream");
    ole.into_inner().into_inner()
}

#[test]
fn detects_encrypted_ooxml_xlsx_container() {
    let tmp = tempfile::tempdir().expect("tempdir");

    let bytes = encrypted_ooxml_bytes();

    // Test both correct and incorrect extensions to ensure content sniffing detects encryption
    // before attempting to open as legacy BIFF.
    for filename in ["encrypted.xlsx", "encrypted.xls", "encrypted.xlsb"] {
        let path = tmp.path().join(filename);
        std::fs::write(&path, &bytes).expect("write encrypted fixture");

        let info = detect_workbook_encryption(&path)
            .expect("detect encryption")
            .expect("expected encrypted workbook to be detected");
        assert_eq!(info.kind, WorkbookEncryptionKind::OoxmlOleEncryptedPackage);

        let err = detect_workbook_format(&path).expect_err("expected encrypted workbook to error");
        assert!(
            matches!(err, Error::EncryptedWorkbook { .. }),
            "expected Error::EncryptedWorkbook, got {err:?}"
        );

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
}

#[test]
fn detects_encrypted_ooxml_xlsx_container_for_model_loader() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let bytes = encrypted_ooxml_bytes();

    for filename in ["encrypted.xlsx", "encrypted.xls"] {
        let path = tmp.path().join(filename);
        std::fs::write(&path, &bytes).expect("write encrypted fixture");

        let err = open_workbook_model(&path).expect_err("expected encrypted workbook to error");
        assert!(
            matches!(err, Error::EncryptedWorkbook { .. }),
            "expected Error::EncryptedWorkbook, got {err:?}"
        );
    }
}
