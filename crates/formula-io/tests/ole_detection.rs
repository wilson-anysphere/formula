use std::io::Cursor;

use formula_io::{detect_workbook_format, open_workbook, open_workbook_model, Error, WorkbookFormat};

fn non_xls_ole_bytes() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    // Create a stream that would exist in other OLE document types (e.g. Word), but not in an
    // Excel BIFF workbook.
    ole.create_stream("WordDocument")
        .expect("create WordDocument stream");
    ole.into_inner().into_inner()
}

#[test]
fn ole_container_without_workbook_stream_is_not_classified_as_xls() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("document.doc");
    std::fs::write(&path, non_xls_ole_bytes()).expect("write ole bytes");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Unknown);

    let err = open_workbook(&path).expect_err("expected open_workbook to fail");
    assert!(
        matches!(err, Error::UnsupportedExtension { .. }),
        "expected UnsupportedExtension, got {err:?}"
    );

    let err = open_workbook_model(&path).expect_err("expected open_workbook_model to fail");
    assert!(
        matches!(err, Error::UnsupportedExtension { .. }),
        "expected UnsupportedExtension, got {err:?}"
    );
}

