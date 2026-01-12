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

#[test]
fn ole_container_with_xls_extension_falls_back_to_xls_open_error() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("document.xls");
    std::fs::write(&path, non_xls_ole_bytes()).expect("write ole bytes");

    // Content sniffing should not classify non-Excel OLE containers as `.xls`.
    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Unknown);

    // `open_workbook` should still attempt to open based on the `.xls` extension and surface an
    // `.xls`-specific open error, rather than incorrectly claiming the extension is unsupported.
    let err = open_workbook(&path).expect_err("expected open_workbook to fail");
    assert!(
        matches!(err, Error::OpenXls { .. }),
        "expected OpenXls, got {err:?}"
    );

    let err = open_workbook_model(&path).expect_err("expected open_workbook_model to fail");
    assert!(
        matches!(err, Error::OpenXls { .. }),
        "expected OpenXls, got {err:?}"
    );
}

#[test]
fn ole_container_with_xlt_or_xla_extension_falls_back_to_xls_open_error() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let bytes = non_xls_ole_bytes();

    for ext in ["xlt", "xla"] {
        let path = tmp.path().join(format!("document.{ext}"));
        std::fs::write(&path, &bytes).expect("write ole bytes");

        // Content sniffing should not classify non-Excel OLE containers as `.xls`.
        let fmt = detect_workbook_format(&path).expect("detect format");
        assert_eq!(fmt, WorkbookFormat::Unknown);

        // `.xlt`/`.xla` are legacy BIFF/OLE extensions; we should still attempt to open them as
        // `.xls` and surface an `.xls`-specific open error, rather than claiming the extension is
        // unsupported.
        let err = open_workbook(&path).expect_err("expected open_workbook to fail");
        assert!(
            matches!(err, Error::OpenXls { .. }),
            "expected OpenXls, got {err:?}"
        );

        let err = open_workbook_model(&path).expect_err("expected open_workbook_model to fail");
        assert!(
            matches!(err, Error::OpenXls { .. }),
            "expected OpenXls, got {err:?}"
        );
    }
}

#[test]
fn ole_container_with_xlsx_extension_falls_back_to_xlsx_open_error() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("document.xlsx");
    std::fs::write(&path, non_xls_ole_bytes()).expect("write ole bytes");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Unknown);

    // `.xlsx` is a supported extension, so attempting to open it should yield an `.xlsx` open
    // error (invalid ZIP/package), not UnsupportedExtension.
    let err = open_workbook(&path).expect_err("expected open_workbook to fail");
    assert!(
        matches!(err, Error::OpenXlsx { .. }),
        "expected OpenXlsx, got {err:?}"
    );

    let err = open_workbook_model(&path).expect_err("expected open_workbook_model to fail");
    assert!(
        matches!(err, Error::OpenXlsx { .. }),
        "expected OpenXlsx, got {err:?}"
    );
}
