use std::io::{Cursor, Write};

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

fn non_biff_workbook_stream_ole_bytes() -> Vec<u8> {
    // OLE container with a stream named `Workbook`, but where the stream bytes do *not* start with
    // a BIFF BOF record.
    //
    // This guards against false-positive `.xls` classification for arbitrary OLE containers that
    // happen to contain a similarly named stream.
    const RECORD_FILEPASS: u16 = 0x002F;

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        // BIFF-like header that is *not* a BOF record:
        // [record_id=FILEPASS][len=0]
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&RECORD_FILEPASS.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        stream.write_all(&bytes).expect("write Workbook bytes");
    }
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
fn ole_container_with_workbook_stream_but_not_biff_is_not_classified_as_xls() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("document.doc");
    std::fs::write(&path, non_biff_workbook_stream_ole_bytes()).expect("write ole bytes");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Unknown);

    // Content sniffing should not route this through the `.xls` importer.
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
