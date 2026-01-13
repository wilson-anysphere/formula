use std::io::{Cursor, Write};

use formula_model::CellRef;

mod common;

use common::xls_fixture_builder;

fn xls_bytes_from_workbook_stream(workbook_stream: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

#[test]
fn imports_workbook_globals_after_filepass_when_record_is_masked() {
    let workbook_stream =
        xls_fixture_builder::build_number_format_workbook_stream_with_filepass(false);

    // Sanity check: without masking, our normal import path should treat the file as encrypted and
    // abort (even though the bytes after FILEPASS are plaintext in this synthetic fixture).
    let bytes_unmasked = xls_bytes_from_workbook_stream(&workbook_stream);
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&bytes_unmasked).expect("write xls bytes");
    let err = formula_xls::import_xls_path(tmp.path()).expect_err("expected encrypted workbook");
    assert!(matches!(err, formula_xls::ImportError::EncryptedWorkbook));

    // Simulate post-decryption plumbing: mask FILEPASS in-memory before running BIFF parsing or
    // constructing the in-memory CFB for calamine.
    let mut decrypted_stream = workbook_stream.clone();
    let masked = formula_xls::mask_biff_filepass_record_id(&mut decrypted_stream);
    assert_eq!(masked, 1, "expected exactly one FILEPASS record to be masked");

    let bytes = xls_bytes_from_workbook_stream(&decrypted_stream);
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&bytes).expect("write xls bytes");
    let result = formula_xls::import_xls_path(tmp.path()).expect("import xls");

    // Bound sheet names should be preserved.
    let sheet = result
        .workbook
        .sheet_by_name("Formats")
        .expect("Formats missing");

    // Workbook-global metadata that appears *after* FILEPASS should be imported, including custom
    // number formats and the XF table.
    let a1 = CellRef::from_a1("A1").unwrap();
    let cell = sheet.cell(a1).expect("A1 missing");
    let fmt = result
        .workbook
        .styles
        .get(cell.style_id)
        .and_then(|s| s.number_format.as_deref());
    assert_eq!(fmt, Some("$#,##0.00"));
}

