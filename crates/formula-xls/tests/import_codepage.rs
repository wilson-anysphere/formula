use std::io::Write;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_workbook_codepage_from_biff_codepage_record() {
    let bytes = xls_fixture_builder::build_note_comment_split_across_continues_codepage_932_fixture_xls();
    let result = import_fixture(&bytes);
    assert_eq!(result.workbook.codepage, 932);
}

