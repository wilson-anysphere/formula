use std::collections::BTreeSet;
use std::io::Write;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_manual_page_breaks_from_biff8_sheet_stream() {
    let bytes = xls_fixture_builder::build_manual_page_breaks_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let settings = workbook.sheet_print_settings_by_name("Sheet1");
    assert_eq!(
        settings.manual_page_breaks.row_breaks_after,
        BTreeSet::from([1u32, 4u32])
    );
    assert_eq!(
        settings.manual_page_breaks.col_breaks_after,
        BTreeSet::from([2u32, 9u32])
    );
}

