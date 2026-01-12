use std::io::Write;

use formula_model::Range;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_autofilter_range_from_filterdatabase_defined_name() {
    let bytes = xls_fixture_builder::build_autofilter_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result.workbook.sheet_by_name("Filter").expect("Filter missing");

    let af = sheet.auto_filter.as_ref().expect("auto_filter missing");
    assert_eq!(af.range, Range::from_a1("A1:C5").unwrap());
    assert!(af.filter_columns.is_empty());
    assert!(af.sort_state.is_none());
    assert!(af.raw_xml.is_empty());
}

