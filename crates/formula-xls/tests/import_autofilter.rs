use std::io::Write;

use formula_model::Range;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

fn import_fixture_without_biff(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path_without_biff(tmp.path()).expect("import xls")
}

#[test]
fn imports_autofilter_range_from_filterdatabase_defined_name() {
    let bytes = xls_fixture_builder::build_defined_names_builtins_fixture_xls();
    let result = import_fixture(&bytes);
    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");

    let auto_filter = sheet
        .auto_filter
        .as_ref()
        .expect("expected Sheet1.auto_filter to be set");

    assert_eq!(auto_filter.range, Range::from_a1("A1:C10").unwrap());
}

#[test]
fn imports_autofilter_fixture_range_and_empty_state() {
    let bytes = xls_fixture_builder::build_autofilter_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result.workbook.sheet_by_name("Filter").expect("Filter missing");
    let auto_filter = sheet.auto_filter.as_ref().expect("auto_filter missing");

    assert_eq!(auto_filter.range, Range::from_a1("A1:C5").unwrap());
    assert!(auto_filter.filter_columns.is_empty());
    assert!(auto_filter.sort_state.is_none());
    assert!(auto_filter.raw_xml.is_empty());
}

#[test]
fn imports_autofilter_range_from_filterdatabase_defined_name_without_biff() {
    let bytes = xls_fixture_builder::build_autofilter_fixture_xls();
    let result = import_fixture_without_biff(&bytes);

    let sheet = result.workbook.sheet_by_name("Filter").expect("Filter missing");

    let af = sheet.auto_filter.as_ref().unwrap_or_else(|| {
        panic!(
            "auto_filter missing; defined_names={:?}; warnings={:?}",
            result.workbook.defined_names, result.warnings
        )
    });
    assert_eq!(af.range, Range::from_a1("A1:C5").unwrap());
    assert!(af.filter_columns.is_empty());
    assert!(af.sort_state.is_none());
    assert!(af.raw_xml.is_empty());
}
