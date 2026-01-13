use std::io::Write;

use formula_model::autofilter::SortCondition;
use formula_model::Range;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_autofilter_sort_state_from_biff_sort_record() {
    let bytes = xls_fixture_builder::build_autofilter_sort_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("FilterSort")
        .expect("FilterSort missing");

    let auto_filter = sheet.auto_filter.as_ref().expect("auto_filter missing");
    assert!(auto_filter.sort_state.is_some(), "expected sort_state");

    let sort_state = auto_filter.sort_state.as_ref().unwrap();
    assert_eq!(
        sort_state.conditions,
        vec![SortCondition {
            // OOXML semantics: when the sort range includes a header row, the sort key range
            // excludes it (start at row 2).
            range: Range::from_a1("B2:B5").unwrap(),
            descending: true,
        }]
    );
}

