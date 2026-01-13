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
fn imports_page_break_edge_cases_and_warns() {
    let bytes = xls_fixture_builder::build_page_break_edge_cases_fixture_xls();
    let result = import_fixture(&bytes);

    let formula_xls::XlsImportResult { workbook, warnings, .. } = result;

    let settings = workbook.sheet_print_settings_by_name("PageBreaks");
    assert_eq!(
        settings.manual_page_breaks.row_breaks_after,
        BTreeSet::from([4u32])
    );
    assert_eq!(
        settings.manual_page_breaks.col_breaks_after,
        BTreeSet::from([2u32])
    );

    assert!(
        warnings
            .iter()
            .any(|w| w.message.contains("ignoring horizontal page break with row=0")),
        "expected row=0 break warning, got {warnings:?}"
    );
    assert!(
        warnings
            .iter()
            .any(|w| w.message.contains("ignoring vertical page break with col=0")),
        "expected col=0 break warning, got {warnings:?}"
    );
}

