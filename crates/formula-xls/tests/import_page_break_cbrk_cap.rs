use std::io::Write;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn caps_page_break_cbrk_and_imports_single_entry() {
    let bytes = xls_fixture_builder::build_page_break_cbrk_cap_fixture_xls();
    let result = import_fixture(&bytes);

    let settings = result.workbook.sheet_print_settings_by_name("PageBreaks");
    assert_eq!(
        settings.manual_page_breaks.row_breaks_after.len(),
        1,
        "expected exactly one row break, got {:?}",
        settings.manual_page_breaks.row_breaks_after
    );
    assert!(
        settings.manual_page_breaks.row_breaks_after.contains(&0),
        "expected row break after 0, got {:?}",
        settings.manual_page_breaks.row_breaks_after
    );
    assert!(
        settings.manual_page_breaks.col_breaks_after.is_empty(),
        "expected no column breaks, got {:?}",
        settings.manual_page_breaks.col_breaks_after
    );

    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("cbrk") || w.message.contains("count")),
        "expected warning mentioning cbrk/count, got {:?}",
        result.warnings
    );
}

