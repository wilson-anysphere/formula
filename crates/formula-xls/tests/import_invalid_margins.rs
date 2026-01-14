use std::io::Write;

use formula_model::PageMargins;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn ignores_invalid_biff_margins_and_warns() {
    let bytes = xls_fixture_builder::build_invalid_margins_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = &result.workbook;

    let margins = workbook.sheet_print_settings_by_name("Margins").page_setup.margins;

    // LEFTMARGIN is out-of-range in the fixture and should be ignored (default retained).
    assert_eq!(margins.left, PageMargins::default().left);

    // Sanity: other margins should still be imported.
    assert_eq!(margins.right, 1.2);
    assert_eq!(margins.header, 0.4);
    assert_eq!(margins.footer, 0.5);

    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("invalid") && w.message.contains("LEFTMARGIN")),
        "expected invalid LEFTMARGIN warning, got {:?}",
        result.warnings
    );
}

#[test]
fn ignores_nan_biff_margins_and_warns() {
    let bytes = xls_fixture_builder::build_invalid_margins_nan_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = &result.workbook;

    let margins = workbook.sheet_print_settings_by_name("Margins").page_setup.margins;

    // LEFTMARGIN is NaN in the fixture and should be ignored (default retained).
    assert_eq!(margins.left, PageMargins::default().left);

    // Sanity: other margins should still be imported.
    assert_eq!(margins.right, 1.2);
    assert_eq!(margins.header, 0.4);
    assert_eq!(margins.footer, 0.5);

    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("invalid") && w.message.contains("LEFTMARGIN")),
        "expected invalid LEFTMARGIN warning, got {:?}",
        result.warnings
    );
}
