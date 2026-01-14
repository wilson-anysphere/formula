use std::io::Write;

use formula_model::{Orientation, PageMargins, Scaling};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_truncated_setup_record_best_effort() {
    let bytes = xls_fixture_builder::build_truncated_setup_fixture_xls();
    let result = import_fixture(&bytes);

    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("truncated SETUP record")),
        "expected truncated SETUP warning, warnings={:?}",
        result.warnings
    );

    let settings = result.workbook.sheet_print_settings_by_name("Sheet1");
    let page_setup = settings.page_setup;

    assert_eq!(page_setup.paper_size.code, 9, "expected A4 paper size");
    assert_eq!(page_setup.orientation, Orientation::Landscape);
    assert_eq!(page_setup.scaling, Scaling::Percent(80));

    assert_eq!(page_setup.margins.header, 0.55);
    assert_eq!(page_setup.margins.footer, PageMargins::default().footer);
}

#[test]
fn ignores_truncated_wsbool_record_and_emits_warning() {
    let bytes = xls_fixture_builder::build_truncated_wsbool_fixture_xls();
    let result = import_fixture(&bytes);

    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("truncated WSBOOL record")),
        "expected truncated WSBOOL warning, warnings={:?}",
        result.warnings
    );

    let settings = result.workbook.sheet_print_settings_by_name("Sheet1");
    assert_eq!(settings.page_setup.scaling, Scaling::Percent(80));
}

#[test]
fn truncated_margin_record_is_ignored_but_later_valid_margin_applies() {
    let bytes = xls_fixture_builder::build_truncated_margin_fixture_xls();
    let result = import_fixture(&bytes);

    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("truncated TOPMARGIN record")),
        "expected truncated TOPMARGIN warning, warnings={:?}",
        result.warnings
    );

    let settings = result.workbook.sheet_print_settings_by_name("Sheet1");
    assert_eq!(settings.page_setup.margins.top, 2.25);
}

