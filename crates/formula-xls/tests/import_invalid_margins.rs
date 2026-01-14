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
fn ignores_invalid_page_margin_values_and_emits_warnings() {
    let bytes = xls_fixture_builder::build_invalid_margins_fixture_xls();
    let result = import_fixture(&bytes);

    let settings = result.workbook.sheet_print_settings_by_name("Sheet1");
    let margins = settings.page_setup.margins;
    let defaults = PageMargins::default();

    // Left margin has a valid record followed by an invalid record; invalid must not clobber.
    assert!((margins.left - 1.0).abs() < 1e-12, "left={}", margins.left);

    // Margins populated only by invalid values should remain default.
    assert_eq!(margins.right, defaults.right);
    assert_eq!(margins.top, defaults.top);
    assert_eq!(margins.bottom, defaults.bottom);
    assert_eq!(margins.header, defaults.header);
    assert_eq!(margins.footer, defaults.footer);

    let warning_messages: Vec<&str> = result.warnings.iter().map(|w| w.message.as_str()).collect();
    for needle in [
        "LEFTMARGIN",
        "RIGHTMARGIN",
        "TOPMARGIN",
        "BOTTOMMARGIN",
        "SETUP header margin",
        "SETUP footer margin",
    ] {
        assert!(
            warning_messages.iter().any(|w| w.contains(needle)),
            "expected warning containing {needle:?}, got warnings={warning_messages:?}"
        );
    }
}
