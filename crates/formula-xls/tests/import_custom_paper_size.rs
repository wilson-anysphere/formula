use std::io::Write;

use formula_model::PaperSize;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn ignores_custom_paper_size_in_setup_record() {
    let bytes = xls_fixture_builder::build_custom_paper_size_fixture_xls();
    let result = import_fixture(&bytes);

    let settings = result.workbook.sheet_print_settings_by_name("Sheet1");
    assert_eq!(settings.page_setup.paper_size, PaperSize::LETTER);

    let warnings: Vec<&str> = result.warnings.iter().map(|w| w.message.as_str()).collect();
    assert!(
        warnings.iter().any(|w| {
            let w = w.to_ascii_lowercase();
            w.contains("paper size") && (w.contains("custom") || w.contains("invalid"))
        }),
        "expected custom/invalid paper size warning, got: {warnings:?}"
    );
}
