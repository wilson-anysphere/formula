use std::io::Write;

use formula_model::{PageSetup, Scaling};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_fit_to_page_without_setup_record() {
    let bytes = xls_fixture_builder::build_fit_to_page_without_setup_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let settings = workbook.sheet_print_settings_by_name("Sheet1");
    assert_eq!(
        settings.page_setup.scaling,
        Scaling::FitTo { width: 0, height: 0 }
    );

    // Only scaling is imported from BIFF8 sheet print settings today; other fields should retain
    // their model defaults.
    let mut expected = PageSetup::default();
    expected.scaling = Scaling::FitTo { width: 0, height: 0 };
    assert_eq!(settings.page_setup, expected);
}

