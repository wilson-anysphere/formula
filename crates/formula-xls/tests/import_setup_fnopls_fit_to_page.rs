use std::io::Write;

use formula_model::{Orientation, PaperSize, Scaling};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_setup_fnopls_does_not_drop_fit_to_or_header_footer_margins() {
    let bytes = xls_fixture_builder::build_setup_fnopls_fit_to_page_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let settings = workbook.sheet_print_settings_by_name("Sheet1");
    assert_eq!(
        settings.page_setup.scaling,
        Scaling::FitTo {
            width: 2,
            height: 3
        }
    );
    assert!(
        (settings.page_setup.margins.header - 0.9).abs() < 1e-12,
        "expected header margin 0.9, got {}",
        settings.page_setup.margins.header
    );
    assert!(
        (settings.page_setup.margins.footer - 1.1).abs() < 1e-12,
        "expected footer margin 1.1, got {}",
        settings.page_setup.margins.footer
    );

    // fNoPls=1 marks paper size and orientation as undefined, so the importer should leave them at
    // defaults.
    assert_eq!(settings.page_setup.paper_size, PaperSize::LETTER);
    assert_eq!(settings.page_setup.orientation, Orientation::Portrait);
}

