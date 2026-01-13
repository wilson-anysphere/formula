use std::io::Write;

use formula_model::{Orientation, Scaling};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_biff8_page_setup_percent_scaling_and_manual_page_breaks() {
    let bytes = xls_fixture_builder::build_page_setup_percent_scaling_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let settings = workbook.sheet_print_settings_by_name("Sheet1");

    assert_eq!(settings.page_setup.orientation, Orientation::Landscape);
    assert_eq!(settings.page_setup.paper_size.code, 9);
    assert_eq!(settings.page_setup.scaling, Scaling::Percent(85));

    assert!((settings.page_setup.margins.left - 0.5).abs() < 1e-9);
    assert!((settings.page_setup.margins.right - 0.6).abs() < 1e-9);
    assert!((settings.page_setup.margins.top - 0.7).abs() < 1e-9);
    assert!((settings.page_setup.margins.bottom - 0.8).abs() < 1e-9);
    assert!((settings.page_setup.margins.header - 0.9).abs() < 1e-9);
    assert!((settings.page_setup.margins.footer - 1.0).abs() < 1e-9);

    // HorzBrk.row is "first row below the break" (0-based) so row=5 => break after row 4.
    assert!(settings.manual_page_breaks.row_breaks_after.contains(&4));
    // VertBrk.col is "first col to the right of the break" (0-based) so col=3 => break after col 2.
    assert!(settings.manual_page_breaks.col_breaks_after.contains(&2));
}

#[test]
fn imports_biff8_page_setup_fit_to_scaling_and_manual_page_breaks() {
    let bytes = xls_fixture_builder::build_page_setup_fit_to_scaling_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let settings = workbook.sheet_print_settings_by_name("Sheet1");

    assert_eq!(settings.page_setup.orientation, Orientation::Landscape);
    assert_eq!(settings.page_setup.paper_size.code, 9);
    assert_eq!(
        settings.page_setup.scaling,
        Scaling::FitTo {
            width: 2,
            height: 3,
        }
    );

    assert!((settings.page_setup.margins.left - 0.5).abs() < 1e-9);
    assert!((settings.page_setup.margins.right - 0.6).abs() < 1e-9);
    assert!((settings.page_setup.margins.top - 0.7).abs() < 1e-9);
    assert!((settings.page_setup.margins.bottom - 0.8).abs() < 1e-9);
    assert!((settings.page_setup.margins.header - 0.9).abs() < 1e-9);
    assert!((settings.page_setup.margins.footer - 1.0).abs() < 1e-9);

    // HorzBrk.row is "first row below the break" (0-based) so row=5 => break after row 4.
    assert!(settings.manual_page_breaks.row_breaks_after.contains(&4));
    // VertBrk.col is "first col to the right of the break" (0-based) so col=3 => break after col 2.
    assert!(settings.manual_page_breaks.col_breaks_after.contains(&2));
}
