use std::io::Write;

use formula_model::Scaling;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_page_setup_margins_with_correct_record_order_precedence() {
    let bytes = xls_fixture_builder::build_page_setup_margins_before_setup_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let settings = workbook.sheet_print_settings_by_name("PageSetup");
    let margins = settings.page_setup.margins;

    // LEFT/RIGHT/TOP/BOTTOMMARGIN records must win over SETUP defaults regardless of record order.
    assert!((margins.left - 1.25).abs() < 1e-12, "left={}", margins.left);
    assert!((margins.right - 1.5).abs() < 1e-12, "right={}", margins.right);
    assert!((margins.top - 2.25).abs() < 1e-12, "top={}", margins.top);
    assert!((margins.bottom - 2.5).abs() < 1e-12, "bottom={}", margins.bottom);

    // Header/footer margins come from SETUP.numHdr/numFtr.
    assert!((margins.header - 0.25).abs() < 1e-12, "header={}", margins.header);
    assert!((margins.footer - 0.5).abs() < 1e-12, "footer={}", margins.footer);
}

#[test]
fn fit_to_page_overrides_scale_percent_when_enabled() {
    let bytes = xls_fixture_builder::build_page_setup_scaling_fit_to_page_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let settings = workbook.sheet_print_settings_by_name("ScaleFitTo");
    assert_eq!(
        settings.page_setup.scaling,
        Scaling::FitTo {
            width: 2,
            height: 3
        }
    );
}

#[test]
fn scale_percent_used_when_fit_to_page_disabled() {
    let bytes = xls_fixture_builder::build_page_setup_scaling_percent_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let settings = workbook.sheet_print_settings_by_name("ScalePercent");
    assert_eq!(settings.page_setup.scaling, Scaling::Percent(77));
}

