use std::io::Write;

use formula_model::CalculationMode;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_biff_workbook_calc_settings() {
    let bytes = xls_fixture_builder::build_calc_settings_fixture_xls();
    let result = import_fixture(&bytes);

    let settings = &result.workbook.calc_settings;
    assert_eq!(settings.calculation_mode, CalculationMode::Manual);
    assert_eq!(settings.calculate_before_save, false);
    assert_eq!(settings.full_precision, false);
    assert_eq!(settings.full_calc_on_load, false);

    assert_eq!(settings.iterative.enabled, true);
    assert_eq!(settings.iterative.max_iterations, 7);
    assert!((settings.iterative.max_change - 0.01).abs() < 1e-12);
}

