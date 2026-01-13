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
fn clamps_invalid_fit_to_dimensions_from_setup_record() {
    let bytes = xls_fixture_builder::build_fit_to_clamp_fixture_xls();
    let result = import_fixture(&bytes);

    let settings = result.workbook.sheet_print_settings_by_name("Sheet1");
    assert_eq!(
        settings.page_setup.scaling,
        Scaling::FitTo {
            width: 32767,
            height: 32767
        }
    );

    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("iFitWidth")),
        "expected iFitWidth warning, got {:?}",
        result.warnings
    );
    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("iFitHeight")),
        "expected iFitHeight warning, got {:?}",
        result.warnings
    );
}
