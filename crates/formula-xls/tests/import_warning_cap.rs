use std::io::Write;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn caps_total_import_warnings() {
    // Deliberately exceed the global warning cap (1000) by generating many BIFF view-state
    // warnings across many worksheets.
    let bytes = xls_fixture_builder::build_many_malformed_selection_records_fixture_xls(25, 60);
    let result = import_fixture(&bytes);

    // The importer should cap warnings and add a single suppression message.
    assert_eq!(result.warnings.len(), 1001, "warnings={:?}", result.warnings);
    assert_eq!(
        result.warnings.last().map(|w| w.message.as_str()),
        Some("additional `.xls` import warnings suppressed")
    );
    assert_eq!(
        result
            .warnings
            .iter()
            .filter(|w| w.message == "additional `.xls` import warnings suppressed")
            .count(),
        1
    );
}

#[test]
fn caps_total_import_warnings_from_sort_state_recovery() {
    // Deliberately exceed the global warning cap (1000) by generating many best-effort
    // sort-state warnings from malformed BIFF8 `SORT` records.
    //
    // This exercises the warning-cap enforcement on warning propagation paths that previously
    // bypassed `push_import_warning`.
    let bytes = xls_fixture_builder::build_many_sort_state_warnings_fixture_xls(1200);
    let result = import_fixture(&bytes);

    assert_eq!(result.warnings.len(), 1001, "warnings={:?}", result.warnings);
    assert_eq!(
        result.warnings.last().map(|w| w.message.as_str()),
        Some("additional `.xls` import warnings suppressed")
    );
    assert_eq!(
        result
            .warnings
            .iter()
            .filter(|w| w.message == "additional `.xls` import warnings suppressed")
            .count(),
        1
    );
}
