use std::io::Write;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn rewrites_sheet_names_in_formulas_after_sanitization() {
    // The fixture contains an invalid (over-long) sheet name and a formula in another sheet
    // referencing that name.
    let bytes = xls_fixture_builder::build_sheet_name_sanitization_formula_fixture_xls();
    let result = import_fixture(&bytes);

    // The invalid sheet name should be sanitized during import.
    let old_name = "A".repeat(40);
    let new_name = "A".repeat(31);
    assert!(result.workbook.sheet_by_name(&new_name).is_some());
    assert!(result.workbook.sheet_by_name(&old_name).is_none());

    // The cross-sheet formula should reference the sanitized name.
    let sheet = result.workbook.sheet_by_name("Ref").expect("Ref missing");
    let formula = sheet
        .formula_a1("A1")
        .expect("parse A1")
        .expect("formula missing");
    let expected = format!("{new_name}!A1");
    assert_eq!(formula, expected.as_str());
}
