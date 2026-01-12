use std::io::Write;

use formula_engine::{parse_formula, ParseOptions};
use formula_model::CellRef;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

fn assert_parseable(formula_body: &str) {
    let formula = format!("={formula_body}");
    parse_formula(&formula, ParseOptions::default())
        .unwrap_or_else(|e| panic!("expected formula to be parseable, formula={formula:?}, err={e:?}"));
}

#[test]
fn rewrites_cross_sheet_formulas_after_sheet_name_truncation() {
    // The fixture contains an invalid (over-long) sheet name and a formula in another sheet
    // referencing that name.
    let bytes = xls_fixture_builder::build_formula_sheet_name_truncation_fixture_xls();
    let result = import_fixture(&bytes);

    // The invalid sheet name should be sanitized during import.
    let old_name = "A".repeat(40);
    let new_name = "A".repeat(formula_model::EXCEL_MAX_SHEET_NAME_LEN);
    assert!(result.workbook.sheet_by_name(&new_name).is_some());
    assert!(result.workbook.sheet_by_name(&old_name).is_none());

    // The cross-sheet formula should reference the sanitized name.
    let sheet = result.workbook.sheet_by_name("Ref").expect("Ref missing");
    let formula = sheet
        .formula(CellRef::from_a1("A1").unwrap())
        .expect("expected formula in Ref!A1");
    let expected = format!("{new_name}!A1");
    assert_eq!(formula, expected.as_str());
    assert_parseable(formula);
}
