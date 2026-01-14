use std::io::Write;

use formula_model::CellRef;

mod common;

use common::{assert_parseable_formula, xls_fixture_builder};

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn recovers_ptgexp_array_formula_when_array_record_missing() {
    let bytes = xls_fixture_builder::build_array_formula_ptgexp_missing_array_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("ArrayMissing")
        .expect("ArrayMissing sheet missing");

    // B2 should be recovered from the base cell's formula (B1) *without* shifting relative
    // references because array formulas are anchored at the base cell.
    let b2 = CellRef::from_a1("B2").unwrap();
    let formula = sheet
        .formula(b2)
        .expect("expected formula in ArrayMissing!B2");
    assert_eq!(formula, "A1+1");
    assert_parseable_formula(formula);
}
