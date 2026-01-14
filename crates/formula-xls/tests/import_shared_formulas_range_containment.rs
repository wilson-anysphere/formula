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
fn recovers_shared_formulas_when_ptgexp_master_is_not_range_start() {
    let bytes = xls_fixture_builder::build_shared_formula_master_not_top_left_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Shared")
        .expect("Shared missing");

    let b1 = sheet.formula(CellRef::from_a1("B1").unwrap());
    let b2 = sheet.formula(CellRef::from_a1("B2").unwrap());

    assert_eq!(b1, Some("A1"));
    assert_eq!(b2, Some("A2"));

    assert_parseable_formula(b1.unwrap());
    assert_parseable_formula(b2.unwrap());
}
