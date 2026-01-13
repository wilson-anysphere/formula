use std::io::Write;

mod common;

use common::{assert_parseable_formula, xls_fixture_builder};

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_biff8_shared_formulas_via_shrfmla() {
    let bytes = xls_fixture_builder::build_shared_formula_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");

    assert_eq!(sheet.formula_a1("B1").unwrap(), Some("A1+1"));
    assert_eq!(sheet.formula_a1("B2").unwrap(), Some("A2+1"));

    assert_parseable_formula(sheet.formula_a1("B1").unwrap().unwrap());
    assert_parseable_formula(sheet.formula_a1("B2").unwrap().unwrap());
}

