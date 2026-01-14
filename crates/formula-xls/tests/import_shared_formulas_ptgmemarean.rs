use std::io::Write;

use formula_model::CellRef;

mod common;

use common::{assert_parseable_formula, xls_fixture_builder};

#[test]
fn imports_shared_formula_with_ptgmemarean_in_shrfmla() {
    let bytes = xls_fixture_builder::build_shared_formula_ptgmemarean_fixture_xls();

    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&bytes).expect("write xls bytes");

    let result = formula_xls::import_xls_path(tmp.path()).expect("import xls");
    let sheet = result
        .workbook
        .sheet_by_name("Shared")
        .expect("Shared sheet missing");

    let b1 = CellRef::from_a1("B1").unwrap();
    let b2 = CellRef::from_a1("B2").unwrap();
    assert_eq!(sheet.formula(b1), Some("A1+1"));

    let formula = sheet.formula(b2).expect("expected B2 formula");
    assert_eq!(formula, "A2+1");
    assert_parseable_formula(formula);
}
