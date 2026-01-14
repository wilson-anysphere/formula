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
fn imports_shared_formulas_with_ptgfuncvar_via_shrfmla_ptgexp() {
    let bytes = xls_fixture_builder::build_shared_formula_ptgfuncvar_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("SharedFormula")
        .expect("SharedFormula sheet missing");

    let b1 = sheet
        .formula(CellRef::from_a1("B1").unwrap())
        .expect("expected formula in B1");
    assert_eq!(b1, "SUM(A1,1)");
    assert_parseable_formula(b1);

    let b2 = sheet
        .formula(CellRef::from_a1("B2").unwrap())
        .expect("expected formula in B2");
    assert_eq!(b2, "SUM(A2,1)");
    assert_parseable_formula(b2);
}

