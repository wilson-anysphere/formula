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
fn array_formulas_with_external_refs_are_decoded_with_fidelity() {
    let bytes = xls_fixture_builder::build_array_formula_external_refs_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("ArrayExt")
        .expect("ArrayExt sheet missing");

    let b1 = sheet
        .formula(CellRef::from_a1("B1").unwrap())
        .expect("expected formula in ArrayExt!B1");
    let b2 = sheet
        .formula(CellRef::from_a1("B2").unwrap())
        .expect("expected formula in ArrayExt!B2");
    assert_eq!(b1, "'[Book1.xlsx]ExtSheet'!$A$1+1");
    assert_eq!(b2, b1);
    assert_parseable_formula(b1);

    let c1 = sheet
        .formula(CellRef::from_a1("C1").unwrap())
        .expect("expected formula in ArrayExt!C1");
    let c2 = sheet
        .formula(CellRef::from_a1("C2").unwrap())
        .expect("expected formula in ArrayExt!C2");
    assert_eq!(c1, "'[Book1.xlsx]ExtSheet'!ExtDefined+1");
    assert_eq!(c2, c1);
    assert_parseable_formula(c1);
}

