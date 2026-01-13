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
fn imports_shared_formulas_in_2d_ranges_with_correct_base_cell() {
    let bytes = xls_fixture_builder::build_shared_formula_2d_fixture_xls();

    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Shared")
        .expect("Shared sheet missing");

    let b1 = sheet
        .formula(CellRef::from_a1("B1").unwrap())
        .expect("expected formula in Shared!B1");
    assert_eq!(b1, "A1+1");
    assert_parseable_formula(b1);

    let c1 = sheet
        .formula(CellRef::from_a1("C1").unwrap())
        .expect("expected formula in Shared!C1");
    assert_eq!(c1, "B1+1");
    assert_parseable_formula(c1);

    let b2 = sheet
        .formula(CellRef::from_a1("B2").unwrap())
        .expect("expected formula in Shared!B2");
    assert_eq!(b2, "A2+1");
    assert_parseable_formula(b2);

    let c2 = sheet
        .formula(CellRef::from_a1("C2").unwrap())
        .expect("expected formula in Shared!C2");
    assert_eq!(c2, "B2+1");
    assert_parseable_formula(c2);
}
