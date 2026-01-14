use std::io::Write;

use formula_model::CellRef;

mod common;

use common::{assert_parseable_formula, xls_fixture_builder};

#[test]
fn imports_shared_formula_ptgmemarean_with_nested_ptgarray_advances_rgcb() {
    let bytes = xls_fixture_builder::build_shared_formula_ptgmemarean_rgcb_fixture_xls();

    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&bytes).expect("write xls bytes");

    let result = formula_xls::import_xls_path(tmp.path()).expect("import xls");
    let sheet = result
        .workbook
        .sheet_by_name("SharedMemRgcb")
        .expect("SharedMemRgcb sheet missing");

    let b1 = CellRef::from_a1("B1").unwrap();
    let b2 = CellRef::from_a1("B2").unwrap();

    assert_eq!(sheet.formula(b1), Some("A1+{5,6;7,8}"));
    assert_eq!(sheet.formula(b2), Some("A2+{5,6;7,8}"));

    assert_parseable_formula(sheet.formula(b1).unwrap());
    assert_parseable_formula(sheet.formula(b2).unwrap());
}

