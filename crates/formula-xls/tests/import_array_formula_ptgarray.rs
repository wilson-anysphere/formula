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
fn imports_array_formula_with_ptgarray_constant_and_rgcb() {
    let bytes = xls_fixture_builder::build_array_formula_ptgarray_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("ArrayConst")
        .expect("ArrayConst sheet missing");

    let b1 = sheet
        .formula(CellRef::from_a1("B1").unwrap())
        .expect("expected formula in ArrayConst!B1");
    let b2 = sheet
        .formula(CellRef::from_a1("B2").unwrap())
        .expect("expected formula in ArrayConst!B2");

    assert_eq!(b1, b2, "expected array formula text to match for all group cells");
    assert!(
        b1.contains("{1,2;3,4}"),
        "expected formula to contain array literal, got {b1:?}"
    );
    assert!(
        !b1.contains("#UNKNOWN!"),
        "expected formula to decode without #UNKNOWN!, got {b1:?}"
    );
    assert_parseable_formula(b1);
}

