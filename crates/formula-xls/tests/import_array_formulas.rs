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
fn imports_biff8_array_formulas() {
    let bytes = xls_fixture_builder::build_array_formula_fixture_xls();
    let result = import_fixture(&bytes);
    assert!(
        result.warnings.is_empty(),
        "unexpected import warnings: {:?}",
        result.warnings
    );

    let sheet = result.workbook.sheet_by_name("Array").expect("sheet");

    let b1 = sheet
        .formula(CellRef::from_a1("B1").unwrap())
        .expect("expected formula in B1");
    let b2 = sheet
        .formula(CellRef::from_a1("B2").unwrap())
        .expect("expected formula in B2");

    assert_eq!(b1, b2);
    assert_eq!(b1, "A1:A2");
    assert_parseable_formula(b1);
}
