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
fn imports_shared_formula_ptgref_relative_flags() {
    let bytes = xls_fixture_builder::build_shared_formula_ptgref_relative_flags_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 missing");

    let b2 = CellRef::from_a1("B2").unwrap();
    let formula = sheet.formula(b2).expect("expected formula in Sheet1!B2");
    assert_eq!(formula, "A2+1");
    assert_parseable_formula(formula);
}

