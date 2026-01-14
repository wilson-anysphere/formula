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
fn materializes_shared_formula_ptgarea_row_oob_as_ref_error() {
    let bytes = xls_fixture_builder::build_shared_formula_ptgarea_row_oob_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("SharedArea")
        .expect("SharedArea missing");

    let follower = CellRef::from_a1("B65536").expect("B65536 ref");
    let formula = sheet
        .formula(follower)
        .expect("expected formula in SharedArea!B65536");
    assert_eq!(formula, "SUM(#REF!)+1");
    assert_parseable_formula(formula);
}

