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
fn imports_shrfmla_only_shared_formula_ptgref_col_oob_as_ref_error() {
    let bytes =
        xls_fixture_builder::build_shared_formula_ptgref_col_oob_shrfmla_only_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("SharedRefColOOB_ShrFmlaOnly")
        .expect("SharedRefColOOB_ShrFmlaOnly missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let b1 = CellRef::from_a1("B1").unwrap();

    assert_eq!(sheet.formula(a1), Some("XFD1+1"));
    assert_eq!(sheet.formula(b1), Some("#REF!+1"));

    assert_parseable_formula(sheet.formula(a1).unwrap());
    assert_parseable_formula(sheet.formula(b1).unwrap());
}

