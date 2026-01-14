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
fn imports_shrfmla_only_shared_formula_ptgrefn_col_oob_relative_offsets() {
    let bytes = xls_fixture_builder::build_shared_formula_ptgrefn_col_oob_shrfmla_only_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("SharedRefN_ColOOB_ShrFmlaOnly")
        .expect("SharedRefN_ColOOB_ShrFmlaOnly missing");

    let a1 = sheet
        .formula(CellRef::from_a1("A1").unwrap())
        .expect("expected formula in SharedRefN_ColOOB_ShrFmlaOnly!A1");
    assert_eq!(a1, "#REF!+1");
    assert_parseable_formula(a1);

    let b1 = sheet
        .formula(CellRef::from_a1("B1").unwrap())
        .expect("expected formula in SharedRefN_ColOOB_ShrFmlaOnly!B1");
    assert_eq!(b1, "A1+1");
    assert_parseable_formula(b1);
}

