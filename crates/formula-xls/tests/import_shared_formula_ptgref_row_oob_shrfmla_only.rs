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
fn imports_shrfmla_only_shared_formula_ptgref_row_oob_as_ref_error() {
    // Fixture uses only a SHRFMLA record (no FORMULA/PtgExp cells) for the shared range, so the
    // importer must expand/materialize the shared rgce itself.
    let bytes = xls_fixture_builder::build_shared_formula_ptgref_row_oob_shrfmla_only_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("SharedRefRowOOB_ShrFmlaOnly")
        .expect("SharedRefRowOOB_ShrFmlaOnly missing");

    let base = sheet
        .formula(CellRef::from_a1("B65535").unwrap())
        .expect("expected formula in SharedRefRowOOB_ShrFmlaOnly!B65535");
    assert_eq!(base, "A65536+1");
    assert_parseable_formula(base);

    let follower = sheet
        .formula(CellRef::from_a1("B65536").unwrap())
        .expect("expected formula in SharedRefRowOOB_ShrFmlaOnly!B65536");
    assert_eq!(follower, "#REF!+1");
    assert_parseable_formula(follower);
}

