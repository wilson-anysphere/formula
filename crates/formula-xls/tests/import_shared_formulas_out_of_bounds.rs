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
fn materializes_shared_formulas_out_of_bounds_refs_as_ref_error() {
    let bytes = xls_fixture_builder::build_shared_formula_out_of_bounds_relative_refs_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("SharedOutOfBounds")
        .expect("SharedOutOfBounds sheet missing");

    // Base cell (B65535): A65536+1.
    let base = CellRef::from_a1("B65535").expect("parse base cell");
    let base_formula = sheet
        .formula(base)
        .expect("expected shared formula base cell to import with a formula");
    assert_eq!(base_formula, "A65536+1");
    assert_parseable_formula(base_formula);

    // Follower cell (B65536, last BIFF8 row): reference shifts to A65537 which is out of bounds,
    // so Excel represents it as #REF! and the importer should surface a parseable formula string.
    let follower = CellRef::from_a1("B65536").expect("parse follower cell");
    let follower_formula = sheet
        .formula(follower)
        .expect("expected shared formula follower cell to import with a formula");
    assert_eq!(follower_formula, "#REF!+1");
    assert_parseable_formula(follower_formula);
}

