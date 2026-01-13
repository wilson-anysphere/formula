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
fn shared_formula_external_3d_refs_are_decoded_with_fidelity() {
    let bytes = xls_fixture_builder::build_shared_formula_external_refs_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Shared")
        .expect("Shared sheet missing");

    // Follower cell should have a fully-qualified external workbook ref, not `#REF!`.
    let b2 = sheet
        .formula(CellRef::from_a1("B2").unwrap())
        .expect("expected formula in Shared!B2");
    assert_eq!(b2, "'[Book1.xlsx]ExtSheet'!$A$1+1");
    assert_parseable_formula(b2);

    // Second shared formula uses PtgNameX to reference an external defined name.
    let c2 = sheet
        .formula(CellRef::from_a1("C2").unwrap())
        .expect("expected formula in Shared!C2");
    assert_eq!(c2, "'[Book1.xlsx]ExtSheet'!ExtDefined+1");
    assert_parseable_formula(c2);
}
