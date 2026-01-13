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
fn rewrites_shared_formula_3d_refs_to_sanitized_sheet_names() {
    let bytes = xls_fixture_builder::build_shared_formula_sheet_name_sanitization_fixture_xls();
    let result = import_fixture(&bytes);

    assert!(
        result.workbook.sheet_by_name("Bad:Name").is_none(),
        "expected invalid sheet name to be sanitized"
    );
    assert!(
        result.workbook.sheet_by_name("Bad_Name").is_some(),
        "expected sanitized sheet to be present"
    );

    let sheet = result.workbook.sheet_by_name("Ref").expect("Ref missing");
    let formula = sheet
        .formula(CellRef::from_a1("A2").unwrap())
        .expect("expected formula in Ref!A2");

    assert!(
        !formula.contains("Bad:Name"),
        "expected formula to no longer reference original sheet name, got {formula:?}"
    );
    assert!(
        formula.contains("Bad_Name"),
        "expected formula to reference sanitized sheet name, got {formula:?}"
    );
    assert_parseable_formula(formula);
}

