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
fn decodes_sheet_scoped_ptgname_using_sanitized_sheet_names() {
    let bytes =
        xls_fixture_builder::build_shared_formula_sheet_scoped_name_sanitization_fixture_xls();
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
    assert!(
        !formula.contains("#NAME?"),
        "expected PtgName to resolve to a defined name, got {formula:?}"
    );

    assert_eq!(formula, "Bad_Name!LocalName");
    assert_parseable_formula(formula);
}

#[test]
fn decodes_sheet_scoped_ptgname_when_sheet_name_contains_apostrophe() {
    let bytes =
        xls_fixture_builder::build_shared_formula_sheet_scoped_name_apostrophe_fixture_xls();
    let result = import_fixture(&bytes);

    assert!(
        result.workbook.sheet_by_name("O'Brien").is_some(),
        "expected O'Brien sheet to be present"
    );

    let sheet = result.workbook.sheet_by_name("Ref").expect("Ref missing");
    let formula = sheet
        .formula(CellRef::from_a1("A2").unwrap())
        .expect("expected formula in Ref!A2");

    assert_eq!(formula, "'O''Brien'!LocalName");
    assert_parseable_formula(formula);
}

#[test]
fn decodes_sheet_scoped_ptgname_when_sheet_name_is_true() {
    let bytes = xls_fixture_builder::build_shared_formula_sheet_scoped_name_true_sheet_fixture_xls();
    let result = import_fixture(&bytes);

    assert!(
        result.workbook.sheet_by_name("TRUE").is_some(),
        "expected TRUE sheet to be present"
    );

    let sheet = result.workbook.sheet_by_name("Ref").expect("Ref missing");
    let formula = sheet
        .formula(CellRef::from_a1("A2").unwrap())
        .expect("expected formula in Ref!A2");

    assert_eq!(formula, "'TRUE'!LocalName");
    assert_parseable_formula(formula);
}

#[test]
fn decodes_sheet_scoped_ptgname_when_sheet_name_looks_like_cell_reference() {
    let bytes = xls_fixture_builder::build_shared_formula_sheet_scoped_name_a1_sheet_fixture_xls();
    let result = import_fixture(&bytes);

    assert!(
        result.workbook.sheet_by_name("A1").is_some(),
        "expected A1 sheet to be present"
    );

    let sheet = result.workbook.sheet_by_name("Ref").expect("Ref missing");
    let formula = sheet
        .formula(CellRef::from_a1("A2").unwrap())
        .expect("expected formula in Ref!A2");

    assert_eq!(formula, "'A1'!LocalName");
    assert_parseable_formula(formula);
}
