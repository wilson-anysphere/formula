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
fn imports_biff8_array_constant_formulas_as_literals() {
    let bytes = xls_fixture_builder::build_formula_array_constant_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("ArrayConst")
        .expect("ArrayConst missing");
    let formula = sheet
        .formula(CellRef::from_a1("A1").unwrap())
        .expect("expected formula in ArrayConst!A1");

    assert_eq!(formula, "SUM({1,2;3,4})");
    assert_parseable_formula(formula);
}

#[test]
fn imports_biff8_array_constant_string_literals_across_continue() {
    let bytes = xls_fixture_builder::build_formula_array_constant_continued_rgcb_string_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("ArrayConstStr")
        .expect("ArrayConstStr missing");
    let formula = sheet
        .formula(CellRef::from_a1("A1").unwrap())
        .expect("expected formula in ArrayConstStr!A1");

    assert_eq!(formula, "SUM({\"ABCDE\"})");
    assert_parseable_formula(formula);
}
