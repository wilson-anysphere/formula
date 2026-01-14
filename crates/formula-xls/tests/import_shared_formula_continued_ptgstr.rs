use std::io::Write;

use formula_model::CellRef;

mod common;

use common::{assert_parseable_formula, xls_fixture_builder};

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

fn assert_sharedstr_b2_formula(bytes: &[u8]) {
    let result = import_fixture(bytes);
    let sheet = result
        .workbook
        .sheet_by_name("SharedStr")
        .expect("SharedStr sheet missing");

    let b2 = CellRef::from_a1("B2").unwrap();
    let formula = sheet.formula(b2).expect("expected formula in SharedStr!B2");

    assert!(
        !formula.contains('\0'),
        "expected formula to contain no embedded NUL bytes, got {formula:?}"
    );
    assert_eq!(formula, "\"ABCDE\"");
    assert_parseable_formula(formula);
}

#[test]
fn imports_shared_formula_with_ptgstr_split_across_continue() {
    let bytes = xls_fixture_builder::build_shared_formula_continued_ptgstr_fixture_xls();
    assert_sharedstr_b2_formula(&bytes);
}

#[test]
fn imports_shared_formula_when_shrfmla_ptgstr_is_split_across_continue() {
    let bytes = xls_fixture_builder::build_shared_formula_shrfmla_continued_ptgstr_fixture_xls();
    assert_sharedstr_b2_formula(&bytes);
}
