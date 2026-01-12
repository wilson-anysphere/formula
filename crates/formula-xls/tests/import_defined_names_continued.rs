use std::io::Write;

use formula_model::{CellRef, DefinedNameScope};

mod common;

use common::{assert_parseable_formula, xls_fixture_builder};

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_defined_names_split_across_continue_records() {
    let bytes = xls_fixture_builder::build_continued_name_record_fixture_xls();
    let result = import_fixture(&bytes);

    let name = result
        .workbook
        .get_defined_name(DefinedNameScope::Workbook, "MyContinuedName")
        .expect("expected defined name to be imported");

    assert_eq!(name.refers_to, "DefinedNames!$A$1");
    assert_parseable_formula(&name.refers_to);
    assert_eq!(
        name.comment.as_deref(),
        Some("This is a long description used to test continued NAME records.")
    );

    // Ensure worksheet formulas that reference the defined name decode correctly (calamine needs
    // the NAME table for `PtgName` tokens).
    let sheet = result
        .workbook
        .sheet_by_name("DefinedNames")
        .expect("expected sheet to be present");
    let formula = sheet
        .formula(CellRef::from_a1("A1").unwrap())
        .expect("expected formula in DefinedNames!A1");
    assert_eq!(formula, "MyContinuedName");
    assert_parseable_formula(formula);
}
