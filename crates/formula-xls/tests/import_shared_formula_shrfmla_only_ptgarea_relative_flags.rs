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
fn imports_shrfmla_only_shared_formula_with_ptgarea_relative_flags() {
    let bytes =
        xls_fixture_builder::build_shared_formula_shrfmla_only_ptgarea_relative_flags_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("SharedOnlyAreaFlags")
        .expect("SharedOnlyAreaFlags missing");

    let b1 = CellRef::from_a1("B1").unwrap();
    let b2 = CellRef::from_a1("B2").unwrap();

    assert_eq!(sheet.formula(b1), Some("SUM(A1:A2)"));
    assert_eq!(sheet.formula(b2), Some("SUM(A2:A3)"));

    assert_parseable_formula(sheet.formula(b1).unwrap());
    assert_parseable_formula(sheet.formula(b2).unwrap());
}

