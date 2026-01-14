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
fn recovers_ptgexp_shared_formula_with_ptgarray_using_shrfmla_rgcb() {
    let bytes = xls_fixture_builder::build_shared_formula_ptgexp_ptgarray_warning_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("ShrfmlaArrayWarn")
        .expect("ShrfmlaArrayWarn missing");

    let b2 = CellRef::from_a1("B2").unwrap();
    let formula = sheet.formula(b2).expect("expected formula in B2");

    assert!(
        formula.contains("{1,2;3,4}"),
        "expected array literal in B2 formula, got {formula:?}"
    );
    assert!(
        !formula.contains("#UNKNOWN!"),
        "expected formula to decode without #UNKNOWN!, got {formula:?}"
    );

    assert_parseable_formula(formula);
}

