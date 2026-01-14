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
fn imports_shared_formula_with_ptgarray_and_wide_ptgexp_payload() {
    // Regression: wide-payload PtgExp formulas are decoded via
    // `worksheet_formulas::parse_biff8_worksheet_ptgexp_formulas`. Ensure we preserve/use SHRFMLA
    // trailing `rgcb` so PtgArray constants decode to `{...}` literals rather than `#UNKNOWN!`.
    let bytes = xls_fixture_builder::build_shared_formula_ptgarray_wide_ptgexp_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("SharedArrayWide")
        .expect("SharedArrayWide missing");

    let b2 = CellRef::from_a1("B2").expect("valid ref");
    let formula = sheet.formula(b2).expect("expected formula in SharedArrayWide!B2");

    assert!(
        formula.contains("{1,2;3,4}"),
        "expected B2 formula to contain array literal, got {formula:?}"
    );
    assert!(
        !formula.contains("#UNKNOWN!"),
        "expected formula to decode without #UNKNOWN!, got {formula:?}"
    );
    assert_parseable_formula(formula);
}

