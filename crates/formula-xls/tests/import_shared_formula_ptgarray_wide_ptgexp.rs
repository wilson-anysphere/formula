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

    let b1 = CellRef::from_a1("B1").unwrap();
    let b2 = CellRef::from_a1("B2").unwrap();

    let f1 = sheet.formula(b1).expect("B1 formula missing");
    let f2 = sheet.formula(b2).expect("B2 formula missing");

    assert!(
        f1.contains("{1,2;3,4}"),
        "expected B1 formula to contain array literal, got {f1:?}"
    );
    assert!(
        f2.contains("{1,2;3,4}"),
        "expected B2 formula to contain array literal, got {f2:?}"
    );

    assert!(
        !f1.contains("#UNKNOWN!") && !f2.contains("#UNKNOWN!"),
        "expected formulas to decode without #UNKNOWN!, B1={f1:?}, B2={f2:?}"
    );

    // Ensure shared-formula materialization occurred and `PtgRefN(col_off=-1)` was decoded relative
    // to each cell (B1 => A1, B2 => A2).
    assert!(
        f1.contains("A1"),
        "expected B1 formula to reference A1, got {f1:?}"
    );
    assert!(
        f2.contains("A2"),
        "expected B2 formula to reference A2, got {f2:?}"
    );

    assert_parseable_formula(f1);
    assert_parseable_formula(f2);
}
