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
fn imports_shrfmla_only_shared_formula_with_ptgarray_and_rgcb() {
    let bytes = xls_fixture_builder::build_shared_formula_shrfmla_only_ptgarray_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("ShrfmlaArray")
        .expect("ShrfmlaArray missing");

    let b1 = CellRef::from_a1("B1").unwrap();
    let b2 = CellRef::from_a1("B2").unwrap();

    let f1 = sheet.formula(b1).expect("expected formula in ShrfmlaArray!B1");
    let f2 = sheet.formula(b2).expect("expected formula in ShrfmlaArray!B2");

    assert!(
        f1.contains("{1,2;3,4}"),
        "expected B1 formula to contain array literal, got {f1:?}"
    );
    assert!(
        f2.contains("{1,2;3,4}"),
        "expected B2 formula to contain array literal, got {f2:?}"
    );

    // Ensure shared-formula expansion occurred (relative ref A1 -> A2).
    assert!(
        f1.contains("A1"),
        "expected B1 formula to reference A1, got {f1:?}"
    );
    assert!(
        f2.contains("A2"),
        "expected B2 formula to reference A2, got {f2:?}"
    );
    assert_ne!(
        f1, f2,
        "expected shared formula to materialize per-cell references (A1 vs A2)"
    );

    assert!(
        !f1.contains("#UNKNOWN!") && !f2.contains("#UNKNOWN!"),
        "expected formulas to decode without #UNKNOWN!, B1={f1:?}, B2={f2:?}"
    );

    assert_parseable_formula(f1);
    assert_parseable_formula(f2);
}

