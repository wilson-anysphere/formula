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
fn imports_shared_formula_ptgname() {
    let bytes = xls_fixture_builder::build_shared_formula_ptgname_fixture_xls();
    let result = import_fixture(&bytes);

    assert!(
        !result
            .warnings
            .iter()
            .any(|w| w.message.contains("PtgName references missing name index")),
        "unexpected missing-name-index warnings: {:?}",
        result.warnings
    );
    let sheet = result
        .workbook
        .sheet_by_name("SharedName")
        .expect("SharedName missing");

    let b1 = sheet
        .formula(CellRef::from_a1("B1").unwrap())
        .expect("expected formula in SharedName!B1");
    let b2 = sheet
        .formula(CellRef::from_a1("B2").unwrap())
        .expect("expected formula in SharedName!B2");

    assert!(b1.contains("MyName"), "formula={b1:?}");
    assert!(!b1.contains("#NAME?"), "formula={b1:?}");
    assert_parseable_formula(b1);

    assert!(b2.contains("MyName"), "formula={b2:?}");
    assert!(!b2.contains("#NAME?"), "formula={b2:?}");
    assert_parseable_formula(b2);
}
