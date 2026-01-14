use std::io::{Cursor, Write};
use std::panic::{catch_unwind, AssertUnwindSafe};

use calamine::{Reader, Xls};
use formula_model::CellRef;

mod common;

use common::{assert_parseable_formula, xls_fixture_builder};

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_shared_formula_bodies_that_reference_defined_names_via_ptgname() {
    let bytes = xls_fixture_builder::build_shared_formula_ptgname_fixture_xls();

    // Regression guard: ensure calamine does not surface the follower-cell formula. The importer
    // should recover it from `SHRFMLA` + `PtgExp` records in the BIFF workbook stream.
    let mut wb: Xls<_> = Xls::new(Cursor::new(bytes.clone())).expect("open xls via calamine");
    let calamine_result = catch_unwind(AssertUnwindSafe(|| wb.worksheet_formula("SharedName")));
    if let Ok(Ok(range)) = calamine_result {
        let start = range.start().unwrap_or((0, 0));
        let b2_present = range.used_cells().any(|(row, col, _)| {
            start.0.saturating_add(row as u32) == 1 && start.1.saturating_add(col as u32) == 1
        });
        assert!(
            !b2_present,
            "expected calamine worksheet_formula() to omit SharedName!B2 for this fixture"
        );
    }

    let result = import_fixture(&bytes);
    let sheet = result
        .workbook
        .sheet_by_name("SharedName")
        .expect("SharedName sheet missing");

    let b2 = sheet
        .formula(CellRef::from_a1("B2").unwrap())
        .expect("expected formula in SharedName!B2");
    assert_eq!(b2, "MyName");
    assert_parseable_formula(b2);
}
