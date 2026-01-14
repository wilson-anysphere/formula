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
fn falls_back_to_biff8_formulas_when_calamine_formula_decode_fails() {
    let bytes = xls_fixture_builder::build_calamine_formula_error_biff_fallback_fixture_xls();

    // Regression guard: this fixture must force calamine's worksheet_formula() to fail (either by
    // returning an error/panic or by dropping the formula), otherwise the importer won't exercise
    // the BIFF fallback path.
    let mut wb: Xls<_> = Xls::new(Cursor::new(bytes.clone())).expect("open xls via calamine");
    let calamine_result = catch_unwind(AssertUnwindSafe(|| wb.worksheet_formula("Sheet1")));
    match calamine_result {
        Ok(Ok(range)) => assert!(
            range.used_cells().next().is_none(),
            "expected calamine worksheet_formula() to return no formulas for this fixture"
        ),
        Ok(Err(_)) | Err(_) => {}
    }

    let result = import_fixture(&bytes);
    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");

    let cell = CellRef::from_a1("B1").unwrap();
    let formula = sheet.formula(cell).expect("expected formula in Sheet1!B1");
    assert_eq!(formula, "#UNKNOWN!");
    assert_parseable_formula(formula);
}
