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
fn imports_shared_formula_with_wide_ptgexp_payload_in_high_columns() {
    // Regression: `parse_biff8_worksheet_ptgexp_formulas` must accept base-cell columns > 255 when
    // recovering wide-payload PtgExp shared formulas (Ref8 SHRFMLA ranges).
    let bytes = xls_fixture_builder::build_shared_formula_ptgexp_wide_payload_high_col_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("SharedWideHighCol")
        .expect("SharedWideHighCol missing");

    let xfc1 = CellRef::from_a1("XFC1").unwrap();
    let xfd1 = CellRef::from_a1("XFD1").unwrap();

    assert_eq!(sheet.formula(xfc1), Some("XFD1+1"));
    assert_eq!(sheet.formula(xfd1), Some("#REF!+1"));

    assert_parseable_formula(sheet.formula(xfc1).unwrap());
    assert_parseable_formula(sheet.formula(xfd1).unwrap());
}

