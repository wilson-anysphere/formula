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
fn imports_biff8_table_formulas_from_wide_ptgtbl_payload() {
    let bytes = xls_fixture_builder::build_table_formula_ptgtbl_wide_payload_fixture_xls();
    let result = import_fixture(&bytes);
    assert!(result.warnings.is_empty(), "warnings={:?}", result.warnings);
    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1");

    // The fixture stores a FORMULA record at D21 whose rgce begins with `PtgTbl` but uses a
    // non-canonical payload width (row u32 + col u16). The corresponding TABLE record describes
    // input cells A1 and B2.
    let cell = CellRef::from_a1("D21").unwrap();
    assert_eq!(sheet.formula(cell), Some("TABLE(A1,B2)"));
    assert_parseable_formula(sheet.formula(cell).unwrap());
}
