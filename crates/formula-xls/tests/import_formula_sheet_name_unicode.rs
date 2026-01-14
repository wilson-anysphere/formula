use std::io::Write;

use calamine::{open_workbook, Reader, Xls};
use formula_model::CellRef;

mod common;

use common::{assert_parseable_formula, xls_fixture_builder};

#[test]
fn rewrites_cross_sheet_formulas_when_sheet_name_is_unicode() {
    let bytes = xls_fixture_builder::build_formula_sheet_name_unicode_sheet_fixture_xls();

    // Use a single temp file path so we can inspect calamine's decoded metadata and then run the
    // importer on the same bytes.
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&bytes).expect("write xls bytes");
    tmp.flush().expect("flush temp file");

    // Capture calamine's view of the sheet names; some versions have trouble decoding BIFF8
    // BoundSheet names stored in the uncompressed (UTF-16LE) form.
    let calamine: Xls<_> = open_workbook(tmp.path()).expect("open workbook via calamine");
    let sheets = calamine.sheets_metadata().to_vec();
    assert!(
        sheets.len() >= 2,
        "expected fixture to contain at least 2 sheets, got {}",
        sheets.len()
    );
    let calamine_sheet0_name = sheets[0].name.replace('\0', "");

    let result = formula_xls::import_xls_path(tmp.path()).expect("import xls");

    assert!(
        result.workbook.sheet_by_name("数据").is_some(),
        "expected unicode sheet name to be imported from BIFF BoundSheet"
    );

    let sheet = result.workbook.sheet_by_name("Ref").expect("Ref missing");
    let formula = sheet
        .formula(CellRef::from_a1("A1").unwrap())
        .expect("expected formula in Ref!A1");

    assert!(
        formula.starts_with("'数据'!"),
        "expected formula to reference the final sheet name with quotes for formula-engine, got {formula:?}"
    );

    if calamine_sheet0_name != "数据" && !calamine_sheet0_name.is_empty() {
        assert!(
            !formula.contains(&calamine_sheet0_name),
            "expected formula to be rewritten away from calamine sheet name {calamine_sheet0_name:?}, got {formula:?}"
        );
    }

    assert_parseable_formula(formula);
}
