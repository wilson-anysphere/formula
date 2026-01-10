use std::path::PathBuf;

use formula_model::{CellValue, SheetVisibility};

#[test]
fn imports_sheet_visibility() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("hidden.xls");

    let result = formula_xls::import_xls_path(&fixture_path).expect("import xls");

    let visible = result
        .workbook
        .sheet_by_name("Visible")
        .expect("Visible sheet missing");
    assert_eq!(visible.visibility, SheetVisibility::Visible);
    assert_eq!(
        visible.value_a1("A1").unwrap(),
        CellValue::String("Visible sheet".to_owned())
    );

    let hidden = result
        .workbook
        .sheet_by_name("Hidden")
        .expect("Hidden sheet missing");
    assert_eq!(hidden.visibility, SheetVisibility::Hidden);
    assert_eq!(
        hidden.value_a1("A1").unwrap(),
        CellValue::String("Hidden sheet".to_owned())
    );
}
