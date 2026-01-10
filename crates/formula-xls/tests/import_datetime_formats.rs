use std::path::PathBuf;

use formula_model::{CellRef, CellValue};

#[test]
fn applies_default_date_number_formats() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("dates.xls");

    let result = formula_xls::import_xls_path(&fixture_path).expect("import xls");
    let sheet = result
        .workbook
        .sheet_by_name("Dates")
        .expect("Dates missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let a2 = CellRef::from_a1("A2").unwrap();
    let a3 = CellRef::from_a1("A3").unwrap();
    let a4 = CellRef::from_a1("A4").unwrap();

    let c1 = sheet.cell(a1).expect("A1 missing");
    let c2 = sheet.cell(a2).expect("A2 missing");
    let c3 = sheet.cell(a3).expect("A3 missing");
    let c4 = sheet.cell(a4).expect("A4 missing");

    assert!(matches!(c1.value, CellValue::Number(_)));
    assert!(matches!(c2.value, CellValue::Number(_)));
    assert!(matches!(c3.value, CellValue::Number(_)));
    assert!(matches!(c4.value, CellValue::Number(_)));

    let fmt1 = result
        .workbook
        .styles
        .get(c1.style_id)
        .and_then(|s| s.number_format.as_deref());
    let fmt2 = result
        .workbook
        .styles
        .get(c2.style_id)
        .and_then(|s| s.number_format.as_deref());
    let fmt3 = result
        .workbook
        .styles
        .get(c3.style_id)
        .and_then(|s| s.number_format.as_deref());
    let fmt4 = result
        .workbook
        .styles
        .get(c4.style_id)
        .and_then(|s| s.number_format.as_deref());

    assert_eq!(fmt1, Some("m/d/yy"));
    assert_eq!(fmt2, Some("m/d/yy h:mm:ss"));
    assert_eq!(fmt3, Some("h:mm:ss"));
    assert_eq!(fmt4, Some("[h]:mm:ss"));
}
