use std::io::Cursor;

use formula_model::import::{import_csv_into_workbook, CsvOptions};
use formula_model::{CellRef, CellValue};

#[test]
fn writes_and_reads_columnar_backed_sheet_with_overlay_override() {
    let csv = "col1,col2\n1,hello\n2,world\n";

    let mut workbook = formula_model::Workbook::new();
    let sheet_id = import_csv_into_workbook(
        &mut workbook,
        "Data",
        Cursor::new(csv.as_bytes()),
        CsvOptions::default(),
    )
    .expect("import csv into workbook");

    // Override a value from the columnar backing store using a sparse overlay cell.
    {
        let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
        sheet.set_value(CellRef::new(0, 1), CellValue::String("OVERRIDE".to_string()));
    }

    let mut cursor = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut cursor).expect("write workbook");
    let bytes = cursor.into_inner();

    let roundtrip =
        formula_xlsx::read_workbook_model_from_bytes(&bytes).expect("read workbook model");
    let sheet = roundtrip.sheets.first().expect("sheet");

    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("OVERRIDE".to_string())
    );
    assert_eq!(sheet.value_a1("A2").unwrap(), CellValue::Number(2.0));
    assert_eq!(
        sheet.value_a1("B2").unwrap(),
        CellValue::String("world".to_string())
    );
}
