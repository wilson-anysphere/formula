use std::io::Cursor;

use formula_model::{CellRef, CellValue, EntityValue, RecordValue, Workbook};

#[test]
fn export_degrades_entity_and_record_values_to_plain_strings() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").unwrap();
    let sheet = workbook.sheet_mut(sheet_id).unwrap();

    sheet.set_value(
        CellRef::new(0, 0),
        CellValue::Entity(EntityValue::new("Entity Display")),
    );
    sheet.set_value(
        CellRef::new(1, 0),
        CellValue::Record(RecordValue::new("Record Display")),
    );

    let mut cursor = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut cursor).unwrap();
    let bytes = cursor.into_inner();

    let roundtrip = formula_xlsx::read_workbook_model_from_bytes(&bytes).unwrap();
    let sheet = roundtrip.sheets.first().expect("sheet present");

    assert_eq!(
        sheet.value(CellRef::new(0, 0)),
        CellValue::String("Entity Display".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::new(1, 0)),
        CellValue::String("Record Display".to_string())
    );
}

