use std::io::Cursor;

use formula_model::drawings::ImageId;
use formula_model::{CellRef, CellValue, EntityValue, ImageValue, RecordValue, Workbook};

#[test]
fn export_degrades_entity_and_record_values_to_plain_strings() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").unwrap();
    let sheet = workbook.sheet_mut(sheet_id).unwrap();

    sheet.set_value(
        CellRef::new(0, 0),
        CellValue::Entity(EntityValue::new("Entity Display")),
    );
    // Record values degrade to their Display implementation when exporting.
    // `formula_model::RecordValue` uses `display_field` (when present) before falling back to
    // `display_value`, so keep the test aligned with that behavior.
    sheet.set_value(
        CellRef::new(1, 0),
        CellValue::Record(
            RecordValue::new("Record Fallback")
                .with_field("Name", "Record Display")
                .with_display_field("Name"),
        ),
    );
    sheet.set_value(
        CellRef::new(2, 0),
        CellValue::Record(
            RecordValue::new("Record Fallback")
                .with_field(
                    "Logo",
                    CellValue::Image(ImageValue {
                        image_id: ImageId::new("logo.png"),
                        alt_text: Some("Logo".to_string()),
                        width: None,
                        height: None,
                    }),
                )
                .with_display_field("Logo"),
        ),
    );
    sheet.set_value(
        CellRef::new(3, 0),
        CellValue::Image(ImageValue {
            image_id: ImageId::new("logo.png"),
            alt_text: Some("Logo".to_string()),
            width: None,
            height: None,
        }),
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
    assert_eq!(
        sheet.value(CellRef::new(2, 0)),
        CellValue::String("Logo".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::new(3, 0)),
        CellValue::String("Logo".to_string())
    );
}
