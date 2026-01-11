use std::path::Path;

use formula_model::CellValue;

#[test]
fn read_workbook_handles_date_t_cells() -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/date-type.xlsx");

    let wb = formula_xlsx::read_workbook(&fixture)?;
    let sheet = &wb.sheets[0];
    let value = sheet.value_a1("C1")?;
    assert_eq!(value, CellValue::String("2024-01-01T00:00:00Z".to_string()));

    Ok(())
}

