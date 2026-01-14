use std::io::Cursor;

use formula_model::{Cell, CellRef, CellValue, Workbook};

#[test]
fn simple_writer_emits_inline_string_phonetic_runs() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");

    let mut cell = Cell::new(CellValue::String("Base".to_string()));
    cell.phonetic = Some("PHO".to_string());
    sheet.set_cell(CellRef::from_a1("A1")?, cell);

    let mut cursor = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer_with_kind(
        &workbook,
        &mut cursor,
        formula_xlsx::WorkbookKind::Workbook,
    )?;
    let bytes = cursor.into_inner();

    let doc = formula_xlsx::load_from_bytes(&bytes)?;
    let sheet = doc
        .workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 exists");
    let cell = sheet.cell(CellRef::from_a1("A1")?).expect("A1 cell exists");

    assert_eq!(cell.value, CellValue::String("Base".to_string()));
    assert_eq!(cell.phonetic.as_deref(), Some("PHO"));

    Ok(())
}

