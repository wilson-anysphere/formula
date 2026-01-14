use formula_model::{Cell, CellRef, CellValue, Workbook};
use formula_storage::{ImportModelWorkbookOptions, Storage};

#[test]
fn model_workbook_codepage_and_cell_phonetic_roundtrip() {
    let storage = Storage::open_in_memory().expect("open storage");

    let mut workbook = Workbook::new();
    workbook.codepage = 932;
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");
    let sheet = workbook.sheet_mut(sheet_id).expect("get sheet");

    let mut cell = Cell::new(CellValue::String("漢字".to_string()));
    cell.phonetic = Some("PHO".to_string());
    sheet.set_cell(CellRef::new(0, 0), cell);

    let workbook_meta = storage
        .import_model_workbook(&workbook, ImportModelWorkbookOptions::new("Book"))
        .expect("import model workbook");

    let exported = storage
        .export_model_workbook(workbook_meta.id)
        .expect("export model workbook");
    assert_eq!(exported.codepage, 932);

    let exported_sheet = exported.sheet_by_name("Sheet1").expect("sheet by name");
    let exported_cell = exported_sheet
        .cell(CellRef::new(0, 0))
        .expect("cell should exist");
    assert_eq!(exported_cell.phonetic.as_deref(), Some("PHO"));
}

