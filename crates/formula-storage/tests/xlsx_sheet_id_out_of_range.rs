use formula_storage::Storage;

#[test]
fn export_model_workbook_ignores_out_of_range_xlsx_sheet_id() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    storage
        .set_sheet_xlsx_metadata(sheet.id, Some(-1), Some("rId7"))
        .expect("set invalid xlsx sheet id");

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");
    let sheet = exported.sheet_by_name("Sheet1").expect("sheet exists");
    assert_eq!(sheet.xlsx_sheet_id, None);
    assert_eq!(sheet.xlsx_rel_id.as_deref(), Some("rId7"));
}

