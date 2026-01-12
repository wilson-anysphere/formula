use formula_model::{CellRef, CellValue, Workbook};
use formula_xlsx::{read_workbook_model_from_bytes, XlsxDocument};

#[test]
fn protection_roundtrips_in_new_xlsx() {
    let mut workbook = Workbook::new();
    workbook.workbook_protection.lock_structure = true;
    workbook.workbook_protection.lock_windows = true;
    workbook.workbook_protection.password_hash = Some(0x83AF);

    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");
    {
        let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
        sheet.set_value(CellRef::new(0, 0), CellValue::Number(42.0));

        sheet.sheet_protection.enabled = true;
        sheet.sheet_protection.password_hash = Some(0xCBEB);
        sheet.sheet_protection.select_locked_cells = false;
        sheet.sheet_protection.format_cells = true;
        sheet.sheet_protection.insert_rows = true;
        sheet.sheet_protection.edit_objects = true;
        sheet.sheet_protection.edit_scenarios = true;
    }

    let doc = XlsxDocument::new(workbook);
    let expected_workbook_protection = doc.workbook.workbook_protection.clone();
    let expected_sheet_protection = doc
        .workbook
        .sheets
        .first()
        .expect("sheet exists")
        .sheet_protection
        .clone();

    let bytes = doc.save_to_vec().expect("write xlsx bytes");
    let roundtrip = read_workbook_model_from_bytes(&bytes).expect("read workbook model");

    assert_eq!(roundtrip.workbook_protection, expected_workbook_protection);
    let sheet = roundtrip.sheets.first().expect("sheet exists");
    assert_eq!(sheet.sheet_protection, expected_sheet_protection);
}

