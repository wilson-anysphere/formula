use formula_model::{CellRef, CellValue, Workbook};
use formula_xlsx::{load_from_bytes, XlsxDocument};

#[test]
fn cell_phonetic_roundtrips_via_inline_string() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;
    let a1 = CellRef::from_a1("A1")?;

    {
        let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
        sheet.set_value(a1, CellValue::String("Base".to_string()));
        let cell = sheet.cell_mut(a1).expect("cell exists");
        cell.phonetic = Some("PHO".to_string());
    }

    let bytes = XlsxDocument::new(workbook).save_to_vec()?;
    let doc = load_from_bytes(&bytes)?;

    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).expect("sheet exists");

    assert_eq!(sheet.value(a1), CellValue::String("Base".to_string()));
    let cell = sheet.cell(a1).expect("cell exists");
    assert_eq!(cell.phonetic.as_deref(), Some("PHO"));

    Ok(())
}

