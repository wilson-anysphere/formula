use formula_model::Workbook;
use formula_xlsx::{load_from_bytes, XlsxDocument};

#[test]
fn new_document_preserves_workbook_date_system_1904() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    workbook.date_system = formula_model::DateSystem::Excel1904;
    workbook.add_sheet("Sheet1")?;

    let doc = XlsxDocument::new(workbook);
    let bytes = doc.save_to_vec()?;

    let loaded = load_from_bytes(&bytes)?;
    assert_eq!(loaded.workbook.date_system, formula_model::DateSystem::Excel1904);

    Ok(())
}

