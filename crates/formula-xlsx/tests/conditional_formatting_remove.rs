use std::path::Path;

use formula_xlsx::{load_from_bytes, XlsxPackage};

#[test]
fn clearing_conditional_formatting_removes_worksheet_blocks() -> Result<(), Box<dyn std::error::Error>>
{
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("conditional_formatting_2007.xlsx");
    let bytes = std::fs::read(&fixture_path)?;

    let mut doc = load_from_bytes(&bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .clear_conditional_formatting();

    let saved = doc.save_to_vec()?;
    let reopened = XlsxPackage::from_bytes(&saved)?;
    let sheet_xml = std::str::from_utf8(
        reopened
            .part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml exists"),
    )?;

    assert!(
        !sheet_xml.contains("conditionalFormatting"),
        "expected sheet1.xml to remove conditionalFormatting blocks, got: {sheet_xml}"
    );
    assert!(
        !sheet_xml.contains("conditionalFormattings"),
        "expected sheet1.xml to remove x14 conditionalFormattings blocks, got: {sheet_xml}"
    );

    Ok(())
}

