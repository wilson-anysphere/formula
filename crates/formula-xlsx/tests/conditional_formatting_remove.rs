use std::path::Path;

use formula_xlsx::{load_from_bytes, XlsxPackage};

fn assert_clearing_conditional_formatting_removes_worksheet_blocks(
    fixture_name: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    const X14_CF_URI: &str = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";

    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join(fixture_name);
    let bytes = std::fs::read(&fixture_path)?;

    // Sanity check: the fixture should actually have conditional formatting.
    let fixture_pkg = XlsxPackage::from_bytes(&bytes)?;
    let fixture_sheet_xml = std::str::from_utf8(
        fixture_pkg
            .part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml exists"),
    )?;
    assert!(
        fixture_sheet_xml.contains("conditionalFormatting"),
        "fixture {fixture_name} should contain conditionalFormatting blocks, got: {fixture_sheet_xml}"
    );

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
    assert!(
        !sheet_xml.contains(X14_CF_URI),
        "expected sheet1.xml to remove x14 conditional formatting <ext uri=...> block (uri={X14_CF_URI}), got: {sheet_xml}"
    );

    Ok(())
}

#[test]
fn clearing_conditional_formatting_removes_worksheet_blocks_2007(
) -> Result<(), Box<dyn std::error::Error>> {
    assert_clearing_conditional_formatting_removes_worksheet_blocks("conditional_formatting_2007.xlsx")
}

#[test]
fn clearing_conditional_formatting_removes_worksheet_blocks_x14(
) -> Result<(), Box<dyn std::error::Error>> {
    assert_clearing_conditional_formatting_removes_worksheet_blocks("conditional_formatting_x14.xlsx")
}
