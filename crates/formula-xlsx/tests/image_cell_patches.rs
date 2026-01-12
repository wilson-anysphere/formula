use std::io::Write as _;

use formula_model::drawings::ImageId;
use formula_model::{CellRef, CellValue, ImageValue};
use formula_xlsx::{CellPatch, PackageCellPatch, WorkbookCellPatches, XlsxPackage};

fn build_minimal_xlsx(sheet_xml: &str) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let cursor = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn assert_cell_inline_str_text(sheet_xml: &str, cell_ref: &str, expected_text: &str) {
    let doc = roxmltree::Document::parse(sheet_xml).expect("parse worksheet xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some(cell_ref))
        .unwrap_or_else(|| panic!("expected {cell_ref} cell"));
    assert_eq!(
        cell.attribute("t"),
        Some("inlineStr"),
        "expected {cell_ref} to be written as inlineStr (no sharedStrings.xml in fixture), got: {sheet_xml}"
    );
    let t = cell
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "t")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(
        t, expected_text,
        "expected {cell_ref} to contain inline string text {expected_text}, got: {sheet_xml}"
    );
}

fn image_value_with_alt_text(alt_text: &str) -> CellValue {
    CellValue::Image(ImageValue {
        image_id: ImageId::new("image1.png"),
        alt_text: Some(alt_text.to_string()),
        width: None,
        height: None,
    })
}

#[test]
fn apply_cell_patches_writes_image_alt_text_fallback() -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"/>
  </sheetData>
</worksheet>"#;
    let bytes = build_minimal_xlsx(worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(image_value_with_alt_text("ExampleAltText")),
    );
    pkg.apply_cell_patches(&patches)?;

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap())?;
    assert_cell_inline_str_text(out_xml, "A1", "ExampleAltText");
    Ok(())
}

#[test]
fn apply_cell_patches_to_bytes_writes_image_alt_text_fallback(
) -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"/>
  </sheetData>
</worksheet>"#;
    let bytes = build_minimal_xlsx(worksheet_xml);
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    let patch = PackageCellPatch::for_sheet_name(
        "Sheet1",
        CellRef::from_a1("A1")?,
        image_value_with_alt_text("ExampleAltText"),
        None,
    );

    let out_bytes = pkg.apply_cell_patches_to_bytes(&[patch])?;
    let out_pkg = XlsxPackage::from_bytes(&out_bytes)?;
    let out_xml = std::str::from_utf8(out_pkg.part("xl/worksheets/sheet1.xml").unwrap())?;
    assert_cell_inline_str_text(out_xml, "A1", "ExampleAltText");
    Ok(())
}
