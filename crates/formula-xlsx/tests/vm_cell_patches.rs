use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{CellPatch, PackageCellPatch, WorkbookCellPatches, XlsxPackage};
use zip::ZipArchive;

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

    let cursor = Cursor::new(Vec::new());
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

fn sheet1_xml(zip_bytes: &[u8]) -> String {
    let mut archive = ZipArchive::new(Cursor::new(zip_bytes)).expect("open zip");
    let mut file = archive.by_name("xl/worksheets/sheet1.xml").expect("sheet exists");
    let mut sheet_xml = String::new();
    file.read_to_string(&mut sheet_xml)
        .expect("read sheet xml");
    sheet_xml
}

fn cell_attr(sheet_xml: &str, cell_ref: &str, attr: &str) -> Option<String> {
    let doc = roxmltree::Document::parse(sheet_xml).expect("parse worksheet xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some(cell_ref))
        .unwrap_or_else(|| panic!("expected {cell_ref} cell"));
    cell.attribute(attr).map(|s| s.to_string())
}

#[test]
fn in_memory_cell_patches_support_vm_overrides() -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" vm="1" cm="7"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;
    let bytes = build_minimal_xlsx(worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    // Insert a new cell with vm/cm.
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("B1")?,
        CellPatch::set_value(CellValue::Number(2.0))
            .with_vm(1)
            .with_cm(9),
    );
    pkg.apply_cell_patches(&patches)?;
    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap())?;
    assert_eq!(cell_attr(out_xml, "B1", "vm"), Some("1".to_string()));
    assert_eq!(cell_attr(out_xml, "B1", "cm"), Some("9".to_string()));

    // Update an existing cell and change vm without touching cm.
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(1.0)).with_vm(2),
    );
    pkg.apply_cell_patches(&patches)?;
    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap())?;
    assert_eq!(cell_attr(out_xml, "A1", "vm"), Some("2".to_string()));
    assert_eq!(cell_attr(out_xml, "A1", "cm"), Some("7".to_string()));

    // Clear vm and ensure it is removed (while preserving cm).
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(1.0)).clear_vm(),
    );
    pkg.apply_cell_patches(&patches)?;
    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap())?;
    assert_eq!(cell_attr(out_xml, "A1", "vm"), None);
    assert_eq!(cell_attr(out_xml, "A1", "cm"), Some("7".to_string()));

    Ok(())
}

#[test]
fn streaming_cell_patches_support_vm_overrides() -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" vm="1" cm="7"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;
    let bytes = build_minimal_xlsx(worksheet_xml);
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    // Insert a new cell with vm/cm.
    let insert = PackageCellPatch::for_sheet_name(
        "Sheet1",
        CellRef::from_a1("B1")?,
        CellValue::Number(2.0),
        None,
    )
    .set_vm(1)
    .set_cm(9);
    let out1 = pkg.apply_cell_patches_to_bytes(&[insert])?;
    let out_xml = sheet1_xml(&out1);
    assert_eq!(cell_attr(&out_xml, "B1", "vm"), Some("1".to_string()));
    assert_eq!(cell_attr(&out_xml, "B1", "cm"), Some("9".to_string()));

    // Update an existing cell and change vm without touching cm.
    let pkg2 = XlsxPackage::from_bytes(&out1)?;
    let update = PackageCellPatch::for_sheet_name(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellValue::Number(1.0),
        None,
    )
    .set_vm(2);
    let out2 = pkg2.apply_cell_patches_to_bytes(&[update])?;
    let out_xml = sheet1_xml(&out2);
    assert_eq!(cell_attr(&out_xml, "A1", "vm"), Some("2".to_string()));
    assert_eq!(cell_attr(&out_xml, "A1", "cm"), Some("7".to_string()));

    // Clear vm and ensure it is removed (while preserving cm).
    let pkg3 = XlsxPackage::from_bytes(&out2)?;
    let clear = PackageCellPatch::for_sheet_name(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellValue::Number(1.0),
        None,
    )
    .clear_vm();
    let out3 = pkg3.apply_cell_patches_to_bytes(&[clear])?;
    let out_xml = sheet1_xml(&out3);
    assert_eq!(cell_attr(&out_xml, "A1", "vm"), None);
    assert_eq!(cell_attr(&out_xml, "A1", "cm"), Some("7".to_string()));

    Ok(())
}

