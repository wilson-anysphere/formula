use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    patch_xlsx_streaming_workbook_cell_patches, CellPatch, WorkbookCellPatches, XlsxPackage,
};
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

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn read_sheet1_xml_from_xlsx(bytes: &[u8]) -> String {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet1.xml exists")
        .read_to_string(&mut sheet_xml)
        .expect("read sheet1.xml");
    sheet_xml
}

#[test]
fn in_memory_cell_patches_can_set_vm_cm_on_inserted_cell() -> Result<(), Box<dyn std::error::Error>>
{
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(2.0)),
    );
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("B1")?,
        CellPatch::set_value(CellValue::Number(3.0))
            .with_vm(9)
            .with_cm(7),
    );
    pkg.apply_cell_patches(&patches)?;

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let doc = roxmltree::Document::parse(out_xml)?;
    let b1 = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("B1"))
        .expect("expected inserted B1 cell");
    assert_eq!(b1.attribute("vm"), Some("9"));
    assert_eq!(b1.attribute("cm"), Some("7"));
    Ok(())
}

#[test]
fn in_memory_cell_patches_can_clear_vm_on_existing_cell() -> Result<(), Box<dyn std::error::Error>>
{
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" vm="9" cm="7"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(2.0)).clear_vm(),
    );
    pkg.apply_cell_patches(&patches)?;

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let doc = roxmltree::Document::parse(out_xml)?;
    let a1 = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");
    assert_eq!(a1.attribute("vm"), None);
    assert_eq!(a1.attribute("cm"), Some("7"), "expected cm to be preserved");
    Ok(())
}

#[test]
fn streaming_workbook_cell_patches_can_set_vm_cm_on_inserted_cell(
) -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(2.0)),
    );
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("B1")?,
        CellPatch::set_value(CellValue::Number(3.0))
            .with_vm(9)
            .with_cm(7),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut out, &patches)?;

    let out_xml = read_sheet1_xml_from_xlsx(out.get_ref());
    let doc = roxmltree::Document::parse(&out_xml)?;
    let b1 = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("B1"))
        .expect("expected inserted B1 cell");
    assert_eq!(b1.attribute("vm"), Some("9"));
    assert_eq!(b1.attribute("cm"), Some("7"));
    Ok(())
}

#[test]
fn streaming_workbook_cell_patches_can_clear_vm_on_existing_cell(
) -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" vm="9" cm="7"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(2.0)).clear_vm(),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut out, &patches)?;

    let out_xml = read_sheet1_xml_from_xlsx(out.get_ref());
    let doc = roxmltree::Document::parse(&out_xml)?;
    let a1 = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");
    assert_eq!(a1.attribute("vm"), None);
    assert_eq!(a1.attribute("cm"), Some("7"), "expected cm to be preserved");
    Ok(())
}

