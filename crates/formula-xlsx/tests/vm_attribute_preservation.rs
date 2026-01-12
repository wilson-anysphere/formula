use std::io::{Cursor, Read as _, Write as _};

use formula_model::{CellRef, CellValue, ErrorValue};
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

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn read_zip_part_to_string(bytes: &[u8], name: &str) -> String {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut file = archive.by_name(name).expect("zip entry exists");
    let mut out = String::new();
    file.read_to_string(&mut out).expect("read xml");
    out
}

fn build_minimal_vm_xlsx() -> Vec<u8> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A2"/>
  <sheetData>
    <row r="2">
      <c r="A2" vm="1"><v>5</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    build_minimal_xlsx(worksheet_xml)
}

fn assert_a2_value_and_preserves_vm(worksheet_xml: &str, expected_value: &str, expected_vm: &str) {
    let doc = roxmltree::Document::parse(worksheet_xml).expect("parse worksheet xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A2"))
        .expect("expected A2 cell");
    assert_eq!(
        cell.attribute("vm"),
        Some(expected_vm),
        "vm should be preserved on value edits for non-placeholder rich-data cells, got: {worksheet_xml}"
    );
    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(
        v, expected_value,
        "expected patched cached value in A2 to be {expected_value}, got: {worksheet_xml}"
    );
}

fn assert_a2_error_value_and_preserves_vm(worksheet_xml: &str, expected_vm: &str) {
    let doc = roxmltree::Document::parse(worksheet_xml).expect("parse worksheet xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A2"))
        .expect("expected A2 cell");
    assert_eq!(
        cell.attribute("vm"),
        Some(expected_vm),
        "vm should be preserved when patching to the rich-value placeholder error (#VALUE!), got: {worksheet_xml}"
    );
    assert_eq!(
        cell.attribute("t"),
        Some("e"),
        "expected patched A2 to use error cell type (t=\"e\"), got: {worksheet_xml}"
    );
    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(
        v,
        ErrorValue::Value.as_str(),
        "expected patched cached error value in A2 to be #VALUE!, got: {worksheet_xml}"
    );
}

#[test]
fn apply_cell_patches_to_bytes_preserves_vm_attribute_on_patched_cell(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_vm_xlsx();
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    let patch = PackageCellPatch::for_sheet_name(
        "Sheet1",
        CellRef::from_a1("A2")?,
        CellValue::Number(9.0),
        None,
    );

    let out_bytes = pkg.apply_cell_patches_to_bytes(&[patch])?;
    let out_xml = read_zip_part_to_string(&out_bytes, "xl/worksheets/sheet1.xml");
    assert_a2_value_and_preserves_vm(&out_xml, "9", "1");

    Ok(())
}

#[test]
fn apply_cell_patches_preserves_vm_attribute_on_patched_cell() -> Result<(), Box<dyn std::error::Error>>
{
    let bytes = build_minimal_vm_xlsx();
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A2")?,
        CellPatch::set_value(CellValue::Number(9.0)),
    );
    pkg.apply_cell_patches(&patches)?;

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap())?;
    assert_a2_value_and_preserves_vm(out_xml, "9", "1");

    Ok(())
}

#[test]
fn apply_cell_patches_to_bytes_preserves_vm_attribute_on_rich_value_placeholder_error(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_vm_xlsx();
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    let patch = PackageCellPatch::for_sheet_name(
        "Sheet1",
        CellRef::from_a1("A2")?,
        CellValue::Error(ErrorValue::Value),
        None,
    );

    let out_bytes = pkg.apply_cell_patches_to_bytes(&[patch])?;
    let out_xml = read_zip_part_to_string(&out_bytes, "xl/worksheets/sheet1.xml");
    assert_a2_error_value_and_preserves_vm(&out_xml, "1");

    Ok(())
}

#[test]
fn apply_cell_patches_preserves_vm_attribute_on_rich_value_placeholder_error(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_vm_xlsx();
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A2")?,
        CellPatch::set_value(CellValue::Error(ErrorValue::Value)),
    );
    pkg.apply_cell_patches(&patches)?;

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap())?;
    assert_a2_error_value_and_preserves_vm(out_xml, "1");

    Ok(())
}
