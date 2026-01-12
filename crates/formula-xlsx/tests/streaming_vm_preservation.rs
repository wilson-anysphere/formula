use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use formula_model::{CellRef, CellValue, ErrorValue};
use formula_xlsx::{patch_xlsx_streaming_workbook_cell_patches, CellPatch, WorkbookCellPatches};
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

#[test]
fn streaming_patch_drops_vm_attribute_on_value_update() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/row-col-attrs.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A2")?,
        CellPatch::set_value(CellValue::Number(99.0)),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut out, &patches)?;

    let mut archive = ZipArchive::new(Cursor::new(out.into_inner()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let cell = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("A2"))
        .expect("expected cell A2 to exist");

    assert_eq!(
        cell.attribute("vm"),
        None,
        "vm should be dropped when patching away from rich-value placeholder semantics, got: {sheet_xml}"
    );

    let v_text = cell
        .children()
        .find(|n| n.has_tag_name((ns, "v")))
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(
        v_text, "99",
        "expected patched value in <v>, got: {sheet_xml}"
    );

    Ok(())
}

#[test]
fn streaming_patch_drops_vm_but_preserves_cm_on_value_update(
) -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1" cm="2"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(2.0)),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut out, &patches)?;

    let mut archive = ZipArchive::new(Cursor::new(out.into_inner()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let cell = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("A1"))
        .expect("expected cell A1 to exist");

    assert_eq!(cell.attribute("vm"), None);
    assert_eq!(cell.attribute("cm"), Some("2"));

    let v_text = cell
        .children()
        .find(|n| n.has_tag_name((ns, "v")))
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v_text, "2");

    Ok(())
}

#[test]
fn streaming_patch_preserves_vm_on_rich_value_placeholder_error(
) -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1" cm="2"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Error(ErrorValue::Value)),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut out, &patches)?;

    let mut archive = ZipArchive::new(Cursor::new(out.into_inner()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let cell = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("A1"))
        .expect("expected cell A1 to exist");

    assert_eq!(
        cell.attribute("vm"),
        Some("1"),
        "vm should be preserved when patching to the rich-value placeholder error (#VALUE!), got: {sheet_xml}"
    );
    assert_eq!(
        cell.attribute("cm"),
        Some("2"),
        "expected cm attribute to be preserved (sanity check), got: {sheet_xml}"
    );
    assert_eq!(
        cell.attribute("t"),
        Some("e"),
        "expected patched cell to use error cell type (t=\"e\"), got: {sheet_xml}"
    );

    let v_text = cell
        .children()
        .find(|n| n.has_tag_name((ns, "v")))
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(
        v_text,
        ErrorValue::Value.as_str(),
        "expected patched value in <v> to be #VALUE!, got: {sheet_xml}"
    );

    Ok(())
}
