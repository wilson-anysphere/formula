use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{patch_xlsx_streaming, WorksheetCellPatch};
use zip::{write::FileOptions, CompressionMethod, ZipArchive, ZipWriter};

fn build_partial_zip_with_vm_cell() -> Vec<u8> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1" cm="2"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    // Intentionally omit `xl/workbook.xml` and `[Content_Types].xml` to simulate callers patching
    // an extracted worksheet part.
    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_minimal_xlsx() -> Vec<u8> {
    // The low-level streaming patcher preserves `vm` when patching "partial" ZIPs that only
    // contain worksheet XML (no `xl/workbook.xml`). Include `xl/workbook.xml` (and a minimal
    // `[Content_Types].xml`) so this fixture exercises the full-workbook behavior where patched
    // cells drop `vm`.
    let content_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

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

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1" cm="2"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types_xml.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn patch_xlsx_streaming_drops_vm_but_preserves_cm_on_existing_cells(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_xlsx();

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::Number(2.0),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.into_inner()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .ok_or("expected A1 cell")?;

    assert_eq!(
        cell.attribute("vm"),
        None,
        "vm should be dropped when patching away from rich-value placeholder semantics, got: {sheet_xml}"
    );
    assert_eq!(
        cell.attribute("cm"),
        Some("2"),
        "expected patched cell to preserve cm attribute, got: {sheet_xml}"
    );

    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(
        v, "2",
        "expected patch to update cached value (sanity check), got: {sheet_xml}"
    );

    Ok(())
}

#[test]
fn patch_xlsx_streaming_preserves_vm_cm_when_patching_partial_zip(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_partial_zip_with_vm_cell();

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::Number(2.0),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.into_inner()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .ok_or("expected A1 cell")?;

    assert_eq!(
        cell.attribute("vm"),
        Some("1"),
        "expected vm to be preserved when patching an extracted worksheet part, got: {sheet_xml}"
    );
    assert_eq!(cell.attribute("cm"), Some("2"));

    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v, "2");

    Ok(())
}
