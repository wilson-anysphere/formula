use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use formula_model::{CellRef, CellValue};
use formula_xlsx::{patch_xlsx_streaming, WorksheetCellPatch};
use zip::ZipArchive;

#[test]
fn streaming_patch_preserves_vm_on_unpatched_cells_when_inserting_missing_cells(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/row-col-attrs.xlsx");
    let bytes = fs::read(&fixture_path)?;

    // B2 does not exist in the fixture; patching it forces the streaming patcher to insert a new
    // cell into row 2, while preserving the existing A2 cell (`vm="1"`).
    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("B2")?,
        CellValue::Number(99.0),
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
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    let a2 = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("A2"))
        .expect("A2 cell should exist");
    assert_eq!(a2.attribute("vm"), Some("1"), "A2 should preserve vm");
    let a2_v = a2
        .children()
        .find(|n| n.has_tag_name((ns, "v")))
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(a2_v, "2", "unpatched A2 value should be unchanged");

    let b2 = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("B2"))
        .expect("B2 cell should be inserted");
    let b2_v = b2
        .children()
        .find(|n| n.has_tag_name((ns, "v")))
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(b2_v, "99");

    Ok(())
}

#[test]
fn streaming_patch_preserves_vm_error_cells_when_inserting_missing_cells(
) -> Result<(), Box<dyn std::error::Error>> {
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipWriter};

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="e" vm="1"><v>#VALUE!</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let mut input = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut input);
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);
        zip.start_file("xl/worksheets/sheet1.xml", options)?;
        zip.write_all(worksheet_xml.as_bytes())?;
        zip.finish()?;
    }
    input.set_position(0);

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("B1")?,
        CellValue::Number(7.0),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(input, &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.into_inner()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    let a1 = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("A1"))
        .expect("A1 cell should exist");
    assert_eq!(a1.attribute("t"), Some("e"));
    assert_eq!(a1.attribute("vm"), Some("1"));
    let a1_v = a1
        .children()
        .find(|n| n.has_tag_name((ns, "v")))
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(a1_v, "#VALUE!");

    let b1 = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("B1"))
        .expect("B1 cell should be inserted");
    let b1_v = b1
        .children()
        .find(|n| n.has_tag_name((ns, "v")))
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(b1_v, "7");

    Ok(())
}

#[test]
fn streaming_patch_preserves_vm_on_patched_cells_while_inserting_others(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/row-col-attrs.xlsx");
    let bytes = fs::read(&fixture_path)?;

    // Patch A2 (existing cell with vm="1") and also insert B2 (missing).
    let patches = vec![
        WorksheetCellPatch::new(
            "xl/worksheets/sheet1.xml",
            CellRef::from_a1("A2")?,
            CellValue::Number(123.0),
            None,
        ),
        WorksheetCellPatch::new(
            "xl/worksheets/sheet1.xml",
            CellRef::from_a1("B2")?,
            CellValue::Number(99.0),
            None,
        ),
    ];

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &patches)?;

    let mut archive = ZipArchive::new(Cursor::new(out.into_inner()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    let a2 = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("A2"))
        .expect("A2 cell should exist");
    assert_eq!(a2.attribute("vm"), Some("1"), "A2 should preserve vm after patching value");
    let a2_v = a2
        .children()
        .find(|n| n.has_tag_name((ns, "v")))
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(a2_v, "123");

    let b2 = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("B2"))
        .expect("B2 cell should be inserted");
    let b2_v = b2
        .children()
        .find(|n| n.has_tag_name((ns, "v")))
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(b2_v, "99");

    Ok(())
}
