use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use formula_model::{CellRef, CellValue};
use formula_xlsx::{patch_xlsx_streaming, WorksheetCellPatch};
use zip::ZipArchive;

#[test]
fn streaming_patch_inserts_dimension_and_updates_row_spans(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/row-col-attrs.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("C1")?,
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
    let worksheet = doc.root_element();

    let children: Vec<_> = worksheet.children().filter(|n| n.is_element()).collect();
    let dimension_idx = children
        .iter()
        .position(|n| n.tag_name().name() == "dimension")
        .expect("expected <dimension> element to be inserted");
    let cols_idx = children
        .iter()
        .position(|n| n.tag_name().name() == "cols")
        .expect("expected <cols> element to exist");
    assert!(
        dimension_idx < cols_idx,
        "<dimension> should appear before <cols> in worksheet schema order"
    );

    let dimension = children[dimension_idx];
    assert_eq!(
        dimension.attribute("ref"),
        Some("A1:C3"),
        "dimension should union original used range (A1:B3) and patched range (C1)"
    );

    let row_1 = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("1"))
        .expect("expected row r=\"1\" to exist");
    assert_eq!(
        row_1.attribute("spans"),
        Some("1:3"),
        "row 1 spans should expand to cover inserted C1 cell"
    );

    Ok(())
}

#[test]
fn streaming_patch_emits_spans_on_inserted_rows() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/row-col-attrs.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("B4")?,
        CellValue::Number(4.0),
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
    let row_4 = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("4"))
        .expect("expected row r=\"4\" to be inserted");
    assert_eq!(
        row_4.attribute("spans"),
        Some("2:2"),
        "inserted row spans should match inserted cell columns"
    );

    Ok(())
}

#[test]
fn streaming_patch_inserts_dimension_after_sheet_pr() -> Result<(), Box<dyn std::error::Error>> {
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipWriter};

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
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
        CellRef::from_a1("C1")?,
        CellValue::Number(2.0),
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
    let worksheet = doc.root_element();
    let children: Vec<_> = worksheet.children().filter(|n| n.is_element()).collect();
    assert_eq!(children[0].tag_name().name(), "sheetPr");
    assert_eq!(children[1].tag_name().name(), "dimension");
    assert_eq!(children[2].tag_name().name(), "sheetData");

    Ok(())
}

#[test]
fn streaming_patch_inserts_dimension_including_merge_cells(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/merged-cells.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::Number(123.0),
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
    let dimension = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "dimension")
        .expect("expected <dimension> element to be inserted");
    assert_eq!(dimension.attribute("ref"), Some("A1:B2"));

    Ok(())
}
