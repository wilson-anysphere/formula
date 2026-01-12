use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    patch_xlsx_streaming_workbook_cell_patches, CellPatch, WorkbookCellPatches, XlsxPackage,
};
use zip::ZipArchive;

fn build_minimal_a1_fixture() -> Vec<u8> {
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

    // Intentionally minimal: a single cell A1 and a tight dimension.
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

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
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_sparse_row_fixture() -> Vec<u8> {
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

    // Sparse row: A1 and C1 exist, B1 is missing but within the used-range bounding box.
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c></row>
  </sheetData>
</worksheet>"#;

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
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_minimal_empty_sheet_fixture() -> Vec<u8> {
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
  <dimension ref="A1"/>
  <sheetData/>
</worksheet>"#;

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
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

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

fn dimension_ref(sheet_xml: &str) -> String {
    let doc = roxmltree::Document::parse(sheet_xml).expect("parse xml");
    doc.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "dimension")
        .and_then(|n| n.attribute("ref"))
        .unwrap_or_default()
        .to_string()
}

fn has_row(sheet_xml: &str, row: &str) -> bool {
    let doc = roxmltree::Document::parse(sheet_xml).expect("parse xml");
    doc.descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some(row))
}

fn has_cell(sheet_xml: &str, a1: &str) -> bool {
    let doc = roxmltree::Document::parse(sheet_xml).expect("parse xml");
    doc.descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some(a1))
}

fn cell_style(sheet_xml: &str, a1: &str) -> Option<String> {
    let doc = roxmltree::Document::parse(sheet_xml).expect("parse xml");
    doc.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some(a1))
        .and_then(|n| n.attribute("s"))
        .map(|s| s.to_string())
}

#[test]
fn streaming_noop_clear_does_not_insert_cell_or_expand_dimension(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_a1_fixture();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell("Sheet1", CellRef::from_a1("Z100")?, CellPatch::clear());

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes.clone()), &mut out, &patches)?;

    let sheet_xml = read_sheet1_xml_from_xlsx(out.get_ref());
    assert!(!has_cell(&sheet_xml, "Z100"));
    assert!(!has_row(&sheet_xml, "100"));
    assert_eq!(dimension_ref(&sheet_xml), "A1");
    assert_eq!(
        out.get_ref().as_slice(),
        bytes.as_slice(),
        "expected streaming patcher to leave the zip unchanged for a no-op clear patch"
    );

    Ok(())
}

#[test]
fn in_memory_noop_clear_does_not_insert_cell_or_expand_dimension(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_a1_fixture();
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;
    let original_sheet = pkg.part("xl/worksheets/sheet1.xml").unwrap().to_vec();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell("Sheet1", CellRef::from_a1("Z100")?, CellPatch::clear());
    pkg.apply_cell_patches(&patches)?;

    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(!has_cell(sheet_xml, "Z100"));
    assert!(!has_row(sheet_xml, "100"));
    assert_eq!(dimension_ref(sheet_xml), "A1");
    assert_eq!(
        pkg.part("xl/worksheets/sheet1.xml").unwrap(),
        original_sheet.as_slice(),
        "expected in-memory patcher to preserve sheet XML bytes for a no-op clear patch"
    );

    Ok(())
}

#[test]
fn streaming_empty_formula_is_treated_as_noop_clear_for_missing_cell(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_a1_fixture();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("Z100")?,
        CellPatch::set_value_with_formula(CellValue::Empty, "="),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes.clone()), &mut out, &patches)?;

    let sheet_xml = read_sheet1_xml_from_xlsx(out.get_ref());
    assert!(!has_cell(&sheet_xml, "Z100"));
    assert!(!has_row(&sheet_xml, "100"));
    assert_eq!(dimension_ref(&sheet_xml), "A1");
    assert_eq!(
        out.get_ref().as_slice(),
        bytes.as_slice(),
        "expected streaming patcher to leave the zip unchanged for a no-op empty-formula patch"
    );

    Ok(())
}

#[test]
fn in_memory_empty_formula_is_treated_as_noop_clear_for_missing_cell(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_a1_fixture();
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;
    let original_sheet = pkg.part("xl/worksheets/sheet1.xml").unwrap().to_vec();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("Z100")?,
        CellPatch::set_value_with_formula(CellValue::Empty, "="),
    );
    pkg.apply_cell_patches(&patches)?;

    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(!has_cell(sheet_xml, "Z100"));
    assert!(!has_row(sheet_xml, "100"));
    assert_eq!(dimension_ref(sheet_xml), "A1");
    assert_eq!(
        pkg.part("xl/worksheets/sheet1.xml").unwrap(),
        original_sheet.as_slice(),
        "expected in-memory patcher to preserve sheet XML bytes for a no-op empty-formula patch"
    );

    Ok(())
}

#[test]
fn streaming_noop_clear_missing_cell_within_used_range_does_not_rewrite_zip(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_sparse_row_fixture();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell("Sheet1", CellRef::from_a1("B1")?, CellPatch::clear());

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes.clone()), &mut out, &patches)?;

    let sheet_xml = read_sheet1_xml_from_xlsx(out.get_ref());
    assert!(!has_cell(&sheet_xml, "B1"));
    assert_eq!(dimension_ref(&sheet_xml), "A1:C1");
    assert_eq!(
        out.get_ref().as_slice(),
        bytes.as_slice(),
        "expected streaming patcher to leave the zip unchanged for a no-op clear patch"
    );

    Ok(())
}

#[test]
fn in_memory_noop_clear_missing_cell_within_used_range_does_not_rewrite_sheet_xml(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_sparse_row_fixture();
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;
    let original_sheet = pkg.part("xl/worksheets/sheet1.xml").unwrap().to_vec();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell("Sheet1", CellRef::from_a1("B1")?, CellPatch::clear());
    pkg.apply_cell_patches(&patches)?;

    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(!has_cell(sheet_xml, "B1"));
    assert_eq!(dimension_ref(sheet_xml), "A1:C1");
    assert_eq!(
        pkg.part("xl/worksheets/sheet1.xml").unwrap(),
        original_sheet.as_slice(),
        "expected in-memory patcher to preserve sheet XML bytes for a no-op clear patch"
    );

    Ok(())
}

#[test]
fn streaming_clear_with_style_inserts_cell_and_expands_dimension(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_a1_fixture();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("Z100")?,
        CellPatch::clear_with_style(1),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut out, &patches)?;

    let sheet_xml = read_sheet1_xml_from_xlsx(out.get_ref());
    assert!(has_cell(&sheet_xml, "Z100"));
    assert!(has_row(&sheet_xml, "100"));
    assert_eq!(dimension_ref(&sheet_xml), "A1:Z100");
    assert_eq!(cell_style(&sheet_xml, "Z100").as_deref(), Some("1"));

    Ok(())
}

#[test]
fn in_memory_clear_with_style_inserts_cell_and_expands_dimension(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_a1_fixture();
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("Z100")?,
        CellPatch::clear_with_style(1),
    );
    pkg.apply_cell_patches(&patches)?;

    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(has_cell(sheet_xml, "Z100"));
    assert!(has_row(sheet_xml, "100"));
    assert_eq!(dimension_ref(sheet_xml), "A1:Z100");
    assert_eq!(cell_style(sheet_xml, "Z100").as_deref(), Some("1"));

    Ok(())
}

#[test]
fn streaming_noop_clear_does_not_expand_empty_sheetdata() -> Result<(), Box<dyn std::error::Error>>
{
    let bytes = build_minimal_empty_sheet_fixture();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell("Sheet1", CellRef::from_a1("Z100")?, CellPatch::clear());

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes.clone()), &mut out, &patches)?;

    let sheet_xml = read_sheet1_xml_from_xlsx(out.get_ref());
    assert!(
        !sheet_xml.contains("</sheetData>"),
        "expected sheetData to remain an empty element"
    );
    assert!(!has_cell(&sheet_xml, "Z100"));
    assert!(!has_row(&sheet_xml, "100"));
    assert_eq!(dimension_ref(&sheet_xml), "A1");
    assert_eq!(
        out.get_ref().as_slice(),
        bytes.as_slice(),
        "expected streaming patcher to leave the zip unchanged for a no-op clear patch"
    );

    Ok(())
}

#[test]
fn in_memory_noop_clear_does_not_expand_empty_sheetdata() -> Result<(), Box<dyn std::error::Error>>
{
    let bytes = build_minimal_empty_sheet_fixture();
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;
    let original_sheet = pkg.part("xl/worksheets/sheet1.xml").unwrap().to_vec();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell("Sheet1", CellRef::from_a1("Z100")?, CellPatch::clear());
    pkg.apply_cell_patches(&patches)?;

    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        !sheet_xml.contains("</sheetData>"),
        "expected sheetData to remain an empty element"
    );
    assert!(!has_cell(sheet_xml, "Z100"));
    assert!(!has_row(sheet_xml, "100"));
    assert_eq!(dimension_ref(sheet_xml), "A1");
    assert_eq!(
        pkg.part("xl/worksheets/sheet1.xml").unwrap(),
        original_sheet.as_slice(),
        "expected in-memory patcher to preserve sheet XML bytes for a no-op clear patch"
    );

    Ok(())
}
