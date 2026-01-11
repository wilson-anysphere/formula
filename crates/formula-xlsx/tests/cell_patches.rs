use std::collections::BTreeSet;
use std::io::{Cursor, Write};
use std::path::Path;

use formula_model::{CellRef, CellValue};
use pretty_assertions::assert_eq;

use formula_xlsx::{CellPatch, WorkbookCellPatches, XlsxPackage};

fn diff_parts(expected: &Path, actual: &Path) -> BTreeSet<String> {
    let report = xlsx_diff::diff_workbooks(expected, actual).expect("diff workbooks");
    for diff in &report.differences {
        assert_ne!(diff.kind, "missing_part", "missing part {}", diff.part);
        assert_ne!(diff.kind, "extra_part", "extra part {}", diff.part);
    }
    report
        .differences
        .iter()
        .map(|d| d.part.clone())
        .collect::<BTreeSet<_>>()
}

fn worksheet_cell_formula(sheet_xml: &str, cell_ref: &str) -> Option<String> {
    let xml_doc = roxmltree::Document::parse(sheet_xml).ok()?;
    let cell = xml_doc.descendants().find(|n| {
        n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some(cell_ref)
    })?;
    let f = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "f")?;
    Some(f.text().unwrap_or_default().to_string())
}

fn worksheet_dimension_ref(sheet_xml: &[u8]) -> Result<String, Box<dyn std::error::Error>> {
    let xml = std::str::from_utf8(sheet_xml)?;
    let doc = roxmltree::Document::parse(xml)?;
    Ok(doc
        .root_element()
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "dimension")
        .and_then(|n| n.attribute("ref"))
        .unwrap_or("A1")
        .to_string())
}

fn build_sheetpr_no_dimension_fixture() -> Vec<u8> {
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
  <sheetPr/>
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

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn apply_cell_patches_preserves_unrelated_parts_for_xlsx() -> Result<(), Box<dyn std::error::Error>>
{
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/basic.xlsx");
    let bytes = std::fs::read(&fixture)?;
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        // Sheet names are case-insensitive in Excel; allow patches keyed by any casing.
        "sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(42.0)),
    );

    pkg.apply_cell_patches(&patches)?;

    let tmpdir = tempfile::tempdir()?;
    let out = tmpdir.path().join("patched.xlsx");
    std::fs::write(&out, pkg.write_to_bytes()?)?;

    let parts = diff_parts(&fixture, &out);
    assert_eq!(
        parts,
        BTreeSet::from(["xl/worksheets/sheet1.xml".to_string()])
    );
    Ok(())
}

#[test]
fn apply_cell_patches_preserves_unknown_cell_types() -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/date-type.xlsx");
    let bytes = std::fs::read(&fixture)?;
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("C1")?,
        CellPatch::set_value(CellValue::String("2027-04-05T00:00:00Z".to_string())),
    );
    pkg.apply_cell_patches(&patches)?;

    let tmpdir = tempfile::tempdir()?;
    let out = tmpdir.path().join("patched.xlsx");
    std::fs::write(&out, pkg.write_to_bytes()?)?;

    // Only the worksheet part should change (no sharedStrings churn).
    let parts = diff_parts(&fixture, &out);
    assert_eq!(
        parts,
        BTreeSet::from(["xl/worksheets/sheet1.xml".to_string()])
    );

    // And the original unknown `t="d"` cell type should be preserved.
    let out_bytes = std::fs::read(&out)?;
    let pkg2 = XlsxPackage::from_bytes(&out_bytes)?;
    let sheet_xml = std::str::from_utf8(
        pkg2.part("xl/worksheets/sheet1.xml")
            .expect("worksheet part exists"),
    )?;
    assert!(
        sheet_xml.contains(r#"<c r="C1" t="d"><v>2027-04-05T00:00:00Z</v></c>"#),
        "expected patched worksheet xml to keep t=\"d\""
    );

    Ok(())
}

#[test]
fn apply_cell_patches_updates_worksheet_dimension_when_range_expands(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/grouped-rows.xlsx");
    let bytes = std::fs::read(&fixture)?;
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("Z100")?,
        CellPatch::set_value(CellValue::Number(1.0)),
    );
    pkg.apply_cell_patches(&patches)?;

    let out_bytes = pkg.write_to_bytes()?;
    let pkg2 = XlsxPackage::from_bytes(&out_bytes)?;
    let sheet_xml = pkg2
        .part("xl/worksheets/sheet1.xml")
        .expect("worksheet part exists");
    assert_eq!(worksheet_dimension_ref(sheet_xml)?, "A1:Z100");

    let tmpdir = tempfile::tempdir()?;
    let out = tmpdir.path().join("patched.xlsx");
    std::fs::write(&out, out_bytes)?;
    let parts = diff_parts(&fixture, &out);
    assert_eq!(
        parts,
        BTreeSet::from(["xl/worksheets/sheet1.xml".to_string()])
    );
    Ok(())
}

#[test]
fn apply_cell_patches_preserves_vba_project_for_xlsm() -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/macros/basic.xlsm");
    let bytes = std::fs::read(&fixture)?;
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let original_vba = pkg
        .part("xl/vbaProject.bin")
        .expect("fixture has vbaProject.bin")
        .to_vec();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(1.0)),
    );
    pkg.apply_cell_patches(&patches)?;

    let out_bytes = pkg.write_to_bytes()?;
    let pkg2 = XlsxPackage::from_bytes(&out_bytes)?;
    assert_eq!(
        pkg2.part("xl/vbaProject.bin").unwrap(),
        original_vba.as_slice()
    );

    let tmpdir = tempfile::tempdir()?;
    let out = tmpdir.path().join("patched.xlsm");
    std::fs::write(&out, out_bytes)?;

    let parts = diff_parts(&fixture, &out);
    assert_eq!(
        parts,
        BTreeSet::from(["xl/worksheets/sheet1.xml".to_string()])
    );
    Ok(())
}

#[test]
fn apply_cell_patches_drops_calc_chain_when_formulas_change(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = include_bytes!("fixtures/calc_settings.xlsx");
    let mut pkg = XlsxPackage::from_bytes(bytes)?;
    assert!(
        pkg.part("xl/calcChain.xml").is_some(),
        "fixture should include calcChain.xml"
    );

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value_with_formula(CellValue::Number(2.0), " =1+1"),
    );
    pkg.apply_cell_patches(&patches)?;

    assert!(pkg.part("xl/calcChain.xml").is_none());
    let workbook_xml = std::str::from_utf8(pkg.part("xl/workbook.xml").unwrap())?;
    assert!(
        workbook_xml.contains("fullCalcOnLoad=\"1\""),
        "workbook.xml should request full recalculation on load when formulas change"
    );

    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap())?;
    let formula = worksheet_cell_formula(sheet_xml, "A1").expect("patched cell should have <f>");
    assert!(
        !formula.trim_start().starts_with('='),
        "patched <f> text must not include a leading '=' (got {formula:?})"
    );
    assert_eq!(formula, "1+1");

    Ok(())
}

#[test]
fn apply_cell_patches_inserts_dimension_and_updates_row_spans(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/row-col-attrs.xlsx");
    let bytes = std::fs::read(&fixture)?;
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("C1")?,
        CellPatch::set_value(CellValue::Number(42.0)),
    );
    pkg.apply_cell_patches(&patches)?;

    let sheet_xml = pkg
        .part("xl/worksheets/sheet1.xml")
        .expect("worksheet part exists");
    assert_eq!(worksheet_dimension_ref(sheet_xml)?, "A1:C3");

    let xml = std::str::from_utf8(sheet_xml)?;
    let doc = roxmltree::Document::parse(xml)?;
    let row1 = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("1"))
        .expect("row r=1 exists");
    assert_eq!(row1.attribute("spans"), Some("1:3"));

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("B4")?,
        CellPatch::set_value(CellValue::Number(4.0)),
    );
    pkg.apply_cell_patches(&patches)?;

    let sheet_xml = pkg
        .part("xl/worksheets/sheet1.xml")
        .expect("worksheet part exists");
    let xml = std::str::from_utf8(sheet_xml)?;
    let doc = roxmltree::Document::parse(xml)?;
    let row4 = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("4"))
        .expect("row r=4 should be inserted");
    assert_eq!(row4.attribute("spans"), Some("2:2"));

    Ok(())
}

#[test]
fn apply_cell_patches_inserts_dimension_after_sheetpr_when_missing(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_sheetpr_no_dimension_fixture();
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("B2")?,
        CellPatch::set_value(CellValue::Number(2.0)),
    );
    pkg.apply_cell_patches(&patches)?;

    let sheet_xml = pkg
        .part("xl/worksheets/sheet1.xml")
        .expect("worksheet part exists");
    assert_eq!(worksheet_dimension_ref(sheet_xml)?, "A1:B2");

    let xml = std::str::from_utf8(sheet_xml)?;
    let pos_sheet_pr = xml.find("<sheetPr").expect("sheetPr exists");
    let pos_dimension = xml.find("<dimension").expect("dimension inserted");
    let pos_sheet_data = xml.find("<sheetData").expect("sheetData exists");
    assert!(
        pos_sheet_pr < pos_dimension && pos_dimension < pos_sheet_data,
        "expected dimension to be inserted after sheetPr and before sheetData: {xml}"
    );

    Ok(())
}
