use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::CellValue;
use zip::ZipArchive;

fn read_zip_part(bytes: &[u8], name: &str) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let cursor = Cursor::new(bytes);
    let mut zip = ZipArchive::new(cursor)?;
    let mut file = zip.by_name(name)?;
    let mut out = Vec::new();
    file.read_to_end(&mut out)?;
    Ok(out)
}

fn read_dimension_ref_from_worksheet_xml(xml_bytes: &[u8]) -> Result<String, Box<dyn std::error::Error>> {
    let xml = std::str::from_utf8(xml_bytes)?;
    let doc = roxmltree::Document::parse(xml)?;
    let dimension = doc
        .root_element()
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "dimension")
        .ok_or("dimension element missing")?;
    Ok(dimension
        .attribute("ref")
        .ok_or("dimension ref missing")?
        .to_string())
}

#[test]
fn worksheet_dimension_shrinks_when_cells_removed() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/grouped-rows.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    // Expand the used range.
    let mut doc = formula_xlsx::load_from_bytes(&fixture_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).unwrap();
    sheet.set_value_a1("Z100", CellValue::Number(1.0))?;
    let expanded = doc.save_to_vec()?;

    // Then remove the boundary cell and ensure the dimension ref shrinks back.
    let mut doc2 = formula_xlsx::load_from_bytes(&expanded)?;
    let sheet_id = doc2.workbook.sheets[0].id;
    let sheet = doc2.workbook.sheet_mut(sheet_id).unwrap();
    sheet.clear_cell_a1("Z100")?;
    let out_bytes = doc2.save_to_vec()?;

    let worksheet_xml = read_zip_part(&out_bytes, "xl/worksheets/sheet1.xml")?;
    let dimension_ref = read_dimension_ref_from_worksheet_xml(&worksheet_xml)?;
    assert_eq!(dimension_ref, "A1:A5");

    // The final workbook should match the original fixture semantically.
    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("out.xlsx");
    std::fs::write(&out_path, &out_bytes)?;
    let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path)?;
    assert!(
        report.differences.is_empty(),
        "expected no diffs after roundtrip expansion+shrink: {:#?}",
        report.differences
    );

    Ok(())
}

#[test]
fn worksheet_dimension_updates_existing_ref_when_range_expands(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/grouped-rows.xlsx");
    let fixture_bytes = std::fs::read(&fixture)?;

    let mut doc = formula_xlsx::load_from_bytes(&fixture_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).unwrap();
    sheet.set_value_a1("Z100", CellValue::Number(1.0))?;

    let out_bytes = doc.save_to_vec()?;
    let worksheet_xml = read_zip_part(&out_bytes, "xl/worksheets/sheet1.xml")?;
    let dimension_ref = read_dimension_ref_from_worksheet_xml(&worksheet_xml)?;
    assert_eq!(dimension_ref, "A1:Z100");

    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("out.xlsx");
    std::fs::write(&out_path, &out_bytes)?;
    let report = xlsx_diff::diff_workbooks(&fixture, &out_path)?;
    assert!(
        !report.differences.is_empty(),
        "expected worksheet differences after edit"
    );
    assert!(
        report
            .differences
            .iter()
            .all(|diff| diff.part == "xl/worksheets/sheet1.xml"),
        "unexpected diffs in non-worksheet parts: {:#?}",
        report.differences
    );

    Ok(())
}

#[test]
fn worksheet_dimension_inserted_when_missing_and_range_expands(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/basic.xlsx");
    let fixture_bytes = std::fs::read(&fixture)?;

    let mut doc = formula_xlsx::load_from_bytes(&fixture_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).unwrap();
    sheet.set_value_a1("Z100", CellValue::Number(1.0))?;

    let out_bytes = doc.save_to_vec()?;
    let worksheet_xml = read_zip_part(&out_bytes, "xl/worksheets/sheet1.xml")?;
    let dimension_ref = read_dimension_ref_from_worksheet_xml(&worksheet_xml)?;
    assert_eq!(dimension_ref, "A1:Z100");

    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("out.xlsx");
    std::fs::write(&out_path, &out_bytes)?;
    let report = xlsx_diff::diff_workbooks(&fixture, &out_path)?;
    assert!(
        !report.differences.is_empty(),
        "expected worksheet differences after edit"
    );
    assert!(
        report
            .differences
            .iter()
            .all(|diff| diff.part == "xl/worksheets/sheet1.xml"),
        "unexpected diffs in non-worksheet parts: {:#?}",
        report.differences
    );

    Ok(())
}
