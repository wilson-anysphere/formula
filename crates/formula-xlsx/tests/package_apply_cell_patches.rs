use std::fs;
use std::io::Read as _;
use std::path::PathBuf;

use formula_model::{CellRef, CellValue};
use formula_xlsx::{load_from_bytes, PackageCellPatch, XlsxPackage};
use xlsx_diff::Severity;

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures")
        .join(rel)
}

#[test]
fn apply_cell_patches_noop_roundtrip_has_no_critical_diffs() -> Result<(), Box<dyn std::error::Error>>
{
    let path = fixture_path("xlsx/basic/basic.xlsx");
    let bytes = fs::read(&path)?;
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    let out_bytes = pkg.apply_cell_patches_to_bytes(&[])?;

    let dir = tempfile::tempdir()?;
    let out_path = dir.path().join("patched.xlsx");
    fs::write(&out_path, out_bytes)?;

    let report = xlsx_diff::diff_workbooks(&path, &out_path)?;
    let critical = report.count(Severity::Critical);
    assert_eq!(
        critical, 0,
        "expected no critical diffs, got {critical}\n{}",
        report
            .differences
            .iter()
            .filter(|d| d.severity == Severity::Critical)
            .map(|d| d.to_string())
            .collect::<String>()
    );

    Ok(())
}

#[test]
fn apply_cell_patches_preserves_cell_style() -> Result<(), Box<dyn std::error::Error>> {
    let path = fixture_path("xlsx/styles/varied_styles.xlsx");
    let bytes = fs::read(&path)?;

    let before = load_from_bytes(&bytes)?;
    let sheet_id = before.workbook.sheets[0].id;
    let sheet = before.workbook.sheet(sheet_id).expect("sheet exists");

    let cell_ref = CellRef::from_a1("B1")?;
    let before_cell = sheet.cell(cell_ref).expect("B1 should exist");
    let before_style = before_cell.style_id;
    assert_ne!(before_style, 0, "fixture should have a non-default style on B1");

    let pkg = XlsxPackage::from_bytes(&bytes)?;
    let parts = pkg.worksheet_parts()?;
    assert_eq!(parts.len(), 1);
    assert_eq!(parts[0].name, "Sheet1");
    assert_eq!(parts[0].worksheet_part, "xl/worksheets/sheet1.xml");

    let patch = PackageCellPatch::for_sheet_name(
        "Sheet1",
        cell_ref,
        CellValue::String("patched".to_string()),
        None,
    );
    let out_bytes = pkg.apply_cell_patches_to_bytes(&[patch])?;

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&out_bytes))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    let b1_count = sheet_xml.matches(r#"r="B1""#).count();
    assert_eq!(
        b1_count, 1,
        "expected worksheet xml to contain exactly one B1 cell, found {b1_count}"
    );

    let after = load_from_bytes(&out_bytes)?;
    let sheet_id = after.workbook.sheets[0].id;
    let sheet = after.workbook.sheet(sheet_id).expect("sheet exists");
    let after_cell = sheet.cell(cell_ref).expect("B1 should exist after patch");

    assert_eq!(after_cell.value, CellValue::String("patched".to_string()));
    assert_eq!(
        after_cell.style_id, before_style,
        "patched cell should preserve existing style"
    );

    Ok(())
}

#[test]
fn apply_cell_patches_preserves_vba_project_bin_bytes() -> Result<(), Box<dyn std::error::Error>> {
    let path = fixture_path("xlsx/macros/basic.xlsm");
    let bytes = fs::read(&path)?;
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    let original = pkg
        .vba_project_bin()
        .expect("fixture should contain vbaProject.bin")
        .to_vec();

    let patch = PackageCellPatch::for_worksheet_part(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::Number(123.0),
        None,
    );
    let out_bytes = pkg.apply_cell_patches_to_bytes(&[patch])?;
    let patched = XlsxPackage::from_bytes(&out_bytes)?;

    let roundtrip = patched
        .vba_project_bin()
        .expect("patched workbook should still contain vbaProject.bin");
    assert_eq!(original, roundtrip);

    Ok(())
}

#[test]
fn apply_cell_patches_preserves_shared_strings() -> Result<(), Box<dyn std::error::Error>> {
    let path = fixture_path("xlsx/basic/shared-strings.xlsx");
    let bytes = fs::read(&path)?;
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    let patch = PackageCellPatch::for_sheet_name(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellValue::String("NewSharedString".to_string()),
        None,
    );
    let out_bytes = pkg.apply_cell_patches_to_bytes(&[patch])?;

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&out_bytes))?;

    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    let sheet_doc = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let cell = sheet_doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("A1"))
        .expect("A1 cell should exist");
    assert_eq!(
        cell.attribute("t"),
        Some("s"),
        "patched cell should remain a shared string cell"
    );

    let mut shared_strings_xml = String::new();
    archive
        .by_name("xl/sharedStrings.xml")?
        .read_to_string(&mut shared_strings_xml)?;
    assert!(
        shared_strings_xml.contains("NewSharedString"),
        "sharedStrings.xml should include the inserted string"
    );

    Ok(())
}

#[test]
fn apply_cell_patches_drops_calc_chain_when_formulas_change(
) -> Result<(), Box<dyn std::error::Error>> {
    use zip::result::ZipError;

    let bytes = include_bytes!("fixtures/calc_settings.xlsx");
    let pkg = XlsxPackage::from_bytes(bytes)?;

    let patch = PackageCellPatch::for_sheet_name(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellValue::Number(2.0),
        Some("=1+1".to_string()),
    );
    let out_bytes = pkg.apply_cell_patches_to_bytes(&[patch])?;

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&out_bytes))?;
    assert!(
        matches!(archive.by_name("xl/calcChain.xml").err(), Some(ZipError::FileNotFound)),
        "expected apply_cell_patches_to_bytes to drop xl/calcChain.xml after formula edits"
    );

    let mut workbook_xml = String::new();
    archive
        .by_name("xl/workbook.xml")?
        .read_to_string(&mut workbook_xml)?;
    let doc = roxmltree::Document::parse(&workbook_xml)?;
    let calc_pr = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "calcPr")
        .expect("workbook.xml should include <calcPr>");
    assert_eq!(
        calc_pr.attribute("fullCalcOnLoad"),
        Some("1"),
        "workbook.xml should request full recalculation on load when formulas change"
    );

    Ok(())
}
