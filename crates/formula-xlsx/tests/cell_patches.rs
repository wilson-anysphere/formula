use std::collections::BTreeSet;
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

#[test]
fn apply_cell_patches_preserves_unrelated_parts_for_xlsx() -> Result<(), Box<dyn std::error::Error>>
{
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/basic.xlsx");
    let bytes = std::fs::read(&fixture)?;
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
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
        CellPatch::set_value_with_formula(CellValue::Number(2.0), "=1+1"),
    );
    pkg.apply_cell_patches(&patches)?;

    assert!(pkg.part("xl/calcChain.xml").is_none());
    let workbook_xml = std::str::from_utf8(pkg.part("xl/workbook.xml").unwrap())?;
    assert!(
        workbook_xml.contains("fullCalcOnLoad=\"1\""),
        "workbook.xml should request full recalculation on load when formulas change"
    );

    Ok(())
}
