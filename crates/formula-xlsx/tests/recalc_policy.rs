use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::CellRef;
use formula_xlsx::{load_from_bytes, CellPatch, WorkbookCellPatches, XlsxPackage};
use zip::ZipArchive;

const DOC_FIXTURE: &[u8] = include_bytes!("fixtures/recalc_policy.xlsx");
const PATCH_FIXTURE: &[u8] = include_bytes!("fixtures/rt_simple.xlsx");

fn recalc_fixture_path() -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/recalc_policy.xlsx")
}

#[test]
fn noop_save_preserves_calc_chain() -> Result<(), Box<dyn std::error::Error>> {
    let doc = load_from_bytes(DOC_FIXTURE).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    let mut archive = ZipArchive::new(Cursor::new(&saved))?;
    archive.by_name("xl/calcChain.xml")?;

    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("out.xlsx");
    std::fs::write(&out_path, &saved)?;

    let report = xlsx_diff::diff_workbooks(&recalc_fixture_path(), &out_path)?;
    if report.has_at_least(xlsx_diff::Severity::Critical) {
        eprintln!("Critical diffs detected for recalc policy no-op fixture");
        for diff in report
            .differences
            .iter()
            .filter(|d| d.severity == xlsx_diff::Severity::Critical)
        {
            eprintln!("{diff}");
        }
        panic!("no-op save did not round-trip cleanly");
    }

    Ok(())
}

#[test]
fn formula_edit_drops_calc_chain_and_sets_full_calc_on_load(
) -> Result<(), Box<dyn std::error::Error>> {
    let mut doc = load_from_bytes(DOC_FIXTURE).expect("load fixture");
    let sheet_id = doc.workbook.sheets[0].id;

    assert!(
        doc.set_cell_formula(
            sheet_id,
            CellRef::from_a1("C1")?,
            Some("=SEQUENCE(2)".to_string()),
        ),
        "expected formula edit to succeed"
    );

    let saved = doc.save_to_vec().expect("save");
    let mut archive = ZipArchive::new(Cursor::new(&saved))?;
    assert!(
        archive.by_name("xl/calcChain.xml").is_err(),
        "expected calcChain.xml to be removed after formula edit"
    );

    let mut workbook_xml = String::new();
    archive
        .by_name("xl/workbook.xml")?
        .read_to_string(&mut workbook_xml)?;
    assert!(
        workbook_xml.contains(r#"fullCalcOnLoad="1""#),
        "expected workbook.xml to set calcPr fullCalcOnLoad=1"
    );

    Ok(())
}

#[test]
fn formula_patch_sets_full_calc_on_load_and_drops_calc_chain() {
    let mut pkg = XlsxPackage::from_bytes(PATCH_FIXTURE).expect("read package");
    let sheet_name = pkg
        .workbook_sheets()
        .expect("parse workbook sheets")
        .first()
        .expect("fixture should have at least one sheet")
        .name
        .clone();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        sheet_name,
        CellRef::from_a1("C1").unwrap(),
        CellPatch::set_value_with_formula(formula_model::CellValue::Number(43.0), "B1+1"),
    );

    pkg.apply_cell_patches(&patches).expect("apply patches");

    assert!(
        pkg.part("xl/calcChain.xml").is_none(),
        "calcChain.xml should be removed after formula edits"
    );

    let workbook_xml = std::str::from_utf8(pkg.part("xl/workbook.xml").unwrap()).unwrap();
    assert!(
        workbook_xml.contains(r#"fullCalcOnLoad="1""#),
        "workbook.xml should force full calc on load"
    );

    let ct = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap()).unwrap();
    assert!(
        !ct.contains("/xl/calcChain.xml"),
        "content types override for calcChain.xml should be removed"
    );

    let rels = std::str::from_utf8(pkg.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
    assert!(
        !rels.contains("calcChain.xml"),
        "workbook.xml.rels relationship targeting calcChain.xml should be removed"
    );
}

#[test]
fn literal_patch_preserves_calc_chain_and_calc_pr() {
    let mut pkg = XlsxPackage::from_bytes(PATCH_FIXTURE).expect("read package");
    let sheet_name = pkg
        .workbook_sheets()
        .expect("parse workbook sheets")
        .first()
        .expect("fixture should have at least one sheet")
        .name
        .clone();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        sheet_name,
        CellRef::from_a1("B1").unwrap(),
        CellPatch::set_value(formula_model::CellValue::Number(43.0)),
    );

    pkg.apply_cell_patches(&patches).expect("apply patches");

    assert!(
        pkg.part("xl/calcChain.xml").is_some(),
        "calcChain.xml should be preserved for non-formula edits"
    );

    let workbook_xml = std::str::from_utf8(pkg.part("xl/workbook.xml").unwrap()).unwrap();
    assert!(
        workbook_xml.contains(r#"fullCalcOnLoad="0""#),
        "non-formula edits should preserve existing calcPr attributes"
    );
}
