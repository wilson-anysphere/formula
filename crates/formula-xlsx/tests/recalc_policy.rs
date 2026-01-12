use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::CellRef;
use formula_xlsx::{load_from_bytes, CellPatch, RecalcPolicy, WorkbookCellPatches, XlsxPackage};
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
fn formula_patch_can_clear_cached_values_when_requested() -> Result<(), Box<dyn std::error::Error>>
{
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
        CellRef::from_a1("C1")?,
        CellPatch::set_value_with_formula(formula_model::CellValue::Number(43.0), "B1+1"),
    );

    pkg.apply_cell_patches_with_recalc_policy(
        &patches,
        RecalcPolicy {
            clear_cached_values_on_formula_change: true,
            ..Default::default()
        },
    )
    .expect("apply patches");

    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap())?;
    let sheet_doc = roxmltree::Document::parse(sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let cell = sheet_doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("C1"))
        .expect("C1 cell should exist");

    assert!(
        cell.descendants().any(|n| n.has_tag_name((ns, "f"))),
        "expected patched formula cell to contain <f>"
    );
    assert!(
        !cell
            .descendants()
            .any(|n| n.has_tag_name((ns, "v")) || n.has_tag_name((ns, "is"))),
        "expected patched formula cell to omit cached value (<v>/<is>)"
    );

    Ok(())
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

#[test]
fn formula_edit_can_clear_cached_values_when_requested() -> Result<(), Box<dyn std::error::Error>> {
    let mut doc = load_from_bytes(DOC_FIXTURE).expect("load fixture");
    let sheet_id = doc.workbook.sheets[0].id;

    assert!(
        doc.set_cell_formula(
            sheet_id,
            CellRef::from_a1("C1")?,
            Some("=SEQUENCE(2)".to_string())
        ),
        "expected formula edit to succeed"
    );

    let saved = doc.save_to_vec_with_recalc_policy(RecalcPolicy {
        clear_cached_values_on_formula_change: true,
        ..Default::default()
    })?;

    let mut archive = ZipArchive::new(Cursor::new(&saved))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let sheet_doc = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let cell = sheet_doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("C1"))
        .expect("C1 cell should exist");

    assert!(
        cell.descendants().any(|n| n.has_tag_name((ns, "f"))),
        "expected saved formula cell to contain <f>"
    );
    assert!(
        !cell
            .descendants()
            .any(|n| n.has_tag_name((ns, "v")) || n.has_tag_name((ns, "is"))),
        "expected saved formula cell to omit cached value (<v>/<is>)"
    );

    Ok(())
}
