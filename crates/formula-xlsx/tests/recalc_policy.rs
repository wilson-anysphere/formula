use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::CellRef;
use formula_xlsx::load_from_bytes;
use zip::ZipArchive;

const FIXTURE: &[u8] = include_bytes!("fixtures/recalc_policy.xlsx");

fn fixture_path() -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/recalc_policy.xlsx")
}

#[test]
fn noop_save_preserves_calc_chain() -> Result<(), Box<dyn std::error::Error>> {
    let doc = load_from_bytes(FIXTURE).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    let mut archive = ZipArchive::new(Cursor::new(&saved))?;
    archive.by_name("xl/calcChain.xml")?;

    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("out.xlsx");
    std::fs::write(&out_path, &saved)?;

    let report = xlsx_diff::diff_workbooks(&fixture_path(), &out_path)?;
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
    let mut doc = load_from_bytes(FIXTURE).expect("load fixture");
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
