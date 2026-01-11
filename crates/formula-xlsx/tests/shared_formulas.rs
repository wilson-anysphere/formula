use std::io::Read;
use std::path::Path;
use std::{fs, io::Cursor};

use formula_model::CellRef;
use pretty_assertions::assert_eq;
use zip::ZipArchive;

#[test]
fn imports_shared_formula_followers_as_formulas() -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/formulas/shared-formula.xlsx");
    let doc = formula_xlsx::load_from_path(&fixture)?;

    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).unwrap();

    assert_eq!(sheet.formula(CellRef::from_a1("A2")?), Some("B2*2"));
    assert_eq!(sheet.formula(CellRef::from_a1("A3")?), Some("B3*2"));

    Ok(())
}

#[test]
fn roundtrip_preserves_textless_shared_formula_xml() -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/formulas/shared-formula.xlsx");
    let doc = formula_xlsx::load_from_path(&fixture)?;
    let bytes = doc.save_to_vec()?;

    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("roundtripped.xlsx");
    fs::write(&out_path, &bytes)?;

    let report = xlsx_diff::diff_workbooks(&fixture, &out_path)?;
    assert_eq!(report.count(xlsx_diff::Severity::Critical), 0);

    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let needle = r#"<f t="shared" si="0"/>"#;
    assert_eq!(sheet_xml.matches(needle).count(), 2);
    assert!(
        !sheet_xml.contains(r#"<f t="shared" si="0">"#),
        "expected follower formulas to remain textless: {sheet_xml}"
    );

    Ok(())
}

