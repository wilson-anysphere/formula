use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::{CellRef, CellValue};
use zip::ZipArchive;

#[test]
fn roundtrip_preserves_unknown_cell_type() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/metadata/date-iso-cell.xlsx");

    let doc = formula_xlsx::load_from_path(&fixture)?;
    let out_bytes = doc.save_to_vec()?;

    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("roundtripped.xlsx");
    std::fs::write(&out_path, out_bytes)?;

    let report = xlsx_diff::diff_workbooks(&fixture, &out_path)?;
    if report.has_at_least(xlsx_diff::Severity::Critical) {
        eprintln!("Critical diffs detected for fixture {}", fixture.display());
        for diff in report
            .differences
            .iter()
            .filter(|d| d.severity == xlsx_diff::Severity::Critical)
        {
            eprintln!("{diff}");
        }
        panic!("fixture {} did not round-trip cleanly", fixture.display());
    }

    Ok(())
}

#[test]
fn edited_iso_date_string_keeps_cell_type_d() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/metadata/date-iso-cell.xlsx");

    let mut doc = formula_xlsx::load_from_path(&fixture)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).unwrap();
    sheet.set_value(
        CellRef::from_a1("A1")?,
        CellValue::String("2024-02-01T00:00:00Z".to_string()),
    );

    let out_bytes = doc.save_to_vec()?;
    let mut archive = ZipArchive::new(Cursor::new(out_bytes))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let xml = roxmltree::Document::parse(&sheet_xml)?;
    let worksheet = xml.root_element();
    let cell = worksheet
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("missing A1 cell");
    assert_eq!(cell.attribute("t"), Some("d"));
    let v_text = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v_text, "2024-02-01T00:00:00Z");

    Ok(())
}
