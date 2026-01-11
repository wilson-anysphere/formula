use std::io::Read;
use std::path::Path;

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;

#[test]
fn document_roundtrip_preserves_other_cell_types() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/date-type.xlsx");
    let bytes = std::fs::read(&fixture)?;

    let doc = load_from_bytes(&bytes)?;
    let saved = doc.save_to_vec()?;

    // Ensure the worksheet XML keeps the original `t="d"` + raw `<v>` payload.
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&saved))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let parsed = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let cell = parsed
        .descendants()
        .find(|n| n.is_element() && n.has_tag_name((ns, "c")) && n.attribute("r") == Some("C1"))
        .ok_or("missing C1")?;
    assert_eq!(cell.attribute("t"), Some("d"));
    let v = cell
        .children()
        .find(|n| n.is_element() && n.has_tag_name((ns, "v")))
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v, "2024-01-01T00:00:00Z");

    // Ensure the overall no-op round-trip doesn't introduce critical diffs.
    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("roundtripped.xlsx");
    std::fs::write(&out_path, &saved)?;

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
fn editing_other_cell_type_updates_value_but_preserves_t_attribute() -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/date-type.xlsx");
    let bytes = std::fs::read(&fixture)?;

    let mut doc = load_from_bytes(&bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).ok_or("missing sheet")?;
    sheet.set_value(
        CellRef::from_a1("C1")?,
        CellValue::String("2025-02-03T00:00:00Z".to_string()),
    );

    let saved = doc.save_to_vec()?;

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&saved))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let parsed = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let cell = parsed
        .descendants()
        .find(|n| n.is_element() && n.has_tag_name((ns, "c")) && n.attribute("r") == Some("C1"))
        .ok_or("missing C1")?;
    assert_eq!(cell.attribute("t"), Some("d"));
    let v = cell
        .children()
        .find(|n| n.is_element() && n.has_tag_name((ns, "v")))
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v, "2025-02-03T00:00:00Z");

    Ok(())
}
