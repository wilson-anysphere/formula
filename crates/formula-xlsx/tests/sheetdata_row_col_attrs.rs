use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::{Cell, CellRef, CellValue, Style, Workbook};
use formula_xlsx::{load_from_bytes, XlsxDocument};
use zip::ZipArchive;

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

#[test]
fn preserves_row_col_metadata_on_roundtrip() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/row-col-attrs.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    let doc = load_from_bytes(&fixture_bytes)?;

    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).expect("sheet exists");

    // Row 2 (1-based) has a custom height, and row 3 is hidden.
    assert_eq!(sheet.row_properties.get(&1).and_then(|p| p.height), Some(20.0));
    assert_eq!(sheet.row_properties.get(&2).map(|p| p.hidden), Some(true));

    // Column B has a custom width and column C is hidden.
    assert_eq!(sheet.col_properties.get(&1).and_then(|p| p.width), Some(25.0));
    assert_eq!(sheet.col_properties.get(&2).map(|p| p.hidden), Some(true));

    let saved = doc.save_to_vec()?;

    // No critical diffs in the workbook structure.
    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("roundtripped.xlsx");
    std::fs::write(&out_path, &saved)?;
    let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path)?;
    assert!(
        !report.has_at_least(xlsx_diff::Severity::Critical),
        "critical diffs detected: {report:?}"
    );

    // Sheet XML should preserve row/col metadata.
    let xml_bytes = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let xml = std::str::from_utf8(&xml_bytes)?;
    let parsed = roxmltree::Document::parse(xml)?;

    let row2 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("2"))
        .expect("row 2 exists");
    assert_eq!(row2.attribute("ht"), Some("20"));
    assert_eq!(row2.attribute("customHeight"), Some("1"));

    let row3 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("3"))
        .expect("row 3 exists");
    assert_eq!(row3.attribute("hidden"), Some("1"));

    let col_b = parsed
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "col"
                && n.attribute("min") == Some("2")
                && n.attribute("max") == Some("2")
        })
        .expect("col B exists");
    assert_eq!(col_b.attribute("width"), Some("25"));
    assert_eq!(col_b.attribute("customWidth"), Some("1"));

    let col_c = parsed
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "col"
                && n.attribute("min") == Some("3")
                && n.attribute("max") == Some("3")
        })
        .expect("col C exists");
    assert_eq!(col_c.attribute("hidden"), Some("1"));

    let cell_a2 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A2"))
        .expect("cell A2 exists");
    let v = cell_a2
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v, "2");

    Ok(())
}

#[test]
fn editing_a_cell_does_not_strip_unrelated_row_col_or_cell_attrs() -> Result<(), Box<dyn std::error::Error>>
{
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/row-col-attrs.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    let mut doc = load_from_bytes(&fixture_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).unwrap();

    // Edit A2 (which has a `vm="1"` rich-value metadata pointer in the file).
    sheet.set_value(CellRef::from_a1("A2")?, CellValue::Number(99.0));

    let saved = doc.save_to_vec()?;
    let xml_bytes = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let xml = std::str::from_utf8(&xml_bytes)?;
    let parsed = roxmltree::Document::parse(xml)?;

    let row2 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("2"))
        .expect("row 2 exists");
    assert_eq!(row2.attribute("ht"), Some("20"));
    assert_eq!(row2.attribute("customHeight"), Some("1"));
    assert_eq!(row2.attribute("spans"), Some("1:1"));

    let row3 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("3"))
        .expect("row 3 exists");
    assert_eq!(row3.attribute("hidden"), Some("1"));

    let col_b = parsed
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "col"
                && n.attribute("min") == Some("2")
                && n.attribute("max") == Some("2")
        })
        .expect("col B exists");
    assert_eq!(col_b.attribute("width"), Some("25"));
    assert_eq!(col_b.attribute("customWidth"), Some("1"));

    let col_c = parsed
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "col"
                && n.attribute("min") == Some("3")
                && n.attribute("max") == Some("3")
        })
        .expect("col C exists");
    assert_eq!(col_c.attribute("hidden"), Some("1"));

    let cell_a2 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A2"))
        .expect("cell A2 exists");
    assert_eq!(cell_a2.attribute("vm"), Some("1"), "expected vm to be preserved");
    let v = cell_a2
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v, "99");

    Ok(())
}

#[test]
fn editing_cell_style_preserves_vm_attribute_when_value_unchanged(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/row-col-attrs.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    let mut doc = load_from_bytes(&fixture_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;

    // Sanity check the fixture: A2 is a normal number cell that also carries `vm="1"`.
    // Style-only edits should not drop that value-metadata pointer.
    {
        let sheet = doc.workbook.sheet(sheet_id).unwrap();
        assert_eq!(sheet.value(CellRef::from_a1("A2")?), CellValue::Number(2.0));
    }

    let original_sheet_xml = std::str::from_utf8(
        doc.parts()
            .get("xl/worksheets/sheet1.xml")
            .expect("original sheet1.xml exists"),
    )?;
    let original_doc = roxmltree::Document::parse(original_sheet_xml)?;
    let original_a2 = original_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A2"))
        .expect("cell A2 exists in original sheet1.xml");
    assert_eq!(original_a2.attribute("vm"), Some("1"));
    assert_eq!(original_a2.attribute("s"), None, "fixture A2 has no style attr");

    // Apply a new (non-default) style to the cell without changing its cached value.
    let style_id = doc.workbook.styles.intern(Style {
        number_format: Some("0.00".to_string()),
        ..Default::default()
    });
    assert_ne!(style_id, 0, "expected non-default style id");
    {
        let sheet = doc.workbook.sheet_mut(sheet_id).unwrap();
        sheet.set_style_id(CellRef::from_a1("A2")?, style_id);
    }

    let saved = doc.save_to_vec()?;
    let xml_bytes = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let xml = std::str::from_utf8(&xml_bytes)?;
    let parsed = roxmltree::Document::parse(xml)?;

    let cell_a2 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A2"))
        .expect("cell A2 exists");
    assert_eq!(
        cell_a2.attribute("vm"),
        Some("1"),
        "vm should be preserved when only the style changes, got: {xml}"
    );
    assert!(
        cell_a2
            .attribute("s")
            .and_then(|s| s.parse::<u32>().ok())
            .is_some_and(|s| s != 0),
        "expected a non-zero style index to be written for A2, got: {xml}"
    );
    let v = cell_a2
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v, "2");

    Ok(())
}

#[test]
fn new_document_writes_row_and_col_properties() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;
    let sheet = workbook.sheet_mut(sheet_id).unwrap();
    sheet.set_cell(CellRef::from_a1("A1")?, Cell::new(CellValue::Number(1.0)));

    sheet.set_row_height(0, Some(30.0));
    sheet.set_row_hidden(1, true);
    sheet.set_col_width(0, Some(15.0));
    sheet.set_col_hidden(1, true);

    let doc = XlsxDocument::new(workbook);
    let saved = doc.save_to_vec()?;

    let xml_bytes = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let xml = std::str::from_utf8(&xml_bytes)?;
    let parsed = roxmltree::Document::parse(xml)?;

    assert!(
        parsed
            .descendants()
            .any(|n| n.is_element() && n.tag_name().name() == "cols"),
        "expected <cols> to be written when col_properties is non-empty"
    );

    let col_a = parsed
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "col"
                && n.attribute("min") == Some("1")
                && n.attribute("max") == Some("1")
        })
        .expect("col A exists");
    assert_eq!(col_a.attribute("width"), Some("15"));
    assert_eq!(col_a.attribute("customWidth"), Some("1"));

    let col_b = parsed
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "col"
                && n.attribute("min") == Some("2")
                && n.attribute("max") == Some("2")
        })
        .expect("col B exists");
    assert_eq!(col_b.attribute("hidden"), Some("1"));

    let row1 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("1"))
        .expect("row 1 exists");
    assert_eq!(row1.attribute("ht"), Some("30"));
    assert_eq!(row1.attribute("customHeight"), Some("1"));

    let row2 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("2"))
        .expect("row 2 exists");
    assert_eq!(row2.attribute("hidden"), Some("1"));

    Ok(())
}
