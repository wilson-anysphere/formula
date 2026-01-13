use std::io::{Cursor, Read};

use formula_model::{Style, Workbook};
use formula_xlsx::write_workbook_to_writer;
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
fn exports_row_and_col_default_style_ids() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;

    let style_id = workbook.styles.intern(Style {
        number_format: Some("0.00".to_string()),
        ..Default::default()
    });
    assert_ne!(style_id, 0, "expected a non-default style id");

    {
        let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
        sheet.set_row_style_id(0, Some(style_id));
        sheet.set_col_style_id(0, Some(style_id));
    }

    let mut cursor = Cursor::new(Vec::new());
    write_workbook_to_writer(&workbook, &mut cursor)?;
    let bytes = cursor.into_inner();

    let sheet_xml = zip_part(&bytes, "xl/worksheets/sheet1.xml");
    let sheet_xml_str = std::str::from_utf8(&sheet_xml)?;

    assert!(
        sheet_xml_str.contains(r#"customFormat="1""#),
        "expected row/col styles to emit customFormat=1, got: {sheet_xml_str}"
    );
    assert!(
        sheet_xml_str.contains(r#" s=""#),
        "expected row default styles to emit row/@s, got: {sheet_xml_str}"
    );
    assert!(
        sheet_xml_str.contains(r#" style=""#),
        "expected col default styles to emit col/@style, got: {sheet_xml_str}"
    );

    let parsed = roxmltree::Document::parse(sheet_xml_str)?;

    let row1 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("1"))
        .expect("expected <row r=\"1\"> to be written for row style defaults");
    assert_eq!(row1.attribute("customFormat"), Some("1"));
    assert!(
        row1.attribute("s")
            .and_then(|s| s.parse::<u32>().ok())
            .is_some(),
        "expected row/@s to be an integer style index, got: {sheet_xml_str}"
    );

    let col_a = parsed
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "col"
                && n.attribute("min") == Some("1")
                && n.attribute("max") == Some("1")
        })
        .expect("expected <col min=\"1\" max=\"1\"> to be written for col style defaults");
    assert_eq!(col_a.attribute("customFormat"), Some("1"));
    assert!(
        col_a.attribute("style")
            .and_then(|s| s.parse::<u32>().ok())
            .is_some(),
        "expected col/@style to be an integer style index, got: {sheet_xml_str}"
    );

    Ok(())
}

