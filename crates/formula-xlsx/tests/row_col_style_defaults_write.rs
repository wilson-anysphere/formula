use std::io::{Cursor, Read};

use formula_model::{Style, Workbook};
use formula_xlsx::XlsxDocument;
use zip::ZipArchive;

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(zip_bytes)).expect("open zip");
    let mut buf = Vec::new();
    archive
        .by_name(name)
        .unwrap_or_else(|_| panic!("missing zip part: {name}"))
        .read_to_end(&mut buf)
        .expect("read zip part");
    buf
}

#[test]
fn new_document_writes_row_and_col_default_styles() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;

    let style_id = workbook.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Default::default()
    });
    assert_ne!(style_id, 0, "expected non-default style id");

    {
        let sheet = workbook.sheet_mut(sheet_id).unwrap();
        // Row/col properties are 0-based. These apply to row 2 and column B.
        sheet.set_row_style_id(1, Some(style_id));
        sheet.set_col_style_id(1, Some(style_id));
    }

    let doc = XlsxDocument::new(workbook);
    let saved = doc.save_to_vec()?;

    let sheet_xml_bytes = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let sheet_xml = std::str::from_utf8(&sheet_xml_bytes)?;
    let parsed = roxmltree::Document::parse(sheet_xml)?;

    let row2 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("2"))
        .expect("expected row 2 to be written");
    assert!(
        row2.attribute("s")
            .and_then(|v| v.parse::<u32>().ok())
            .is_some_and(|xf| xf != 0),
        "expected row 2 to have a non-zero style xf index, got: {sheet_xml}"
    );
    assert_eq!(
        row2.attribute("customFormat"),
        Some("1"),
        "expected row 2 to set customFormat when a default row style is present"
    );

    let col_b = parsed
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "col"
                && n.attribute("min") == Some("2")
                && n.attribute("max") == Some("2")
        })
        .expect("expected column B to be written");
    assert!(
        col_b
            .attribute("style")
            .and_then(|v| v.parse::<u32>().ok())
            .is_some_and(|xf| xf != 0),
        "expected column B to have a non-zero style xf index, got: {sheet_xml}"
    );
    assert_eq!(
        col_b.attribute("customFormat"),
        Some("1"),
        "expected column B to set customFormat when a default col style is present"
    );

    Ok(())
}

