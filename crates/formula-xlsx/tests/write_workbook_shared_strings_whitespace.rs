use std::io::{Cursor, Read};

use formula_model::{CellRef, CellValue, Workbook};
use zip::ZipArchive;

const SS_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

#[test]
fn write_workbook_emits_xml_space_preserve_for_strings_with_outer_whitespace(
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
    sheet.set_value(
        CellRef::from_a1("A1")?,
        CellValue::String("  hello  ".to_string()),
    );
    sheet.set_value(
        CellRef::from_a1("A2")?,
        CellValue::String("hello".to_string()),
    );

    let mut buffer = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buffer)?;

    let xml_bytes = zip_part(&buffer.into_inner(), "xl/sharedStrings.xml");
    let xml = std::str::from_utf8(&xml_bytes)?;
    let doc = roxmltree::Document::parse(xml)?;

    let preserved = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().namespace() == Some(SS_NS)
                && n.tag_name().name() == "t"
                && n.text() == Some("  hello  ")
        })
        .expect("expected shared string for \"  hello  \"");
    assert_eq!(preserved.attribute((XML_NS, "space")), Some("preserve"));

    let normal = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().namespace() == Some(SS_NS)
                && n.tag_name().name() == "t"
                && n.text() == Some("hello")
        })
        .expect("expected shared string for \"hello\"");
    assert_eq!(normal.attribute((XML_NS, "space")), None);

    Ok(())
}

