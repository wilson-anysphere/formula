use std::io::{Cursor, Read};

use formula_model::rich_text::{RichText, RichTextRunStyle};
use formula_model::{CellRef, CellValue, Workbook};
use zip::ZipArchive;

fn zip_part(zip_bytes: &[u8], name: &str) -> String {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = String::new();
    file.read_to_string(&mut buf).expect("read part");
    buf
}

#[test]
fn semantic_export_preserves_rich_text_in_shared_strings_and_dedupes_by_style(
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");

    let bold = RichText::from_segments(vec![(
        "Hello".to_string(),
        RichTextRunStyle {
            bold: Some(true),
            ..Default::default()
        },
    )]);
    let italic = RichText::from_segments(vec![(
        "Hello".to_string(),
        RichTextRunStyle {
            italic: Some(true),
            ..Default::default()
        },
    )]);

    sheet.set_value(CellRef::from_a1("A1")?, CellValue::RichText(bold));
    sheet.set_value(CellRef::from_a1("A2")?, CellValue::RichText(italic));

    let mut buffer = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buffer)?;
    let bytes = buffer.into_inner();

    let shared_xml = zip_part(&bytes, "xl/sharedStrings.xml");
    assert!(
        shared_xml.contains("<r>"),
        "expected sharedStrings.xml to contain rich text runs (<r> tags)"
    );

    let shared_doc = roxmltree::Document::parse(&shared_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let sst = shared_doc.root_element();
    assert_eq!(sst.attribute("uniqueCount"), Some("2"));
    let si_count = shared_doc
        .descendants()
        .filter(|n| n.has_tag_name((ns, "si")))
        .count();
    assert_eq!(si_count, 2, "expected exactly 2 <si> entries");

    let sheet_xml = zip_part(&bytes, "xl/worksheets/sheet1.xml");
    let sheet_doc = roxmltree::Document::parse(&sheet_xml)?;

    let cell_v = |addr: &str| -> Option<String> {
        let cell = sheet_doc
            .descendants()
            .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some(addr))?;
        cell.children()
            .find(|n| n.has_tag_name((ns, "v")))
            .and_then(|n| n.text())
            .map(|s| s.to_string())
    };

    let a1 = cell_v("A1").expect("expected A1 cell value");
    let a2 = cell_v("A2").expect("expected A2 cell value");
    assert_ne!(
        a1, a2,
        "expected A1 and A2 to reference different shared string indices"
    );

    Ok(())
}

