use std::io::{Cursor, Read, Seek};

use formula_model::{Cell, CellRef, CellValue, Workbook};
use zip::ZipArchive;

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn col_covers(n: roxmltree::Node<'_, '_>, idx: u32) -> bool {
    if !n.is_element() || n.tag_name().name() != "col" {
        return false;
    }
    let min = n.attribute("min").and_then(|v| v.parse::<u32>().ok());
    let max = n.attribute("max").and_then(|v| v.parse::<u32>().ok());
    match (min, max) {
        (Some(min), Some(max)) => min <= idx && idx <= max,
        _ => false,
    }
}

#[test]
#[cfg(not(target_arch = "wasm32"))]
fn write_workbook_emits_outline_metadata() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Outline")?;
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
    sheet.set_cell(CellRef::from_a1("A1")?, Cell::new(CellValue::Number(1.0)));

    // Rows 2-3 are detail rows at level 1; row 4 is the collapsed summary row.
    sheet.outline.rows.entry_mut(2).level = 1;
    sheet.outline.rows.entry_mut(3).level = 1;
    sheet.outline.rows.entry_mut(4).collapsed = true;

    // Columns B-C are detail columns at level 1; column D is the collapsed summary col.
    sheet.outline.cols.entry_mut(2).level = 1;
    sheet.outline.cols.entry_mut(3).level = 1;
    sheet.outline.cols.entry_mut(4).collapsed = true;

    sheet.outline.recompute_outline_hidden_rows();
    sheet.outline.recompute_outline_hidden_cols();

    let mut cursor = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut cursor)?;
    cursor.rewind()?;
    let bytes = cursor.into_inner();

    let xml_bytes = zip_part(&bytes, "xl/worksheets/sheet1.xml");
    let xml = std::str::from_utf8(&xml_bytes)?;
    let parsed = roxmltree::Document::parse(xml)?;

    let outline_pr = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "outlinePr")
        .expect("outlinePr should be written for outlined sheets");
    assert_eq!(outline_pr.attribute("summaryBelow"), Some("1"));
    assert_eq!(outline_pr.attribute("summaryRight"), Some("1"));
    assert_eq!(outline_pr.attribute("showOutlineSymbols"), Some("1"));

    let row2 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("2"))
        .expect("row 2 exists");
    assert_eq!(row2.attribute("outlineLevel"), Some("1"));
    assert_eq!(row2.attribute("hidden"), Some("1"));

    let row4 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("4"))
        .expect("row 4 exists");
    assert_eq!(row4.attribute("collapsed"), Some("1"));

    let col_b = parsed
        .descendants()
        .find(|n| col_covers(*n, 2))
        .expect("col B exists");
    assert_eq!(col_b.attribute("outlineLevel"), Some("1"));
    assert_eq!(col_b.attribute("hidden"), Some("1"));

    let col_d = parsed
        .descendants()
        .find(|n| col_covers(*n, 4))
        .expect("col D exists");
    assert_eq!(col_d.attribute("collapsed"), Some("1"));

    Ok(())
}
