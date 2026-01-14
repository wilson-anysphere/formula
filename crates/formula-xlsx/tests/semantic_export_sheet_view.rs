use std::io::{Cursor, Read};

use formula_model::{CellRef, SheetSelection, Workbook};
use formula_xlsx::write_workbook_to_writer;
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
fn semantic_export_emits_sheet_views_for_non_default_view() -> Result<(), Box<dyn std::error::Error>>
{
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1".to_string())?;
    let sheet = workbook.sheet_mut(sheet_id).unwrap();

    sheet.view.zoom = 1.25;
    sheet.view.show_grid_lines = false;
    sheet.view.show_headings = false;
    sheet.view.show_zeros = false;
    sheet.view.pane.frozen_rows = 2;
    sheet.view.pane.frozen_cols = 1;
    sheet.view.pane.top_left_cell = Some(CellRef::from_a1("B3")?);
    sheet.view.selection = Some(SheetSelection::from_sqref(
        CellRef::from_a1("D5")?,
        "D5:E6",
    )?);

    let mut cursor = Cursor::new(Vec::new());
    write_workbook_to_writer(&workbook, &mut cursor)?;
    let bytes = cursor.into_inner();

    let sheet_xml = zip_part(&bytes, "xl/worksheets/sheet1.xml");
    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let sheet_view = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "sheetView")
        .expect("expected sheetView element");

    assert_eq!(sheet_view.attribute("zoomScale"), Some("125"));
    assert_eq!(sheet_view.attribute("showGridLines"), Some("0"));
    assert_eq!(sheet_view.attribute("showRowColHeaders"), Some("0"));
    assert_eq!(sheet_view.attribute("showZeros"), Some("0"));

    let pane = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "pane")
        .expect("expected pane element");
    assert_eq!(pane.attribute("state"), Some("frozen"));
    assert_eq!(pane.attribute("xSplit"), Some("1"));
    assert_eq!(pane.attribute("ySplit"), Some("2"));
    assert_eq!(pane.attribute("topLeftCell"), Some("B3"));

    let selection = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "selection")
        .expect("expected selection element");
    assert_eq!(selection.attribute("activeCell"), Some("D5"));
    assert_eq!(selection.attribute("sqref"), Some("D5:E6"));

    Ok(())
}

