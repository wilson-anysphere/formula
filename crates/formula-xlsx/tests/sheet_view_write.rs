use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, SheetSelection};
use formula_xlsx::load_from_bytes;
use zip::ZipArchive;

fn build_minimal_xlsx(sheet_xml: &str) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn read_sheet_xml(bytes: &[u8]) -> Result<String, Box<dyn std::error::Error>> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    Ok(sheet_xml)
}

#[test]
fn writes_full_sheet_view_state() -> Result<(), Box<dyn std::error::Error>> {
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
</worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml);
    let mut doc = load_from_bytes(&bytes)?;

    {
        let sheet = &mut doc.workbook.sheets[0];
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
    }

    let out = doc.save_to_vec()?;
    let sheet_xml = read_sheet_xml(&out)?;

    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let sheet_view = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "sheetView")
        .expect("expected sheetView element");

    assert_eq!(
        sheet_view.attribute("zoomScale"),
        Some("125"),
        "expected zoomScale=125, got:\n{sheet_xml}"
    );
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
