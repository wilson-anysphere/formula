use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, Range};
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
fn worksheet_view_roundtrips_and_preserves_unknown_sheetview_children(
) -> Result<(), Box<dyn std::error::Error>> {
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0" zoomScale="125" showGridLines="0" showRowColHeaders="0" showZeros="0">
      <pane state="frozen" xSplit="1" ySplit="2" topLeftCell="B3"/>
      <selection activeCell="C4" sqref="C4 D5:E6"/>
      <extLst>
        <ext uri="{12345678-1234-1234-1234-1234567890AB}">
          <x:dummy xmlns:x="http://example.com/dummy"/>
        </ext>
      </extLst>
    </sheetView>
  </sheetViews>
  <sheetData/>
</worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml);
    let doc = load_from_bytes(&bytes)?;
    let sheet = &doc.workbook.sheets[0];

    assert!(!sheet.view.show_grid_lines);
    assert!(!sheet.view.show_headings);
    assert!(!sheet.view.show_zeros);
    assert_eq!(sheet.view.zoom, 1.25);
    assert_eq!(sheet.view.pane.frozen_cols, 1);
    assert_eq!(sheet.view.pane.frozen_rows, 2);

    let selection = sheet
        .view
        .selection
        .as_ref()
        .expect("selection should be parsed");
    assert_eq!(selection.active_cell, CellRef::from_a1("C4")?);
    assert_eq!(selection.sqref(), "C4 D5:E6");
    assert_eq!(
        selection.ranges,
        vec![Range::from_a1("C4")?, Range::from_a1("D5:E6")?]
    );

    // Legacy fields should mirror the full view model.
    assert_eq!(sheet.zoom, sheet.view.zoom);
    assert_eq!(sheet.frozen_rows, sheet.view.pane.frozen_rows);
    assert_eq!(sheet.frozen_cols, sheet.view.pane.frozen_cols);

    // No-op save should not drop selection/extLst (unknown child content).
    let out = doc.save_to_vec()?;
    let out_xml = read_sheet_xml(&out)?;
    assert!(
        out_xml.contains(r#"<ext uri="{12345678-1234-1234-1234-1234567890AB}">"#)
            && out_xml.contains("<extLst")
            && out_xml.contains("<x:dummy"),
        "expected extLst subtree to be preserved, got:\n{out_xml}"
    );
    assert!(
        out_xml.contains(r#"<selection activeCell="C4" sqref="C4 D5:E6"/>"#),
        "expected selection payload to be preserved, got:\n{out_xml}"
    );

    Ok(())
}
