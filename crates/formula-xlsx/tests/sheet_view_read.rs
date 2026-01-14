use std::io::{Cursor, Write};

use formula_model::{CellRef, Range};
use formula_xlsx::{load_from_bytes, read_workbook_model_from_bytes};

fn build_minimal_xlsx(workbook_xml: &str, sheet_xml: &str) -> Vec<u8> {
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options =
        zip::write::FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn assert_sheet_view(sheet: &formula_model::Worksheet) {
    assert_eq!(sheet.zoom, 1.25);
    assert_eq!(sheet.view.zoom, 1.25);

    assert_eq!(sheet.frozen_cols, 2);
    assert_eq!(sheet.frozen_rows, 3);
    assert_eq!(sheet.view.pane.frozen_cols, 2);
    assert_eq!(sheet.view.pane.frozen_rows, 3);
    assert_eq!(sheet.view.pane.top_left_cell, Some(CellRef::from_a1("C4").unwrap()));

    assert!(!sheet.view.show_grid_lines);
    assert!(!sheet.view.show_headings);
    assert!(!sheet.view.show_zeros);

    let selection = sheet.view.selection.as_ref().expect("expected selection");
    assert_eq!(selection.active_cell, CellRef::from_a1("D5").unwrap());
    assert_eq!(selection.ranges, vec![Range::from_a1("D5:E6").unwrap()]);
}

#[test]
fn reads_sheet_views_into_worksheet_view() {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
 <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
   <sheetViews>
     <sheetView zoomScale="125" showGridLines="0" showHeadings="0" showZeros="0">
      <pane state="frozen" xSplit=" 2 " ySplit=" 3 " topLeftCell=" C4 "/>
      <selection activeCell=" D5 " sqref=" D5:E6 "/>
     </sheetView>
   </sheetViews>
   <sheetData/>
 </worksheet>"#;

    let bytes = build_minimal_xlsx(workbook_xml, sheet_xml);

    let full = load_from_bytes(&bytes).expect("load_from_bytes");
    assert_sheet_view(&full.workbook.sheets[0]);

    let fast = read_workbook_model_from_bytes(&bytes).expect("read_workbook_model_from_bytes");
    assert_sheet_view(&fast.sheets[0]);
}
