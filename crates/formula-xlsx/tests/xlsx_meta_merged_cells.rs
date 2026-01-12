use std::io::{Cursor, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
use zip::ZipWriter;

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
    let mut zip = ZipWriter::new(cursor);
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

#[test]
fn cell_meta_resolves_merged_cells_to_anchor() -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1">
      <c r="A1" vm="1" cm="2"><v>42</v></c>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A1:B2"/>
  </mergeCells>
</worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);
    let mut doc = load_from_bytes(&bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;

    let anchor = CellRef::from_a1("A1")?;
    let interior = CellRef::from_a1("B2")?;

    // `cell_meta` is a convenience helper: callers working against the model's
    // merged-cell semantics should be able to query any cell in the merged
    // region and see the anchor metadata.
    let meta = doc
        .cell_meta(sheet_id, interior)
        .expect("expected cell metadata for merged region");
    assert_eq!(meta.vm.as_deref(), Some("1"));
    assert_eq!(meta.cm.as_deref(), Some("2"));

    // Editing a non-anchor cell in a merged region edits the anchor cell. Ensure
    // the metadata map is updated for the anchor and does not create a separate
    // entry for the non-anchor cell.
    doc.set_cell_value(sheet_id, interior, CellValue::String("Hello".to_string()));
    assert!(
        doc.xlsx_meta().cell_meta.contains_key(&(sheet_id, anchor)),
        "expected metadata to remain keyed by the merged-region anchor cell"
    );
    assert!(
        !doc.xlsx_meta().cell_meta.contains_key(&(sheet_id, interior)),
        "expected no metadata entry for non-anchor merged cell"
    );

    Ok(())
}

