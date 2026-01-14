use std::io::{Cursor, Seek, SeekFrom, Write};

use formula_xlsx::{load_from_bytes, load_from_reader};
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_minimal_xlsx() -> Vec<u8> {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml).unwrap();
    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels).unwrap();
    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn load_from_reader_seeks_to_start() {
    let bytes = build_minimal_xlsx();

    let doc_from_bytes = load_from_bytes(&bytes).expect("load_from_bytes");

    let mut cursor = Cursor::new(bytes.as_slice());
    cursor.seek(SeekFrom::End(0)).unwrap();
    let doc_from_reader = load_from_reader(cursor).expect("load_from_reader");

    assert_eq!(doc_from_reader.workbook.sheets.len(), doc_from_bytes.workbook.sheets.len());
    assert_eq!(doc_from_reader.workbook.sheets[0].name, doc_from_bytes.workbook.sheets[0].name);
}

