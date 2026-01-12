use std::io::{Cursor, Write};

use formula_model::Range;
use formula_xlsx::{load_from_bytes, read_workbook_model_from_bytes, worksheet_parts_from_reader};
use zip::write::FileOptions;
use zip::ZipArchive;
use zip::ZipWriter;

fn build_minimal_xlsx_with_leading_slash_entries() -> Vec<u8> {
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
  <mergeCells count="1">
    <mergeCell ref="A1:B2"/>
  </mergeCells>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options =
        FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    fn add_file(
        zip: &mut ZipWriter<Cursor<Vec<u8>>>,
        options: FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(&mut zip, options, "/xl/workbook.xml", workbook_xml);
    add_file(
        &mut zip,
        options,
        "/xl/_rels/workbook.xml.rels",
        workbook_rels,
    );
    add_file(&mut zip, options, "/xl/worksheets/sheet1.xml", worksheet_xml);

    zip.finish().unwrap().into_inner()
}

#[test]
fn worksheet_parts_from_reader_tolerates_leading_slash_entries() {
    let bytes = build_minimal_xlsx_with_leading_slash_entries();
    let parts = worksheet_parts_from_reader(Cursor::new(bytes)).expect("worksheet parts");
    assert_eq!(parts.len(), 1);
    assert_eq!(parts[0].name, "Sheet1");
    assert_eq!(parts[0].worksheet_part, "xl/worksheets/sheet1.xml");
}

#[test]
fn read_workbook_model_from_bytes_tolerates_leading_slash_entries() {
    let bytes = build_minimal_xlsx_with_leading_slash_entries();
    let workbook = read_workbook_model_from_bytes(&bytes).expect("read workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
}

#[test]
fn load_from_bytes_tolerates_leading_slash_entries() {
    let bytes = build_minimal_xlsx_with_leading_slash_entries();
    let doc = load_from_bytes(&bytes).expect("load xlsx document");
    assert_eq!(doc.workbook.sheets.len(), 1);
    assert_eq!(doc.workbook.sheets[0].name, "Sheet1");
}

#[test]
fn merge_cells_reader_tolerates_leading_slash_entries() {
    let bytes = build_minimal_xlsx_with_leading_slash_entries();
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("zip");
    let merges =
        formula_xlsx::merge_cells::read_merge_cells_from_xlsx(&mut archive, "xl/worksheets/sheet1.xml")
            .expect("merge cells");
    assert_eq!(merges, vec![Range::from_a1("A1:B2").unwrap()]);
}

