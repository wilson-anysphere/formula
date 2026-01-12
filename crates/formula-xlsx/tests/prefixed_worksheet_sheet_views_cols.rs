use std::io::{Cursor, Read, Write};

use formula_model::ColProperties;
use formula_xlsx::{load_from_bytes, XlsxDocument};
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

fn sheet_mut(doc: &mut XlsxDocument) -> &mut formula_model::Worksheet {
    &mut doc.workbook.sheets[0]
}

#[test]
fn writer_inserts_prefixed_sheet_views_and_cols() -> Result<(), Box<dyn std::error::Error>> {
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetData/>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml);
    let mut doc = load_from_bytes(&bytes)?;

    {
        let sheet = sheet_mut(&mut doc);
        sheet.zoom = 1.25;
        sheet.frozen_rows = 2;
        sheet.frozen_cols = 1;
        sheet.col_properties.insert(
            0,
            ColProperties {
                width: Some(12.0),
                hidden: false,
                style_id: None,
            },
        );
    }

    let out = doc.save_to_vec()?;
    let sheet_xml = read_sheet_xml(&out)?;

    roxmltree::Document::parse(&sheet_xml)?;

    assert!(
        sheet_xml.contains("<x:sheetViews") && sheet_xml.contains("</x:sheetViews>"),
        "expected prefixed <x:sheetViews> insertion, got:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<x:sheetView") && sheet_xml.contains("</x:sheetView>"),
        "expected prefixed <x:sheetView> insertion, got:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<x:pane"),
        "expected prefixed <x:pane> insertion, got:\n{sheet_xml}"
    );

    assert!(
        sheet_xml.contains("<x:cols>") && sheet_xml.contains("</x:cols>"),
        "expected prefixed <x:cols> insertion, got:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<x:col "),
        "expected prefixed <x:col> insertion, got:\n{sheet_xml}"
    );

    assert!(
        !sheet_xml.contains("<sheetViews"),
        "should not introduce an unprefixed <sheetViews> element"
    );
    assert!(
        !sheet_xml.contains("<cols>"),
        "should not introduce an unprefixed <cols> element"
    );

    Ok(())
}
