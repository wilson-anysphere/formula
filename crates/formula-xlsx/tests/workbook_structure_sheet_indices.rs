use std::io::{Cursor, Read, Write};

use formula_xlsx::load_from_bytes;
use pretty_assertions::assert_eq;
use quick_xml::events::Event;
use quick_xml::Reader;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

fn build_fixture_xlsx() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews>
    <workbookView activeTab="1" firstSheet="0"/>
  </bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="LocalName" localSheetId="1">Sheet2!A1</definedName>
  </definedNames>
</workbook>
"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>
"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options)
        .expect("zip file");
    zip.write_all(content_types.as_bytes()).expect("zip write");

    zip.start_file("_rels/.rels", options).expect("zip file");
    zip.write_all(root_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/workbook.xml", options)
        .expect("zip file");
    zip.write_all(workbook_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("zip file");
    zip.write_all(workbook_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("zip file");
    zip.write_all(sheet_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/worksheets/sheet2.xml", options)
        .expect("zip file");
    zip.write_all(sheet_xml.as_bytes()).expect("zip write");

    zip.finish().expect("finish zip").into_inner()
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn workbook_view_indices(xml: &[u8]) -> (Option<usize>, Option<usize>) {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut active_tab = None;
    let mut first_sheet = None;
    loop {
        match reader.read_event_into(&mut buf).expect("read xml") {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"workbookView" => {
                for attr in e.attributes().flatten() {
                    let val = attr.unescape_value().expect("attr").into_owned();
                    match attr.key.as_ref() {
                        b"activeTab" => active_tab = val.parse::<usize>().ok(),
                        b"firstSheet" => first_sheet = val.parse::<usize>().ok(),
                        _ => {}
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    (active_tab, first_sheet)
}

fn defined_name_local_sheet_ids(xml: &[u8]) -> Vec<(String, Option<usize>)> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).expect("read xml") {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"definedName" => {
                let mut name = None;
                let mut local = None;
                for attr in e.attributes().flatten() {
                    let val = attr.unescape_value().expect("attr").into_owned();
                    match attr.key.as_ref() {
                        b"name" => name = Some(val),
                        b"localSheetId" => local = val.parse::<usize>().ok(),
                        _ => {}
                    }
                }
                if let Some(name) = name {
                    out.push((name, local));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

#[test]
fn reorder_updates_local_sheet_ids_and_workbook_view_indices() {
    let fixture = build_fixture_xlsx();
    let mut doc = load_from_bytes(&fixture).expect("load fixture");

    let sheet2_id = doc.workbook.sheets[1].id;
    assert!(doc.workbook.reorder_sheet(sheet2_id, 0));

    let saved = doc.save_to_vec().expect("save");
    let workbook_xml = zip_part(&saved, "xl/workbook.xml");

    assert_eq!(workbook_view_indices(&workbook_xml), (Some(0), Some(1)));
    assert_eq!(
        defined_name_local_sheet_ids(&workbook_xml),
        vec![("LocalName".to_string(), Some(0))]
    );
}

#[test]
fn delete_drops_defined_names_scoped_to_removed_sheet() {
    let fixture = build_fixture_xlsx();
    let mut doc = load_from_bytes(&fixture).expect("load fixture");

    let sheet2_id = doc.workbook.sheets[1].id;
    doc.workbook.delete_sheet(sheet2_id).expect("delete");

    let saved = doc.save_to_vec().expect("save");
    let workbook_xml = zip_part(&saved, "xl/workbook.xml");

    assert_eq!(workbook_view_indices(&workbook_xml), (Some(0), Some(0)));
    assert_eq!(defined_name_local_sheet_ids(&workbook_xml), Vec::new());
}
