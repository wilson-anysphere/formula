use std::collections::BTreeMap;
use std::io::{Cursor, Read};

use formula_xlsx::{assert_xml_semantic_eq, load_from_bytes};
use quick_xml::events::Event;
use quick_xml::Reader;
use zip::ZipArchive;

const FIXTURE: &[u8] = include_bytes!("fixtures/rt_simple.xlsx");

#[test]
fn roundtrip_simple_preserves_key_parts() {
    let doc = load_from_bytes(FIXTURE).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    // XML parts we actively rewrite should remain semantically identical.
    assert_xml_semantic_eq(
        &zip_part(FIXTURE, "xl/workbook.xml"),
        &zip_part(&saved, "xl/workbook.xml"),
    )
    .unwrap();
    assert_xml_semantic_eq(
        &zip_part(FIXTURE, "xl/worksheets/sheet1.xml"),
        &zip_part(&saved, "xl/worksheets/sheet1.xml"),
    )
    .unwrap();
    assert_xml_semantic_eq(
        &zip_part(FIXTURE, "xl/sharedStrings.xml"),
        &zip_part(&saved, "xl/sharedStrings.xml"),
    )
    .unwrap();
    assert_xml_semantic_eq(
        &zip_part(FIXTURE, "xl/styles.xml"),
        &zip_part(&saved, "xl/styles.xml"),
    )
    .unwrap();

    // Parts we don't model should be preserved byte-for-byte.
    assert_eq!(
        zip_part(FIXTURE, "xl/theme/theme1.xml"),
        zip_part(&saved, "xl/theme/theme1.xml"),
        "theme xml should be preserved verbatim"
    );
    assert_eq!(
        zip_part(FIXTURE, "xl/calcChain.xml"),
        zip_part(&saved, "xl/calcChain.xml"),
        "calcChain.xml should be preserved verbatim"
    );

    // Relationship IDs must not change.
    assert_eq!(
        workbook_sheet_rids(&zip_part(FIXTURE, "xl/workbook.xml")),
        workbook_sheet_rids(&zip_part(&saved, "xl/workbook.xml")),
    );
    assert_eq!(
        relationships_map(&zip_part(FIXTURE, "xl/_rels/workbook.xml.rels")),
        relationships_map(&zip_part(&saved, "xl/_rels/workbook.xml.rels")),
    );
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn workbook_sheet_rids(xml: &[u8]) -> Vec<(String, String, String)> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).expect("read xml") {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"sheet" => {
                let mut name = String::new();
                let mut sheet_id = String::new();
                let mut rid = String::new();
                for attr in e.attributes().flatten() {
                    let key = attr.key.as_ref();
                    let val = attr.unescape_value().expect("attr").into_owned();
                    match key {
                        b"name" => name = val,
                        b"sheetId" => sheet_id = val,
                        b"r:id" => rid = val,
                        _ => {}
                    }
                }
                out.push((name, sheet_id, rid));
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn relationships_map(xml: &[u8]) -> BTreeMap<String, String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut map = BTreeMap::new();
    loop {
        match reader.read_event_into(&mut buf).expect("read xml") {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                let mut id = None;
                let mut target = None;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Id" => id = Some(attr.unescape_value().unwrap().into_owned()),
                        b"Target" => target = Some(attr.unescape_value().unwrap().into_owned()),
                        _ => {}
                    }
                }
                if let (Some(id), Some(target)) = (id, target) {
                    map.insert(id, target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    map
}
