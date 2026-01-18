use std::io::{Cursor, Read};

use formula_xlsx::{assert_xml_semantic_eq, load_from_bytes};
use zip::ZipArchive;

const FIXTURE: &[u8] = include_bytes!("fixtures/rich_text_shared_strings.xlsx");

#[test]
fn roundtrip_preserves_rich_text_shared_strings() {
    let doc = load_from_bytes(FIXTURE).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    assert_xml_semantic_eq(
        &zip_part(FIXTURE, "xl/sharedStrings.xml"),
        &zip_part(&saved, "xl/sharedStrings.xml"),
    )
    .unwrap();
    assert_xml_semantic_eq(
        &zip_part(FIXTURE, "xl/worksheets/sheet1.xml"),
        &zip_part(&saved, "xl/worksheets/sheet1.xml"),
    )
    .unwrap();
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

