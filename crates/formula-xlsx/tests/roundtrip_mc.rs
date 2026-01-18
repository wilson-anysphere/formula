use std::io::{Cursor, Read};

use formula_xlsx::{assert_xml_semantic_eq, load_from_bytes};
use zip::ZipArchive;

const FIXTURE: &[u8] = include_bytes!("fixtures/rt_mc.xlsx");

#[test]
fn roundtrip_preserves_markup_compatibility_blocks() {
    let doc = load_from_bytes(FIXTURE).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    let original_sheet = zip_part(FIXTURE, "xl/worksheets/sheet1.xml");
    let saved_sheet = zip_part(&saved, "xl/worksheets/sheet1.xml");

    assert_xml_semantic_eq(&original_sheet, &saved_sheet).unwrap();
    assert!(std::str::from_utf8(&saved_sheet)
        .expect("utf8")
        .contains("mc:AlternateContent"));
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

