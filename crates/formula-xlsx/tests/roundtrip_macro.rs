use std::io::{Cursor, Read};

use formula_xlsx::{assert_xml_semantic_eq, load_from_bytes};
use zip::ZipArchive;

const FIXTURE: &[u8] = include_bytes!("fixtures/rt_macro.xlsm");

#[test]
fn roundtrip_preserves_vba_project_bytes() {
    let doc = load_from_bytes(FIXTURE).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    assert_eq!(
        zip_part(FIXTURE, "xl/vbaProject.bin"),
        zip_part(&saved, "xl/vbaProject.bin"),
        "vbaProject.bin must be preserved byte-for-byte"
    );

    assert_xml_semantic_eq(
        &zip_part(FIXTURE, "xl/workbook.xml"),
        &zip_part(&saved, "xl/workbook.xml"),
    )
    .unwrap();
    assert_xml_semantic_eq(
        &zip_part(FIXTURE, "xl/_rels/workbook.xml.rels"),
        &zip_part(&saved, "xl/_rels/workbook.xml.rels"),
    )
    .unwrap();
}

#[test]
fn roundtrip_preserves_signed_vba_project_bytes() {
    let fixture_path =
        concat!(env!("CARGO_MANIFEST_DIR"), "/../../fixtures/xlsx/macros/signed-basic.xlsm");
    let bytes = std::fs::read(fixture_path).expect("read fixture");
    let doc = load_from_bytes(&bytes).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    assert_eq!(
        zip_part(&bytes, "xl/vbaProject.bin"),
        zip_part(&saved, "xl/vbaProject.bin"),
        "signed vbaProject.bin must be preserved byte-for-byte"
    );
}

#[test]
#[cfg(feature = "vba")]
fn parses_vba_modules_for_ui_display() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/../../fixtures/xlsx/macros/basic.xlsm");
    let bytes = std::fs::read(fixture_path).expect("read fixture");
    let doc = load_from_bytes(&bytes).expect("load fixture");

    let vba_project = doc
        .vba_project()
        .expect("parse vba")
        .expect("vba present");
    assert!(
        vba_project.modules.iter().any(|module| module.name == "Module1"),
        "expected Module1 to be present"
    );
    let module = vba_project
        .modules
        .iter()
        .find(|m| m.name == "Module1")
        .unwrap();
    assert!(module.code.contains("Sub Hello"));
    assert!(module.code.contains("MsgBox"));
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}
