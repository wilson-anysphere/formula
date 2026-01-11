use std::io::{Cursor, Read};

use formula_vba::parse_vba_digital_signature;

#[test]
fn unsigned_fixture_reports_no_signature() {
    let fixture_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/macros/basic.xlsm"
    );
    let bytes = std::fs::read(fixture_path).expect("fixture xlsm exists");

    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("valid xlsm zip");
    let mut file = archive
        .by_name("xl/vbaProject.bin")
        .expect("vbaProject.bin present");
    let mut vba = Vec::new();
    file.read_to_end(&mut vba).expect("read vbaProject.bin");

    let sig = parse_vba_digital_signature(&vba).expect("signature parse should succeed");
    assert!(sig.is_none(), "fixture should be unsigned");
}

