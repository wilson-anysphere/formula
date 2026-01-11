use std::io::{Cursor, Read};

use formula_xlsx::load_from_bytes;
use formula_xlsx::XlsxPackage;
use zip::ZipArchive;

const FIXTURE: &[u8] = include_bytes!("fixtures/pivot-fixture.xlsx");

#[test]
fn roundtrip_preserves_pivot_table_and_cache_parts() {
    let doc = load_from_bytes(FIXTURE).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    for part in [
        "xl/pivotTables/pivotTable1.xml",
        "xl/pivotCache/pivotCacheDefinition1.xml",
        "xl/pivotCache/pivotCacheRecords1.xml",
    ] {
        assert_eq!(
            zip_part(FIXTURE, part),
            zip_part(&saved, part),
            "pivot part {part} changed during round-trip"
        );
    }
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

#[test]
fn parses_pivot_table_definition_layout() {
    let package = XlsxPackage::from_bytes(FIXTURE).expect("read package");
    let pivot = package
        .pivot_table_definition("xl/pivotTables/pivotTable1.xml")
        .expect("parse pivot table definition");

    assert_eq!(pivot.row_fields, vec![0]);
    assert_eq!(pivot.data_fields.len(), 1);

    let data_field = &pivot.data_fields[0];
    assert_eq!(data_field.fld, Some(2));
    assert_eq!(data_field.name.as_deref(), Some("Sum of Sales"));
    assert_eq!(data_field.subtotal.as_deref(), Some("sum"));
}
