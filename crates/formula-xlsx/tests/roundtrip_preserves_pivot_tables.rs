use std::io::{Cursor, Read};

use formula_xlsx::load_from_bytes;
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

