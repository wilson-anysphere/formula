use formula_xlsx::{load_from_bytes, PivotCacheValue, XlsxPackage};

use pretty_assertions::assert_eq;

const FIXTURE: &[u8] = include_bytes!("fixtures/pivot-fixture.xlsx");

#[test]
fn parses_pivot_cache_records_fixture() {
    let pkg = XlsxPackage::from_bytes(FIXTURE).expect("read pkg");

    let pivots = pkg.pivots().expect("parse pivots");
    assert_eq!(pivots.pivot_cache_records.len(), 1);
    assert_eq!(pivots.pivot_cache_records[0].count, Some(4));

    let mut reader = pkg
        .pivot_cache_records(&pivots.pivot_cache_records[0].path)
        .expect("pivotCacheRecords part exists");
    let records = reader.parse_all_records();

    assert_eq!(records.len(), 4);
    assert_eq!(
        records[0],
        vec![
            PivotCacheValue::String("East".to_string()),
            PivotCacheValue::String("A".to_string()),
            PivotCacheValue::Number(100.0),
        ]
    );
}

#[test]
fn parses_pivot_cache_records_fixture_from_document() {
    let doc = load_from_bytes(FIXTURE).expect("load fixture");
    let mut reader = doc
        .pivot_cache_records("xl/pivotCache/pivotCacheRecords1.xml")
        .expect("pivotCacheRecords part exists");
    let records = reader.parse_all_records();

    assert_eq!(records.len(), 4);
    assert_eq!(
        records[0],
        vec![
            PivotCacheValue::String("East".to_string()),
            PivotCacheValue::String("A".to_string()),
            PivotCacheValue::Number(100.0),
        ]
    );
}
