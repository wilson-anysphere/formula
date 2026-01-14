use chrono::NaiveDate;
use formula_xlsx::pivots::cache_records::{
    pivot_cache_datetime_to_naive_date, PivotCacheRecordsReader, PivotCacheValue,
};
use formula_xlsx::XlsxPackage;

use pretty_assertions::assert_eq;

const FIXTURE: &[u8] = include_bytes!("fixtures/pivot-cache-datetime.xlsx");

#[test]
fn parses_datetime_values_in_pivot_cache_records() {
    let pkg = XlsxPackage::from_bytes(FIXTURE).expect("read fixture");
    let xml = pkg
        .part("xl/pivotCache/pivotCacheRecords1.xml")
        .expect("pivot cache records part present");

    let mut reader = PivotCacheRecordsReader::new(xml);
    let records = reader.parse_all_records();

    assert_eq!(
        records[0][0],
        PivotCacheValue::DateTime("2024-01-15T00:00:00Z".to_string())
    );
}

#[test]
fn converts_datetime_string_to_naive_date() {
    assert_eq!(
        pivot_cache_datetime_to_naive_date("2024-01-15T00:00:00Z"),
        NaiveDate::from_ymd_opt(2024, 1, 15)
    );
}

#[test]
fn converts_iso_date_prefix_with_trailing_chars_to_naive_date() {
    assert_eq!(
        pivot_cache_datetime_to_naive_date("2024-01-15Z"),
        NaiveDate::from_ymd_opt(2024, 1, 15)
    );
}

#[test]
fn converts_compact_ymd_to_naive_date() {
    assert_eq!(
        pivot_cache_datetime_to_naive_date("20240115"),
        NaiveDate::from_ymd_opt(2024, 1, 15)
    );
}

#[test]
fn converts_excel_serial_date_to_naive_date() {
    assert_eq!(
        pivot_cache_datetime_to_naive_date("1"),
        NaiveDate::from_ymd_opt(1900, 1, 1)
    );
}
