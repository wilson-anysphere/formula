use std::io::{Cursor, Write};

use chrono::NaiveDate;
use formula_engine::pivot::{PivotCache, PivotTable, PivotValue};
use formula_xlsx::pivots::engine_bridge::{
    apply_pivot_slicer_parts_to_engine_config, pivot_cache_to_engine_source, pivot_table_to_engine_config,
};
use formula_xlsx::{PivotCacheRecordsReader, PivotTableDefinition, XlsxPackage};
use pretty_assertions::assert_eq;
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_pkg(entries: &[(&str, &[u8])]) -> XlsxPackage {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }
    let bytes = zip.finish().unwrap().into_inner();
    XlsxPackage::from_bytes(&bytes).expect("parse test pkg")
}

#[test]
fn timeline_parts_date_range_filters_engine_pivot_output() {
    let timeline_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          name="Timeline1">
  <timelineCache r:id="rIdCache1"/>
</timeline>"#;

    let timeline_rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdCache1" Target="../timelineCaches/timelineCacheDefinition1.xml"/>
</Relationships>"#;

    // Connect the timeline cache definition to a pivot table part and include a date range selection.
    let timeline_cache_definition_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                         name="Cache1"
                         sourceName="OrderDate">
  <pivotTables>
    <pivotTable r:id="rIdPivot1"/>
  </pivotTables>
  <selection startDate="2024-01-02" endDate="2024-01-02"/>
</timelineCacheDefinition>"#;

    let timeline_cache_definition_rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPivot1" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#;

    let cache_def_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheFields count="2">
    <cacheField name="OrderDate"/>
    <cacheField name="Sales"/>
  </cacheFields>
</pivotCacheDefinition>"#;

    let cache_records_xml = br#"
        <pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <r><d v="2024-01-01T00:00:00Z"/><n v="100"/></r>
          <r><d v="2024-01-02T00:00:00Z"/><n v="200"/></r>
          <r><d v="2024-01-03T00:00:00Z"/><n v="300"/></r>
        </pivotCacheRecords>
    "#;

    let pivot_table_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1"
  cacheId="1">
  <rowFields count="1">
    <field x="0"/>
  </rowFields>
  <dataFields count="1">
    <dataField fld="1" name="Sum of Sales" subtotal="sum"/>
  </dataFields>
</pivotTableDefinition>"#;

    let pkg = build_pkg(&[
        ("xl/timelines/timeline1.xml", timeline_xml),
        ("xl/timelines/_rels/timeline1.xml.rels", timeline_rels_xml),
        (
            "xl/timelineCaches/timelineCacheDefinition1.xml",
            timeline_cache_definition_xml,
        ),
        (
            "xl/timelineCaches/_rels/timelineCacheDefinition1.xml.rels",
            timeline_cache_definition_rels_xml,
        ),
        ("xl/pivotCache/pivotCacheDefinition1.xml", cache_def_xml),
        ("xl/pivotCache/pivotCacheRecords1.xml", cache_records_xml),
        ("xl/pivotTables/pivotTable1.xml", pivot_table_xml),
    ]);

    let parts = pkg.pivot_slicer_parts().expect("parse timelines");
    assert_eq!(parts.timelines.len(), 1);
    assert_eq!(
        parts.timelines[0].connected_pivot_tables,
        vec!["xl/pivotTables/pivotTable1.xml"]
    );
    assert_eq!(parts.timelines[0].source_name.as_deref(), Some("OrderDate"));
    assert_eq!(parts.timelines[0].selection.start.as_deref(), Some("2024-01-02"));
    assert_eq!(parts.timelines[0].selection.end.as_deref(), Some("2024-01-02"));

    let cache_def = pkg
        .pivot_cache_definition("xl/pivotCache/pivotCacheDefinition1.xml")
        .expect("parse cache def")
        .expect("cache def exists");

    let mut reader = PivotCacheRecordsReader::new(cache_records_xml);
    let records = reader.parse_all_records();
    let source = pivot_cache_to_engine_source(&cache_def, records.into_iter());
    let cache = PivotCache::from_range(&source).expect("engine cache");

    let table = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", pivot_table_xml)
        .expect("parse pivot table def");
    let cfg = pivot_table_to_engine_config(&table, &cache_def);

    let pivot_all = PivotTable::new("PivotTable1", &source, cfg.clone()).expect("pivot");
    let result_all = pivot_all.calculate().expect("calculate");

    let d1 = NaiveDate::from_ymd_opt(2024, 1, 1).expect("valid date");
    let d2 = NaiveDate::from_ymd_opt(2024, 1, 2).expect("valid date");
    let d3 = NaiveDate::from_ymd_opt(2024, 1, 3).expect("valid date");

    assert_eq!(
        result_all.data,
        vec![
            vec![
                PivotValue::Text("OrderDate".to_string()),
                PivotValue::Text("Sum of Sales".to_string())
            ],
            vec![PivotValue::Date(d1), PivotValue::Number(100.0)],
            vec![PivotValue::Date(d2), PivotValue::Number(200.0)],
            vec![PivotValue::Date(d3), PivotValue::Number(300.0)],
            vec![PivotValue::Text("Grand Total".to_string()), PivotValue::Number(600.0)],
        ]
    );

    let mut cfg_filtered = cfg;
    apply_pivot_slicer_parts_to_engine_config(
        &mut cfg_filtered,
        "xl/pivotTables/pivotTable1.xml",
        &cache_def,
        &cache,
        &parts,
    );

    let pivot_filtered = PivotTable::new("PivotTable1", &source, cfg_filtered).expect("pivot");
    let result_filtered = pivot_filtered.calculate().expect("calculate");

    assert_eq!(
        result_filtered.data,
        vec![
            vec![
                PivotValue::Text("OrderDate".to_string()),
                PivotValue::Text("Sum of Sales".to_string())
            ],
            vec![PivotValue::Date(d2), PivotValue::Number(200.0)],
            vec![PivotValue::Text("Grand Total".to_string()), PivotValue::Number(200.0)],
        ]
    );
}

