use std::collections::HashSet;
use std::io::{Cursor, Write};

use formula_engine::pivot::{PivotCache, PivotTable, PivotValue};
use formula_xlsx::pivots::engine_bridge::{
    apply_pivot_slicer_parts_to_engine_config, pivot_cache_to_engine_source,
    pivot_table_to_engine_config,
};
use formula_xlsx::{PivotCacheRecordsReader, PivotCacheValue, PivotTableDefinition, XlsxPackage};
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
fn slicer_cache_x_indices_roundtrip_into_engine_filters() {
    // Minimal slicer that stores selection via shared-item indices (`x="0"` etc).
    let slicer_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        name="Slicer1">
  <slicerCache r:id="rIdCache1"/>
</slicer>"#;

    let slicer_rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdCache1" Target="../slicerCaches/slicerCache1.xml"/>
</Relationships>"#;

    // Connect the slicer cache to a pivot table part and include selection items using `x`.
    let slicer_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            name="Cache1" sourceName="Region">
  <slicerCachePivotTables>
    <slicerCachePivotTable r:id="rIdPivot1"/>
  </slicerCachePivotTables>
  <slicerCacheItems>
    <slicerCacheItem x="0" s="1"/>
    <slicerCacheItem x="1" s="0"/>
  </slicerCacheItems>
</slicerCache>"#;

    let slicer_cache_rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPivot1" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#;

    // Minimal pivot cache definition: shared-items for Region, inline numbers for Sales.
    let cache_def_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheFields count="2">
    <cacheField name="Region">
      <sharedItems count="2">
        <s v="East"/>
        <s v="West"/>
      </sharedItems>
    </cacheField>
    <cacheField name="Sales"/>
  </cacheFields>
</pivotCacheDefinition>"#;

    // Records store Region using `<x v="..."/>` indices.
    let cache_records_xml = br#"
        <pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <r><x v="0"/><n v="100"/></r>
          <r><x v="0"/><n v="150"/></r>
          <r><x v="1"/><n v="200"/></r>
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
        ("xl/slicers/slicer1.xml", slicer_xml),
        ("xl/slicers/_rels/slicer1.xml.rels", slicer_rels_xml),
        ("xl/slicerCaches/slicerCache1.xml", slicer_cache_xml),
        (
            "xl/slicerCaches/_rels/slicerCache1.xml.rels",
            slicer_cache_rels_xml,
        ),
        ("xl/pivotCache/pivotCacheDefinition1.xml", cache_def_xml),
        ("xl/pivotCache/pivotCacheRecords1.xml", cache_records_xml),
        ("xl/pivotTables/pivotTable1.xml", pivot_table_xml),
    ]);

    let parts = pkg.pivot_slicer_parts().expect("parse slicers");
    assert_eq!(parts.slicers.len(), 1);
    assert_eq!(
        parts.slicers[0].connected_pivot_tables,
        vec!["xl/pivotTables/pivotTable1.xml"]
    );
    assert_eq!(parts.slicers[0].source_name.as_deref(), Some("Region"));
    assert_eq!(parts.slicers[0].selection.available_items, vec!["0", "1"]);
    assert_eq!(
        parts.slicers[0]
            .selection
            .selected_items
            .as_ref()
            .map(|set| set.clone()),
        Some(HashSet::from(["0".to_string()])),
    );

    let cache_def = pkg
        .pivot_cache_definition("xl/pivotCache/pivotCacheDefinition1.xml")
        .expect("parse cache def")
        .expect("cache def exists");

    let mut reader = PivotCacheRecordsReader::new(cache_records_xml);
    let records = reader.parse_all_records();
    assert!(
        matches!(
            records.get(0).and_then(|row| row.get(0)),
            Some(PivotCacheValue::Index(0))
        ),
        "expected record to use shared-item index"
    );

    let source = pivot_cache_to_engine_source(&cache_def, records.into_iter());
    let cache = PivotCache::from_range(&source).expect("engine cache");

    let table = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", pivot_table_xml)
        .expect("parse pivot table def");
    let cfg = pivot_table_to_engine_config(&table, &cache_def);

    let pivot_all = PivotTable::new("PivotTable1", &source, cfg.clone()).expect("pivot");
    let result_all = pivot_all.calculate().expect("calculate");
    assert_eq!(
        result_all.data,
        vec![
            vec![
                PivotValue::Text("Region".to_string()),
                PivotValue::Text("Sum of Sales".to_string())
            ],
            vec![
                PivotValue::Text("East".to_string()),
                PivotValue::Number(250.0)
            ],
            vec![
                PivotValue::Text("West".to_string()),
                PivotValue::Number(200.0)
            ],
            vec![
                PivotValue::Text("Grand Total".to_string()),
                PivotValue::Number(450.0)
            ],
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
                PivotValue::Text("Region".to_string()),
                PivotValue::Text("Sum of Sales".to_string())
            ],
            vec![
                PivotValue::Text("East".to_string()),
                PivotValue::Number(250.0)
            ],
            vec![
                PivotValue::Text("Grand Total".to_string()),
                PivotValue::Number(250.0)
            ],
        ]
    );
}
