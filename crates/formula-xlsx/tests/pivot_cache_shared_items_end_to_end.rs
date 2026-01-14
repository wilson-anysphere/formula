use std::io::{Cursor, Write};

use formula_engine::pivot::{AggregationType, PivotFieldRef, PivotTable, PivotValue};
use formula_xlsx::pivots::engine_bridge::{pivot_cache_to_engine_source, pivot_table_to_engine_config};
use formula_xlsx::{PivotCacheValue, XlsxPackage};

use pretty_assertions::assert_eq;

#[test]
fn pivot_cache_shared_item_indices_flow_through_engine_bridge() {
    let bytes = build_synthetic_pivot_package();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read xlsx");

    let cache_def = pkg
        .pivot_cache_definition("xl/pivotCache/pivotCacheDefinition1.xml")
        .expect("parse cache def")
        .expect("cache definition exists");

    let mut cache_records = pkg
        .pivot_cache_records("xl/pivotCache/pivotCacheRecords1.xml")
        .expect("open cache records");
    let records = cache_records.parse_all_records();

    assert_eq!(
        records,
        vec![
            vec![
                PivotCacheValue::Index(0),
                PivotCacheValue::Index(0),
                PivotCacheValue::Number(100.0)
            ],
            vec![
                PivotCacheValue::Index(0),
                PivotCacheValue::Index(1),
                PivotCacheValue::Number(150.0)
            ],
            vec![
                PivotCacheValue::Index(1),
                PivotCacheValue::Index(0),
                PivotCacheValue::Number(200.0)
            ],
            vec![
                PivotCacheValue::Index(1),
                PivotCacheValue::Index(1),
                PivotCacheValue::Number(250.0)
            ],
        ]
    );

    let source = pivot_cache_to_engine_source(&cache_def, records.into_iter());
    assert_eq!(
        source[0],
        vec![
            PivotValue::Text("Region".to_string()),
            PivotValue::Text("Product".to_string()),
            PivotValue::Text("Sales".to_string()),
        ]
    );
    assert_eq!(
        source[1],
        vec![
            PivotValue::Text("East".to_string()),
            PivotValue::Text("A".to_string()),
            PivotValue::Number(100.0),
        ]
    );

    let table = pkg
        .pivot_table_definition("xl/pivotTables/pivotTable1.xml")
        .expect("parse pivot table");
    let cfg = pivot_table_to_engine_config(&table, &cache_def);

    assert_eq!(cfg.row_fields.len(), 1);
    assert_eq!(
        cfg.row_fields[0].source_field,
        PivotFieldRef::CacheFieldName("Region".to_string())
    );
    assert_eq!(cfg.value_fields.len(), 1);
    assert_eq!(
        cfg.value_fields[0].source_field,
        PivotFieldRef::CacheFieldName("Sales".to_string())
    );
    assert_eq!(cfg.value_fields[0].aggregation, AggregationType::Sum);

    let pivot = PivotTable::new("PivotTable1", &source, cfg).expect("create pivot");
    let result = pivot.calculate().expect("calculate");

    assert_eq!(
        result.data,
        vec![
            vec![
                PivotValue::Text("Region".to_string()),
                PivotValue::Text("Sum of Sales".to_string())
            ],
            vec![PivotValue::Text("East".to_string()), PivotValue::Number(250.0)],
            vec![PivotValue::Text("West".to_string()), PivotValue::Number(450.0)],
            vec![
                PivotValue::Text("Grand Total".to_string()),
                PivotValue::Number(700.0)
            ],
        ]
    );
}

fn build_synthetic_pivot_package() -> Vec<u8> {
    let cache_definition_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="1" recordCount="4">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:C5" sheet="Sheet1"/>
  </cacheSource>
  <cacheFields count="3">
    <cacheField name="Region">
      <sharedItems count="2">
        <s v="East"/>
        <s v="West"/>
      </sharedItems>
    </cacheField>
    <cacheField name="Product">
      <sharedItems count="2">
        <s v="A"/>
        <s v="B"/>
      </sharedItems>
    </cacheField>
    <cacheField name="Sales">
      <sharedItems count="0"/>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>"#;

    let cache_records_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4">
  <r><x v="0"/><x v="0"/><n v="100"/></r>
  <r><x v="0"/><x v="1"/><n v="150"/></r>
  <r><x v="1"/><x v="0"/><n v="200"/></r>
  <r><x v="1"/><x v="1"/><n v="250"/></r>
</pivotCacheRecords>"#;

    let pivot_table_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1" cacheId="1">
  <rowFields count="1">
    <field x="0"/>
  </rowFields>
  <dataFields count="1">
    <dataField fld="2" subtotal="sum"/>
  </dataFields>
</pivotTableDefinition>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    for (name, xml) in [
        (
            "xl/pivotCache/pivotCacheDefinition1.xml",
            cache_definition_xml,
        ),
        ("xl/pivotCache/pivotCacheRecords1.xml", cache_records_xml),
        ("xl/pivotTables/pivotTable1.xml", pivot_table_xml),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(xml.as_bytes()).unwrap();
    }

    zip.finish().unwrap().into_inner()
}
