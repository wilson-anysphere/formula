use formula_engine::pivot::{PivotKeyPart, SortOrder};
use formula_xlsx::pivots::engine_bridge::pivot_table_to_engine_config;
use formula_xlsx::{PivotCacheDefinition, PivotCacheField, PivotCacheValue, PivotTableDefinition};

use pretty_assertions::assert_eq;

#[test]
fn maps_indexed_manual_sort_items_via_shared_items() {
    let table_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pivotFields count="1">
    <pivotField axis="axisRow" sortType="manual">
      <items count="2">
        <item x="1"/>
        <item x="0"/>
      </items>
    </pivotField>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
</pivotTableDefinition>"#;

    let table = PivotTableDefinition::parse("xl/pivotTables/pivotTable1.xml", table_xml)
        .expect("parse pivot table definition");

    let cache_def = PivotCacheDefinition {
        cache_fields: vec![PivotCacheField {
            name: "Region".to_string(),
            shared_items: Some(vec![
                PivotCacheValue::String("East".to_string()),
                PivotCacheValue::String("West".to_string()),
            ]),
            ..Default::default()
        }],
        ..Default::default()
    };

    let cfg = pivot_table_to_engine_config(&table, &cache_def);
    assert_eq!(cfg.row_fields.len(), 1);
    assert_eq!(cfg.row_fields[0].sort_order, SortOrder::Manual);
    assert_eq!(
        cfg.row_fields[0].manual_sort.as_deref(),
        Some(&[
            PivotKeyPart::Text("West".to_string()),
            PivotKeyPart::Text("East".to_string()),
        ][..])
    );
}

