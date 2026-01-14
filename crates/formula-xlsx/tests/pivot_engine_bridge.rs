use formula_engine::pivot::{AggregationType, PivotFieldRef, PivotTable, PivotValue};
use formula_xlsx::pivots::engine_bridge::{
    pivot_cache_to_engine_source, pivot_table_to_engine_config,
};
use formula_xlsx::XlsxPackage;

use pretty_assertions::assert_eq;

const FIXTURE: &[u8] = include_bytes!("fixtures/pivot-fixture.xlsx");

#[test]
fn converts_pivot_cache_and_table_to_engine_types() {
    let pkg = XlsxPackage::from_bytes(FIXTURE).expect("read xlsx");

    let cache_def = pkg
        .pivot_cache_definition("xl/pivotCache/pivotCacheDefinition1.xml")
        .expect("parse cache def")
        .expect("cache definition exists");
    let mut cache_records = pkg
        .pivot_cache_records("xl/pivotCache/pivotCacheRecords1.xml")
        .expect("open cache records");
    let records = cache_records.parse_all_records();

    let pivots = pkg.pivots().expect("parse pivots");
    let table = pivots
        .pivot_tables
        .first()
        .expect("fixture includes pivot table");

    let source = pivot_cache_to_engine_source(&cache_def, records.into_iter());

    assert_eq!(
        source[0],
        vec![
            PivotValue::Text("Region".to_string()),
            PivotValue::Text("Product".to_string()),
            PivotValue::Text("Sales".to_string()),
        ]
    );
    assert_eq!(source.len(), 5, "header + 4 record rows");

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
            vec![
                PivotValue::Text("East".to_string()),
                PivotValue::Number(250.0)
            ],
            vec![
                PivotValue::Text("West".to_string()),
                PivotValue::Number(450.0)
            ],
            vec![
                PivotValue::Text("Grand Total".to_string()),
                PivotValue::Number(700.0)
            ],
        ]
    );
}
