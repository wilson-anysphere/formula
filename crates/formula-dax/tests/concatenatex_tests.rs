mod common;

use common::build_model;
use formula_dax::{DaxEngine, FilterContext, RowContext, Value};

#[test]
fn concatenatex_values_returns_all_distinct_regions() {
    let model = build_model();
    let engine = DaxEngine::new();

    let value = engine
        .evaluate(
            &model,
            "CONCATENATEX(VALUES(Customers[Region]), Customers[Region], \",\")",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    let out = match value {
        Value::Text(out) => out,
        other => panic!("expected text result, got {other:?}"),
    };

    let parts: Vec<&str> = out.as_ref().split(',').collect();
    assert_eq!(parts.len(), 2, "expected two distinct regions");
    assert!(parts.contains(&"East"));
    assert!(parts.contains(&"West"));
}

#[test]
fn concatenatex_respects_filter_context() {
    let model = build_model();
    let engine = DaxEngine::new();

    let filter = FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    let value = engine
        .evaluate(
            &model,
            "CONCATENATEX(VALUES(Customers[Region]), Customers[Region], \",\")",
            &filter,
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, Value::from("East"));
}

#[test]
fn concatenatex_orders_by_expression_desc() {
    let model = build_model();
    let engine = DaxEngine::new();

    let value = engine
        .evaluate(
            &model,
            "CONCATENATEX(Customers, Customers[Name], \",\", Customers[Name], DESC)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, Value::from("Carol,Bob,Alice"));
}
