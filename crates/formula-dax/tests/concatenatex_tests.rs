mod common;

use common::build_model;
use formula_dax::{DataModel, DaxEngine, FilterContext, RowContext, Table, Value};

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

#[test]
fn concatenatex_order_by_text_is_case_insensitive() {
    let mut model = DataModel::new();
    let mut t = Table::new("T", vec!["Name"]);
    t.push_row(vec![Value::from("b")]).unwrap();
    t.push_row(vec![Value::from("a")]).unwrap();
    t.push_row(vec![Value::from("B")]).unwrap();
    model.add_table(t).unwrap();

    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            "CONCATENATEX(T, T[Name], \",\", T[Name], ASC)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    // Case-insensitive sort key with deterministic case-sensitive tiebreak (B before b).
    assert_eq!(value, Value::from("a,B,b"));
}

#[test]
fn concatenatex_order_by_text_can_mix_numeric_and_text_keys() {
    // Our engine allows order-by expressions that produce mixed scalar types. If any key evaluates
    // to text, we sort by text (coercing other types to text).
    let mut model = DataModel::new();
    let mut t = Table::new("T", vec!["Key"]);
    t.push_row(vec![Value::from("A")]).unwrap();
    t.push_row(vec![Value::from(2)]).unwrap();
    model.add_table(t).unwrap();

    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            "CONCATENATEX(T, T[Key], \",\", T[Key], ASC)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, Value::from("2,A"));
}

#[test]
fn concatenatex_supports_table_constructor_value_column() {
    // DAX table constructors expose a single implicit column named [Value]. Ensure CONCATENATEX
    // can iterate those virtual rows and resolve [Value] in the row context.
    let model = DataModel::new();
    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            "CONCATENATEX({\"A\",\"B\"}, [Value], \",\")",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, Value::from("A,B"));
}

#[test]
fn concatenatex_table_constructor_can_order_by_value() {
    let model = DataModel::new();
    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            "CONCATENATEX({\"b\",\"a\"}, [Value], \",\", [Value], ASC)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, Value::from("a,b"));
}
