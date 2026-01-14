mod common;

use common::build_model;
use formula_dax::{DataModel, DaxEngine, FilterContext, RowContext, Value};
use pretty_assertions::assert_eq;

#[test]
fn containsrow_with_one_column_table_literal() {
    let model = DataModel::new();
    let engine = DaxEngine::new();

    let value = engine
        .evaluate(
            &model,
            "CONTAINSROW({1,2,3}, 2)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(true));

    let value = engine
        .evaluate(
            &model,
            "CONTAINSROW({1,2,3}, 4)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(false));
}

#[test]
fn containsrow_with_table_literal_var() {
    let model = DataModel::new();
    let engine = DaxEngine::new();

    let value = engine
        .evaluate(
            &model,
            "VAR t = {1,2,3} RETURN CONTAINSROW(t, 2)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(true));
}

#[test]
fn containsrow_with_values_column() {
    let model = build_model();
    let engine = DaxEngine::new();

    let value = engine
        .evaluate(
            &model,
            "CONTAINSROW(VALUES(Customers[Region]), \"East\")",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(true));
}

#[test]
fn containsrow_with_table_literal_and_column_ref() {
    let model = build_model();
    let engine = DaxEngine::new();

    // Common membership pattern: `CONTAINSROW({"A","B"}, Table[Col])` inside a row iterator.
    let value = engine
        .evaluate(
            &model,
            "COUNTROWS(FILTER(Customers, CONTAINSROW({\"East\"}, Customers[Region])))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(2));
}

#[test]
fn containsrow_with_multi_column_table_expression() {
    let model = build_model();
    let engine = DaxEngine::new();

    // Virtual two-column table (SUMMARIZE).
    let value = engine
        .evaluate(
            &model,
            "CONTAINSROW(SUMMARIZE(Orders, Orders[OrderId], Orders[CustomerId]), 100, 1)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(true));

    let value = engine
        .evaluate(
            &model,
            "CONTAINSROW(SUMMARIZE(Orders, Orders[OrderId], Orders[CustomerId]), 100, 2)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(false));

    // Physical table: match all columns in order.
    let value = engine
        .evaluate(
            &model,
            "CONTAINSROW(Orders, 100, 1, 10)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(true));
}
