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
