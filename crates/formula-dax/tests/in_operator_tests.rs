mod common;

use common::build_model;
use formula_dax::{DaxEngine, FilterContext, RowContext, Value};
use pretty_assertions::assert_eq;

#[test]
fn in_operator_scalar_table_constructor() {
    let model = build_model();
    let engine = DaxEngine::new();
    let filter = FilterContext::empty();
    let row_ctx = RowContext::default();

    let value = engine
        .evaluate(&model, "1 IN {1,2}", &filter, &row_ctx)
        .unwrap();
    assert_eq!(value, Value::from(true));

    let value = engine
        .evaluate(&model, "3 IN {1,2}", &filter, &row_ctx)
        .unwrap();
    assert_eq!(value, Value::from(false));
}

#[test]
fn in_operator_calculate_column_filter() {
    let model = build_model();
    let engine = DaxEngine::new();
    let filter = FilterContext::empty();
    let row_ctx = RowContext::default();

    let value = engine
        .evaluate(
            &model,
            "CALCULATE(COUNTROWS(Orders), Orders[CustomerId] IN {1,3})",
            &filter,
            &row_ctx,
        )
        .unwrap();
    assert_eq!(value, Value::from(3i64));
}

#[test]
fn in_operator_scalar_table_expression_rhs() {
    let model = build_model();
    let engine = DaxEngine::new();
    let filter = FilterContext::empty();
    let row_ctx = RowContext::default();

    // Physical one-column table expressions like VALUES(column) should work.
    let value = engine
        .evaluate(&model, "1 IN VALUES(Orders[CustomerId])", &filter, &row_ctx)
        .unwrap();
    assert_eq!(value, Value::from(true));

    let value = engine
        .evaluate(&model, "4 IN VALUES(Orders[CustomerId])", &filter, &row_ctx)
        .unwrap();
    assert_eq!(value, Value::from(false));

    // Virtual one-column tables should also work (e.g. SUMMARIZE).
    let value = engine
        .evaluate(&model, "1 IN SUMMARIZE(Orders, Orders[CustomerId])", &filter, &row_ctx)
        .unwrap();
    assert_eq!(value, Value::from(true));

    let value = engine
        .evaluate(&model, "4 IN SUMMARIZE(Orders, Orders[CustomerId])", &filter, &row_ctx)
        .unwrap();
    assert_eq!(value, Value::from(false));
}

#[test]
fn in_operator_table_expression_requires_one_column() {
    let model = build_model();
    let engine = DaxEngine::new();
    let filter = FilterContext::empty();
    let row_ctx = RowContext::default();

    let err = engine
        .evaluate(&model, "1 IN Orders", &filter, &row_ctx)
        .unwrap_err();
    let message = err.to_string();
    assert!(message.contains("one-column"));
}
