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

