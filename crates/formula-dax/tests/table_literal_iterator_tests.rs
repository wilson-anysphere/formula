mod common;

use common::build_model;
use formula_dax::{DaxEngine, FilterContext, RowContext, Value};
use pretty_assertions::assert_eq;

#[test]
fn sumx_table_constructor_value_column_sums() {
    let model = build_model();
    let engine = DaxEngine::new();

    let result = engine
        .evaluate(
            &model,
            "SUMX({1,2,3}, [Value])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(result, Value::from(6.0));
}

#[test]
fn sumx_table_constructor_measure_does_not_error_and_repeats_measure() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let engine = DaxEngine::new();

    // The table constructor creates a virtual row context with a synthetic [Value] column that is
    // not part of the model. Context transition should ignore that binding (it cannot produce
    // filters), so the measure evaluates in the outer filter context for each row.
    let result = engine
        .evaluate(
            &model,
            "SUMX({1,2,3}, [Total Sales])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(result, Value::from(129.0));
}

