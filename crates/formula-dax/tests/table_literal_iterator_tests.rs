use formula_dax::{DataModel, DaxEngine, FilterContext, RowContext, Value};
use pretty_assertions::assert_eq;

#[test]
fn countrows_over_table_literal_counts_rows() {
    let model = DataModel::new();
    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            "COUNTROWS({1,2,3})",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, Value::from(3i64));
}

#[test]
fn sumx_over_table_literal_uses_value_column() {
    let model = DataModel::new();
    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            "SUMX({1,2,3}, [Value])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, Value::from(6.0));
}

#[test]
fn filter_over_table_literal_can_reference_value_column() {
    let model = DataModel::new();
    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            "COUNTROWS(FILTER({1,2,3}, [Value] > 1))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, Value::from(2i64));
}

