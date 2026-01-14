use formula_dax::{DataModel, DaxEngine, FilterContext, RowContext, Value};
use pretty_assertions::assert_eq;

#[test]
fn logical_operators_short_circuit_and_or() {
    let model = DataModel::new();
    let engine = DaxEngine::new();

    // The RHS would raise a type error if evaluated.
    let value = engine
        .evaluate(
            &model,
            "FALSE() && (1 = \"a\")",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(false));

    let value = engine
        .evaluate(
            &model,
            "TRUE() || (1 = \"a\")",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(true));
}

