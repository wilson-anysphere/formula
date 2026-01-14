use formula_dax::{DataModel, DaxEngine, FilterContext, RowContext, Value};
use pretty_assertions::assert_eq;

#[test]
fn number_literals_support_exponent_notation() {
    let model = DataModel::new();
    let engine = DaxEngine::new();

    let v = engine
        .evaluate(
            &model,
            "1e3",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(v, Value::from(1000.0));

    let v = engine
        .evaluate(
            &model,
            "1e-3",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(v, Value::from(0.001));

    let v = engine
        .evaluate(
            &model,
            "1.5E2",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(v, Value::from(150.0));
}

