mod common;

use common::build_model;
use formula_dax::{DaxEngine, DaxError, FilterContext, RowContext, Value};
use pretty_assertions::assert_eq;

#[test]
fn summarize_virtual_row_context_behaves_like_dax_in_iterators() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let engine = DaxEngine::new();

    let total = engine
        .evaluate(
            &model,
            "SUMX(SUMMARIZE(Orders, Customers[Region]), [Total Sales])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(total, Value::from(43.0));

    let groups = engine
        .evaluate(
            &model,
            "COUNTROWS(SUMMARIZE(Orders, Customers[Region]))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(groups, 2.into());

    let err = engine
        .evaluate(
            &model,
            "SUMX(SUMMARIZE(Orders, Customers[Region]), Orders[Amount])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap_err();
    assert!(
        matches!(err, DaxError::Eval(_)),
        "expected row-context error, got: {err:?}"
    );
}
