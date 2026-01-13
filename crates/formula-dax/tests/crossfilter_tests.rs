mod common;

use common::build_model;
use formula_dax::{DaxEngine, FilterContext, RowContext};
use pretty_assertions::assert_eq;

#[test]
fn crossfilter_both_makes_fact_filter_shrink_dimension() {
    let model = build_model();
    let filter = FilterContext::empty().with_column_equals("Orders", "Amount", 20.0.into());

    let without_crossfilter = DaxEngine::new()
        .evaluate(
            &model,
            "COUNTROWS(Customers)",
            &filter,
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(without_crossfilter, 3.into());

    let with_crossfilter = DaxEngine::new()
        .evaluate(
            &model,
            "CALCULATE(COUNTROWS(Customers), CROSSFILTER(Orders[CustomerId], Customers[CustomerId], \"BOTH\"))",
            &filter,
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(with_crossfilter, 1.into());
}

#[test]
fn crossfilter_none_disables_relationship_propagation_for_measure() {
    let model = build_model();
    let customer_filter = FilterContext::empty().with_column_equals("Customers", "CustomerId", 1.into());

    let without_crossfilter = DaxEngine::new()
        .evaluate(
            &model,
            "COUNTROWS(Orders)",
            &customer_filter,
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(without_crossfilter, 2.into());

    let with_crossfilter_none = DaxEngine::new()
        .evaluate(
            &model,
            "CALCULATE(COUNTROWS(Orders), CROSSFILTER(Orders[CustomerId], Customers[CustomerId], \"NONE\"))",
            &customer_filter,
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(with_crossfilter_none, 4.into());
}
