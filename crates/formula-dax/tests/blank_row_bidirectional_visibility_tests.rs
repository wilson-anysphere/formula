use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, Relationship, RowContext,
    Table, Value,
};
use pretty_assertions::assert_eq;

fn build_model_blank_row_bidirectional() -> DataModel {
    let mut model = DataModel::new();

    let mut customers = Table::new("Customers", vec!["CustomerId", "Region"]);
    customers.push_row(vec![1.into(), "East".into()]).unwrap();
    customers.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(customers).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "CustomerId"]);
    orders.push_row(vec![100.into(), 1.into()]).unwrap();
    orders.push_row(vec![101.into(), 999.into()]).unwrap();
    model.add_table(orders).unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders_Customers".into(),
            from_table: "Orders".into(),
            from_column: "CustomerId".into(),
            to_table: "Customers".into(),
            to_column: "CustomerId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Both,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
}

#[test]
fn values_virtual_blank_row_respects_bidirectional_filter_context() {
    let model = build_model_blank_row_bidirectional();
    let engine = DaxEngine::new();

    // A) No extra filters: East, West, BLANK.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(VALUES(Customers[Region]))",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        3.into()
    );

    // B) Filter fact to a matched key: BLANK should disappear.
    let matched_filter = FilterContext::empty().with_column_equals("Orders", "CustomerId", 1.into());
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(VALUES(Customers[Region]))",
                &matched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "SELECTEDVALUE(Customers[Region])",
                &matched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        Value::from("East")
    );

    // C) Filter fact to an unmatched key: only BLANK should remain.
    let unmatched_filter =
        FilterContext::empty().with_column_equals("Orders", "CustomerId", 999.into());
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(VALUES(Customers[Region]))",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "SELECTEDVALUE(Customers[Region])",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        Value::Blank
    );
}

#[test]
fn distinctcount_virtual_blank_row_respects_bidirectional_filter_context() {
    let model = build_model_blank_row_bidirectional();
    let engine = DaxEngine::new();

    // D) Mirror the same visibility behavior for DISTINCTCOUNT.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "DISTINCTCOUNT(Customers[Region])",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        3.into()
    );

    let matched_filter = FilterContext::empty().with_column_equals("Orders", "CustomerId", 1.into());
    assert_eq!(
        engine
            .evaluate(
                &model,
                "DISTINCTCOUNT(Customers[Region])",
                &matched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );

    let unmatched_filter =
        FilterContext::empty().with_column_equals("Orders", "CustomerId", 999.into());
    assert_eq!(
        engine
            .evaluate(
                &model,
                "DISTINCTCOUNT(Customers[Region])",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

