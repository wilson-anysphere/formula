use formula_dax::{
    pivot, Cardinality, CrossFilterDirection, DataModel, FilterContext, GroupByColumn, PivotMeasure,
    Relationship, Table, Value,
};

#[test]
fn isinscope_uses_pivot_scope_metadata() {
    let mut model = DataModel::new();

    let mut customers = Table::new("Customers", vec!["CustomerId", "Region"]);
    customers
        .push_row(vec![Value::from(1), Value::from("East")])
        .unwrap();
    customers
        .push_row(vec![Value::from(2), Value::from("West")])
        .unwrap();
    model.add_table(customers).unwrap();

    let mut sales = Table::new("Sales", vec!["SaleId", "CustomerId", "Amount"]);
    sales
        .push_row(vec![Value::from(10), Value::from(1), Value::from(100.0)])
        .unwrap();
    sales
        .push_row(vec![Value::from(11), Value::from(2), Value::from(50.0)])
        .unwrap();
    model.add_table(sales).unwrap();

    model
        .add_relationship(Relationship {
            name: "Sales_Customers".into(),
            from_table: "Sales".into(),
            from_column: "CustomerId".into(),
            to_table: "Customers".into(),
            to_column: "CustomerId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        // Identifiers are case-insensitive; verify that scope metadata and identifier resolution do
        // not depend on exact casing.
        .add_measure("InScopeRegion", "ISINSCOPE(customers[region])")
        .unwrap();

    assert_eq!(
        model
            .evaluate_measure("InScopeRegion", &FilterContext::empty())
            .unwrap(),
        Value::Boolean(false)
    );

    assert_eq!(
        model
            .evaluate_measure(
                "InScopeRegion",
                &FilterContext::empty().with_column_equals("Customers", "Region", "East".into()),
            )
            .unwrap(),
        Value::Boolean(false)
    );

    let group_by = vec![GroupByColumn::new("Customers", "Region")];
    let measures = vec![PivotMeasure::new("InScopeRegion", "[InScopeRegion]").unwrap()];

    let result = pivot(
        &model,
        "Sales",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(result.rows.len(), 2);
    for row in &result.rows {
        assert_eq!(row.len(), 2);
        assert_eq!(row[1], Value::Boolean(true));
    }
}

#[test]
fn isinscope_uses_pivot_scope_metadata_for_unicode_identifiers() {
    let mut model = DataModel::new();

    let mut streets = Table::new("Straße", vec!["StraßenId", "Region"]);
    streets
        .push_row(vec![Value::from(1), Value::from("East")])
        .unwrap();
    streets
        .push_row(vec![Value::from(2), Value::from("West")])
        .unwrap();
    model.add_table(streets).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "StraßenId", "Amount"]);
    orders
        .push_row(vec![Value::from(10), Value::from(1), Value::from(100.0)])
        .unwrap();
    orders
        .push_row(vec![Value::from(11), Value::from(2), Value::from(50.0)])
        .unwrap();
    model.add_table(orders).unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders_Straße".into(),
            from_table: "Orders".into(),
            from_column: "straßenid".into(),
            to_table: "STRASSE".into(),
            to_column: "STRASSENID".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_measure("InScopeRegion", "ISINSCOPE(STRASSE[region])")
        .unwrap();

    assert_eq!(
        model
            .evaluate_measure("InScopeRegion", &FilterContext::empty())
            .unwrap(),
        Value::Boolean(false)
    );

    assert_eq!(
        model
            .evaluate_measure(
                "InScopeRegion",
                &FilterContext::empty().with_column_equals("STRASSE", "REGION", "East".into()),
            )
            .unwrap(),
        Value::Boolean(false)
    );

    let group_by = vec![GroupByColumn::new("STRASSE", "region")];
    let measures = vec![PivotMeasure::new("InScopeRegion", "[InScopeRegion]").unwrap()];

    let result = pivot(
        &model,
        "Orders",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(result.rows.len(), 2);
    for row in &result.rows {
        assert_eq!(row.len(), 2);
        assert_eq!(row[1], Value::Boolean(true));
    }
}
