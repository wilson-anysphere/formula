use formula_dax::{
    pivot, Cardinality, CrossFilterDirection, DaxEngine, FilterContext, GroupByColumn, PivotMeasure,
    Relationship, RowContext, Table, Value,
};

fn build_many_to_many_model() -> formula_dax::DataModel {
    let mut model = formula_dax::DataModel::new();

    let mut orders = Table::new("Orders", vec!["OrderId", "ProductId", "Amount"]);
    orders
        .push_row(vec![1.into(), 1.into(), 10.0.into()])
        .unwrap();
    model.add_table(orders).unwrap();

    // Duplicate ProductId rows create a many-to-many relationship between Orders and Products.
    let mut products = Table::new("Products", vec!["ProductId", "Category"]);
    products.push_row(vec![1.into(), "A".into()]).unwrap();
    products.push_row(vec![1.into(), "B".into()]).unwrap();
    model.add_table(products).unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders_Products".into(),
            from_table: "Orders".into(),
            from_column: "ProductId".into(),
            to_table: "Products".into(),
            to_column: "ProductId".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
}

fn build_many_to_many_multi_attr_model() -> formula_dax::DataModel {
    let mut model = formula_dax::DataModel::new();

    let mut orders = Table::new("Orders", vec!["OrderId", "ProductId", "Amount"]);
    orders
        .push_row(vec![1.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![2.into(), 2.into(), 20.0.into()])
        .unwrap();
    model.add_table(orders).unwrap();

    // Duplicate ProductId rows create a many-to-many relationship between Orders and Products.
    // Include multiple attributes to ensure group-by combinations stay correlated per related row.
    let mut products = Table::new("Products", vec!["ProductId", "Category", "Color"]);
    products.push_row(vec![1.into(), "A".into(), "Red".into()]).unwrap();
    products
        .push_row(vec![1.into(), "B".into(), "Blue".into()])
        .unwrap();
    products
        .push_row(vec![2.into(), "C".into(), "Green".into()])
        .unwrap();
    model.add_table(products).unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders_Products".into(),
            from_table: "Orders".into(),
            from_column: "ProductId".into(),
            to_table: "Products".into(),
            to_column: "ProductId".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
}

#[test]
fn summarize_many_to_many_expands_groups() {
    let model = build_many_to_many_model();

    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            "COUNTROWS(SUMMARIZE(Orders, Products[Category]))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, Value::from(2_i64));
}

#[test]
fn pivot_many_to_many_expands_groups_and_duplicates_measures() {
    let model = build_many_to_many_model();

    let measures = vec![PivotMeasure::new("Total Amount", "SUM(Orders[Amount])").unwrap()];
    let group_by = vec![GroupByColumn::new("Products", "Category")];

    let result = pivot(
        &model,
        "Orders",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(result.columns, vec!["Products[Category]", "Total Amount"]);
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("A"), 10.0.into()],
            vec![Value::from("B"), 10.0.into()],
        ]
    );
}

#[test]
fn pivot_many_to_many_respects_dimension_filters() {
    let model = build_many_to_many_model();

    let measures = vec![PivotMeasure::new("Total Amount", "SUM(Orders[Amount])").unwrap()];
    let group_by = vec![GroupByColumn::new("Products", "Category")];
    let filter = FilterContext::empty().with_column_equals("Products", "Category", "A".into());

    let result = pivot(&model, "Orders", &group_by, &measures, &filter).unwrap();
    assert_eq!(result.rows, vec![vec![Value::from("A"), 10.0.into()]]);
}

#[test]
fn summarize_many_to_many_preserves_correlation_for_same_table_columns() {
    let model = build_many_to_many_multi_attr_model();

    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            "COUNTROWS(SUMMARIZE(Orders, Products[Category], Products[Color]))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    // Expect (A,Red), (B,Blue), (C,Green). If group keys were built as {A,B,C}×{Red,Blue,Green},
    // we'd see spurious combinations like (A,Blue).
    assert_eq!(value, Value::from(3_i64));
}

#[test]
fn pivot_many_to_many_preserves_correlation_for_same_table_columns() {
    let model = build_many_to_many_multi_attr_model();

    let measures = vec![PivotMeasure::new("Total Amount", "SUM(Orders[Amount])").unwrap()];
    let group_by = vec![
        GroupByColumn::new("Products", "Category"),
        GroupByColumn::new("Products", "Color"),
    ];

    let result = pivot(
        &model,
        "Orders",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(
        result.columns,
        vec!["Products[Category]", "Products[Color]", "Total Amount"]
    );
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("A"), Value::from("Red"), 10.0.into()],
            vec![Value::from("B"), Value::from("Blue"), 10.0.into()],
            vec![Value::from("C"), Value::from("Green"), 20.0.into()],
        ]
    );
}

#[test]
fn summarizecolumns_many_to_many_expands_and_preserves_correlation() {
    let model = build_many_to_many_multi_attr_model();
    let engine = DaxEngine::new();

    let value = engine
        .evaluate(
            &model,
            "COUNTROWS(SUMMARIZECOLUMNS(Orders[OrderId], Products[Category], Products[Color]))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(3_i64));
}
