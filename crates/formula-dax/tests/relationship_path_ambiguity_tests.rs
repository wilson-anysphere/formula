use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, DaxError, FilterContext, Relationship,
    RowContext, Table, Value,
};

fn build_ambiguous_snowflake_model() -> DataModel {
    // Sales -> Products -> Categories is a snowflake path, but Sales also has a direct relationship
    // to Categories. That makes navigation between Sales and Categories ambiguous.
    let mut model = DataModel::new();

    let mut categories = Table::new("Categories", vec!["CategoryId", "CategoryName"]);
    categories
        .push_row(vec![1.into(), Value::from("A")])
        .unwrap();
    model.add_table(categories).unwrap();

    let mut products = Table::new("Products", vec!["ProductId", "CategoryId"]);
    products.push_row(vec![10.into(), 1.into()]).unwrap();
    model.add_table(products).unwrap();

    let mut sales = Table::new("Sales", vec!["SaleId", "ProductId", "CategoryId"]);
    sales.push_row(vec![100.into(), 10.into(), 1.into()]).unwrap();
    model.add_table(sales).unwrap();

    // Snowflake chain: Sales -> Products -> Categories.
    model
        .add_relationship(Relationship {
            name: "Sales_Products".into(),
            from_table: "Sales".into(),
            from_column: "ProductId".into(),
            to_table: "Products".into(),
            to_column: "ProductId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();
    model
        .add_relationship(Relationship {
            name: "Products_Categories".into(),
            from_table: "Products".into(),
            from_column: "CategoryId".into(),
            to_table: "Categories".into(),
            to_column: "CategoryId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    // Direct relationship: Sales -> Categories.
    model
        .add_relationship(Relationship {
            name: "Sales_Categories".into(),
            from_table: "Sales".into(),
            from_column: "CategoryId".into(),
            to_table: "Categories".into(),
            to_column: "CategoryId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
}

#[test]
fn related_errors_on_ambiguous_relationship_paths() {
    let mut model = build_ambiguous_snowflake_model();
    let err = model
        .add_calculated_column("Sales", "CategoryName", "RELATED(Categories[CategoryName])")
        .unwrap_err();

    match err {
        DaxError::Eval(message) => {
            let message_lc = message.to_ascii_lowercase();
            assert!(
                message_lc.contains("ambiguous active relationship path between sales and categories"),
                "unexpected error message: {message}"
            );
            assert!(
                message_lc.contains("sales -> products -> categories"),
                "expected error to include snowflake path, got: {message}"
            );
            assert!(
                message_lc.contains("sales -> categories"),
                "expected error to include direct path, got: {message}"
            );
        }
        other => panic!("expected DaxError::Eval, got {other:?}"),
    }
}

#[test]
fn relatedtable_errors_on_ambiguous_relationship_paths() {
    let model = build_ambiguous_snowflake_model();
    let engine = DaxEngine::new();

    let mut row_ctx = RowContext::default();
    row_ctx.push("Categories", 0);

    let err = engine
        .evaluate(
            &model,
            "COUNTROWS(RELATEDTABLE(Sales))",
            &FilterContext::empty(),
            &row_ctx,
        )
        .unwrap_err();

    match err {
        DaxError::Eval(message) => {
            let message_lc = message.to_ascii_lowercase();
            assert!(
                message_lc.contains("ambiguous active relationship path between categories and sales"),
                "unexpected error message: {message}"
            );
        }
        other => panic!("expected DaxError::Eval, got {other:?}"),
    }
}
