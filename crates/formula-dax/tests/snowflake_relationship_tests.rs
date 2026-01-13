use formula_dax::{
    pivot, pivot_crosstab, Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext,
    GroupByColumn, PivotMeasure, PivotResultGrid, Relationship, RowContext, Table, Value,
};
use pretty_assertions::assert_eq;

fn build_snowflake_model() -> DataModel {
    let mut model = DataModel::new();

    let mut categories = Table::new("Categories", vec!["CategoryId", "CategoryName"]);
    categories
        .push_row(vec![1.into(), Value::from("A")])
        .unwrap();
    categories
        .push_row(vec![2.into(), Value::from("B")])
        .unwrap();
    model.add_table(categories).unwrap();

    let mut products = Table::new("Products", vec!["ProductId", "CategoryId"]);
    products.push_row(vec![10.into(), 1.into()]).unwrap();
    products.push_row(vec![11.into(), 1.into()]).unwrap();
    products.push_row(vec![20.into(), 2.into()]).unwrap();
    model.add_table(products).unwrap();

    let mut sales = Table::new("Sales", vec!["SaleId", "ProductId", "Amount"]);
    sales
        .push_row(vec![100.into(), 10.into(), 10.0.into()])
        .unwrap(); // A
    sales
        .push_row(vec![101.into(), 11.into(), 5.0.into()])
        .unwrap(); // A
    sales
        .push_row(vec![102.into(), 20.into(), 7.0.into()])
        .unwrap(); // B
    model.add_table(sales).unwrap();

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

    model
}

#[test]
fn related_supports_multi_hop_snowflake_navigation() {
    let mut model = build_snowflake_model();
    model
        .add_calculated_column("Sales", "CategoryName", "RELATED(Categories[CategoryName])")
        .unwrap();

    let sales = model.table("Sales").unwrap();
    let values: Vec<Value> = (0..sales.row_count())
        .map(|row| sales.value(row, "CategoryName").unwrap())
        .collect();

    assert_eq!(
        values,
        vec![Value::from("A"), Value::from("A"), Value::from("B")]
    );
}

#[test]
fn pivot_grouping_supports_multi_hop_snowflake_dimensions() {
    let model = build_snowflake_model();

    let group_by = vec![GroupByColumn::new("Categories", "CategoryName")];
    let measures = vec![PivotMeasure::new("Total Amount", "SUM(Sales[Amount])").unwrap()];

    let result = pivot(
        &model,
        "Sales",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(
        result.columns,
        vec![
            "Categories[CategoryName]".to_string(),
            "Total Amount".to_string()
        ]
    );
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("A"), 15.0.into()],
            vec![Value::from("B"), 7.0.into()],
        ]
    );
}

#[test]
fn pivot_crosstab_supports_multi_hop_snowflake_dimensions() {
    let model = build_snowflake_model();

    let row_fields = vec![GroupByColumn::new("Categories", "CategoryName")];
    let column_fields = vec![GroupByColumn::new("Products", "ProductId")];
    let measures = vec![PivotMeasure::new("Total Amount", "SUM(Sales[Amount])").unwrap()];

    let result = pivot_crosstab(
        &model,
        "Sales",
        &row_fields,
        &column_fields,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(
        result,
        PivotResultGrid {
            data: vec![
                vec![
                    Value::from("Categories[CategoryName]"),
                    Value::from("10"),
                    Value::from("11"),
                    Value::from("20"),
                ],
                vec![Value::from("A"), 10.0.into(), 5.0.into(), Value::Blank],
                vec![Value::from("B"), Value::Blank, Value::Blank, 7.0.into()],
            ]
        }
    );
}

#[test]
fn summarize_grouping_supports_multi_hop_snowflake_dimensions() {
    let model = build_snowflake_model();
    let engine = DaxEngine::new();

    let value = engine
        .evaluate(
            &model,
            "COUNTROWS(SUMMARIZE(Sales, Categories[CategoryName]))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, 2.into());
}

#[test]
fn relatedtable_supports_multi_hop_snowflake_navigation() {
    let mut model = build_snowflake_model();
    model
        .add_calculated_column(
            "Categories",
            "Total Amount",
            "SUMX(RELATEDTABLE(Sales), Sales[Amount])",
        )
        .unwrap();

    let categories = model.table("Categories").unwrap();
    let values: Vec<Value> = (0..categories.row_count())
        .map(|row| categories.value(row, "Total Amount").unwrap())
        .collect();

    assert_eq!(values, vec![15.0.into(), 7.0.into()]);
}
