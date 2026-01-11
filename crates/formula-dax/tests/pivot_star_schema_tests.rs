use formula_dax::{
    pivot, Cardinality, CrossFilterDirection, DataModel, FilterContext, GroupByColumn,
    PivotMeasure, Relationship, Table, Value,
};
use pretty_assertions::assert_eq;

fn build_star_schema_model() -> DataModel {
    let mut model = DataModel::new();

    let mut customers = Table::new("Customers", vec!["CustomerId", "Region"]);
    customers.push_row(vec![1.into(), "East".into()]).unwrap();
    customers.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(customers).unwrap();

    let mut products = Table::new("Products", vec!["ProductId", "Category"]);
    products.push_row(vec![10.into(), "A".into()]).unwrap();
    products.push_row(vec![11.into(), "B".into()]).unwrap();
    model.add_table(products).unwrap();

    let mut sales = Table::new("Sales", vec!["SaleId", "CustomerId", "ProductId", "Amount"]);
    sales
        .push_row(vec![100.into(), 1.into(), 10.into(), 10.0.into()])
        .unwrap(); // East, A
    sales
        .push_row(vec![101.into(), 1.into(), 11.into(), 5.0.into()])
        .unwrap(); // East, B
    sales
        .push_row(vec![102.into(), 2.into(), 10.into(), 7.0.into()])
        .unwrap(); // West, A
    sales
        .push_row(vec![103.into(), 2.into(), 11.into(), 3.0.into()])
        .unwrap(); // West, B
    sales
        .push_row(vec![104.into(), 1.into(), 10.into(), 2.0.into()])
        .unwrap(); // East, A
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
        .add_measure("Total Sales", "SUM(Sales[Amount])")
        .unwrap();
    model
        .add_measure("Double Sales", "[Total Sales] * 2")
        .unwrap();

    model
}

#[test]
fn pivot_star_schema_groups_by_multiple_dimensions_and_supports_nested_measures() {
    let model = build_star_schema_model();

    let measures = vec![
        PivotMeasure::new("Total Sales", "[Total Sales]").unwrap(),
        PivotMeasure::new("Double Sales", "[Double Sales]").unwrap(),
    ];
    let group_by = vec![
        GroupByColumn::new("Customers", "Region"),
        GroupByColumn::new("Products", "Category"),
    ];

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
            "Customers[Region]".to_string(),
            "Products[Category]".to_string(),
            "Total Sales".to_string(),
            "Double Sales".to_string(),
        ]
    );
    assert_eq!(
        result.rows,
        vec![
            vec![
                Value::from("East"),
                Value::from("A"),
                12.0.into(),
                24.0.into(),
            ],
            vec![
                Value::from("East"),
                Value::from("B"),
                5.0.into(),
                10.0.into(),
            ],
            vec![
                Value::from("West"),
                Value::from("A"),
                7.0.into(),
                14.0.into(),
            ],
            vec![
                Value::from("West"),
                Value::from("B"),
                3.0.into(),
                6.0.into(),
            ],
        ]
    );
}

#[test]
fn pivot_star_schema_respects_dimension_filters() {
    let model = build_star_schema_model();

    let measures = vec![
        PivotMeasure::new("Total Sales", "[Total Sales]").unwrap(),
        PivotMeasure::new("Double Sales", "[Double Sales]").unwrap(),
    ];
    let group_by = vec![GroupByColumn::new("Products", "Category")];
    let filter = FilterContext::empty().with_column_equals("Customers", "Region", "East".into());

    let result = pivot(&model, "Sales", &group_by, &measures, &filter).unwrap();
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("A"), 12.0.into(), 24.0.into()],
            vec![Value::from("B"), 5.0.into(), 10.0.into()],
        ]
    );
}
