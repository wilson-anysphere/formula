use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions,
};
use formula_dax::{
    pivot, Cardinality, CrossFilterDirection, DataModel, DaxEngine, DaxError, FilterContext,
    GroupByColumn, PivotMeasure, Relationship, Table, Value,
};
use pretty_assertions::assert_eq;
use std::sync::Arc;

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

fn build_star_schema_columnar_model() -> DataModel {
    let mut model = DataModel::new();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };

    let customers_schema = vec![
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Region".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut customers = ColumnarTableBuilder::new(customers_schema, options);
    customers.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("East")),
    ]);
    customers.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("West")),
    ]);
    model
        .add_table(Table::from_columnar("Customers", customers.finalize()))
        .unwrap();

    let products_schema = vec![
        ColumnSchema {
            name: "ProductId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Category".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut products = ColumnarTableBuilder::new(products_schema, options);
    products.append_row(&[
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::String(Arc::<str>::from("A")),
    ]);
    products.append_row(&[
        formula_columnar::Value::Number(11.0),
        formula_columnar::Value::String(Arc::<str>::from("B")),
    ]);
    model
        .add_table(Table::from_columnar("Products", products.finalize()))
        .unwrap();

    let sales_schema = vec![
        ColumnSchema {
            name: "SaleId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "ProductId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut sales = ColumnarTableBuilder::new(sales_schema, options);
    sales.append_row(&[
        formula_columnar::Value::Number(100.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(10.0),
    ]); // East, A
    sales.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(11.0),
        formula_columnar::Value::Number(5.0),
    ]); // East, B
    sales.append_row(&[
        formula_columnar::Value::Number(102.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(7.0),
    ]); // West, A
    sales.append_row(&[
        formula_columnar::Value::Number(103.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(11.0),
        formula_columnar::Value::Number(3.0),
    ]); // West, B
    sales.append_row(&[
        formula_columnar::Value::Number(104.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(2.0),
    ]); // East, A
    model
        .add_table(Table::from_columnar("Sales", sales.finalize()))
        .unwrap();

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

fn build_star_schema_model_with_duplicate_dimension_attributes() -> DataModel {
    let mut model = DataModel::new();

    let mut customers = Table::new("Customers", vec!["CustomerId", "Region"]);
    customers.push_row(vec![1.into(), "East".into()]).unwrap();
    customers.push_row(vec![2.into(), "East".into()]).unwrap();
    customers.push_row(vec![3.into(), "West".into()]).unwrap();
    model.add_table(customers).unwrap();

    let mut products = Table::new("Products", vec!["ProductId", "Category"]);
    products.push_row(vec![10.into(), "A".into()]).unwrap();
    products.push_row(vec![11.into(), "A".into()]).unwrap();
    products.push_row(vec![12.into(), "B".into()]).unwrap();
    model.add_table(products).unwrap();

    let mut sales = Table::new("Sales", vec!["SaleId", "CustomerId", "ProductId", "Amount"]);
    sales
        .push_row(vec![100.into(), 1.into(), 10.into(), 10.0.into()])
        .unwrap(); // East, A
    sales
        .push_row(vec![101.into(), 2.into(), 11.into(), 5.0.into()])
        .unwrap(); // East, A (different customer/product id)
    sales
        .push_row(vec![102.into(), 3.into(), 10.into(), 7.0.into()])
        .unwrap(); // West, A
    sales
        .push_row(vec![103.into(), 1.into(), 12.into(), 2.0.into()])
        .unwrap(); // East, B
    sales
        .push_row(vec![104.into(), 2.into(), 12.into(), 3.0.into()])
        .unwrap(); // East, B
    sales
        .push_row(vec![105.into(), 3.into(), 12.into(), 4.0.into()])
        .unwrap(); // West, B
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

fn build_star_schema_columnar_model_with_duplicate_dimension_attributes() -> DataModel {
    let mut model = DataModel::new();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };

    let customers_schema = vec![
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Region".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut customers = ColumnarTableBuilder::new(customers_schema, options);
    customers.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("East")),
    ]);
    customers.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("East")),
    ]);
    customers.append_row(&[
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::String(Arc::<str>::from("West")),
    ]);
    model
        .add_table(Table::from_columnar("Customers", customers.finalize()))
        .unwrap();

    let products_schema = vec![
        ColumnSchema {
            name: "ProductId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Category".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut products = ColumnarTableBuilder::new(products_schema, options);
    products.append_row(&[
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::String(Arc::<str>::from("A")),
    ]);
    products.append_row(&[
        formula_columnar::Value::Number(11.0),
        formula_columnar::Value::String(Arc::<str>::from("A")),
    ]);
    products.append_row(&[
        formula_columnar::Value::Number(12.0),
        formula_columnar::Value::String(Arc::<str>::from("B")),
    ]);
    model
        .add_table(Table::from_columnar("Products", products.finalize()))
        .unwrap();

    let sales_schema = vec![
        ColumnSchema {
            name: "SaleId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "ProductId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut sales = ColumnarTableBuilder::new(sales_schema, options);
    sales.append_row(&[
        formula_columnar::Value::Number(100.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(10.0),
    ]); // East, A
    sales.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(11.0),
        formula_columnar::Value::Number(5.0),
    ]); // East, A (different customer/product id)
    sales.append_row(&[
        formula_columnar::Value::Number(102.0),
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(7.0),
    ]); // West, A
    sales.append_row(&[
        formula_columnar::Value::Number(103.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(12.0),
        formula_columnar::Value::Number(2.0),
    ]); // East, B
    sales.append_row(&[
        formula_columnar::Value::Number(104.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(12.0),
        formula_columnar::Value::Number(3.0),
    ]); // East, B
    sales.append_row(&[
        formula_columnar::Value::Number(105.0),
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::Number(12.0),
        formula_columnar::Value::Number(4.0),
    ]); // West, B
    model
        .add_table(Table::from_columnar("Sales", sales.finalize()))
        .unwrap();

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

#[test]
fn pivot_includes_blank_group_for_unmatched_relationship_keys() {
    let mut model = DataModel::new();

    let mut customers = Table::new("Customers", vec!["CustomerId", "Region"]);
    customers.push_row(vec![1.into(), "East".into()]).unwrap();
    customers.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(customers).unwrap();

    let mut sales = Table::new("Sales", vec!["SaleId", "CustomerId", "Amount"]);
    sales
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    sales
        .push_row(vec![101.into(), 2.into(), 5.0.into()])
        .unwrap();
    sales
        .push_row(vec![102.into(), 999.into(), 7.0.into()])
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
        .add_measure("Total Sales", "SUM(Sales[Amount])")
        .unwrap();

    let measures = vec![PivotMeasure::new("Total Sales", "[Total Sales]").unwrap()];
    let group_by = vec![GroupByColumn::new("Customers", "Region")];

    let result = pivot(
        &model,
        "Sales",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();
    assert_eq!(
        result.rows,
        vec![
            vec![Value::Blank, 7.0.into()],
            vec![Value::from("East"), 10.0.into()],
            vec![Value::from("West"), 5.0.into()],
        ]
    );

    // Applying a filter that explicitly excludes BLANK while including all real members should
    // remove the virtual blank row, and therefore exclude unmatched foreign keys from the pivot.
    let non_blank_filter = DaxEngine::new()
        .apply_calculate_filters(&model, &FilterContext::empty(), &["Customers[Region] <> BLANK()"])
        .unwrap();
    let result = pivot(&model, "Sales", &group_by, &measures, &non_blank_filter).unwrap();
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("East"), 10.0.into()],
            vec![Value::from("West"), 5.0.into()],
        ]
    );
}

#[test]
fn pivot_star_schema_columnar_matches_in_memory() {
    let vec_model = build_star_schema_model();
    let col_model = build_star_schema_columnar_model();

    let measures = vec![
        PivotMeasure::new("Total Sales", "[Total Sales]").unwrap(),
        PivotMeasure::new("Double Sales", "[Double Sales]").unwrap(),
    ];
    let group_by = vec![
        GroupByColumn::new("Customers", "Region"),
        GroupByColumn::new("Products", "Category"),
    ];

    let vec_result = pivot(
        &vec_model,
        "Sales",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();
    let col_result = pivot(
        &col_model,
        "Sales",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();
    assert_eq!(vec_result, col_result);

    let east_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    let vec_result = pivot(&vec_model, "Sales", &group_by, &measures, &east_filter).unwrap();
    let col_result = pivot(&col_model, "Sales", &group_by, &measures, &east_filter).unwrap();
    assert_eq!(vec_result, col_result);
}

#[test]
fn pivot_star_schema_columnar_rolls_up_duplicate_dimension_attributes() {
    let vec_model = build_star_schema_model_with_duplicate_dimension_attributes();
    let col_model = build_star_schema_columnar_model_with_duplicate_dimension_attributes();

    let measures = vec![
        PivotMeasure::new("Total Sales", "[Total Sales]").unwrap(),
        PivotMeasure::new("Double Sales", "[Double Sales]").unwrap(),
    ];
    let group_by = vec![
        GroupByColumn::new("Customers", "Region"),
        GroupByColumn::new("Products", "Category"),
    ];

    let vec_result = pivot(
        &vec_model,
        "Sales",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();
    let col_result = pivot(
        &col_model,
        "Sales",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();
    assert_eq!(vec_result, col_result);

    assert_eq!(
        col_result.rows,
        vec![
            vec![Value::from("East"), Value::from("A"), 15.0.into(), 30.0.into()],
            vec![Value::from("East"), Value::from("B"), 5.0.into(), 10.0.into()],
            vec![Value::from("West"), Value::from("A"), 7.0.into(), 14.0.into()],
            vec![Value::from("West"), Value::from("B"), 4.0.into(), 8.0.into()],
        ]
    );
}

#[test]
fn pivot_star_schema_errors_on_ambiguous_relationship_paths() {
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };

    let customers_schema = vec![
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Region".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut customers = ColumnarTableBuilder::new(customers_schema, options);
    customers.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("East")),
    ]);
    customers.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("West")),
    ]);

    let sales_schema = vec![
        ColumnSchema {
            name: "SaleId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "CustomerId1".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "CustomerId2".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut sales = ColumnarTableBuilder::new(sales_schema, options);
    sales.append_row(&[
        formula_columnar::Value::Number(100.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(10.0),
    ]);

    let mut model = DataModel::new();
    model
        .add_table(Table::from_columnar("Customers", customers.finalize()))
        .unwrap();
    model
        .add_table(Table::from_columnar("Sales", sales.finalize()))
        .unwrap();

    model
        .add_relationship(Relationship {
            name: "Sales_Customers_1".into(),
            from_table: "Sales".into(),
            from_column: "CustomerId1".into(),
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
            name: "Sales_Customers_2".into(),
            from_table: "Sales".into(),
            from_column: "CustomerId2".into(),
            to_table: "Customers".into(),
            to_column: "CustomerId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_measure("Total Sales", "SUM(Sales[Amount])")
        .unwrap();

    let measures = vec![PivotMeasure::new("Total Sales", "[Total Sales]").unwrap()];
    let group_by = vec![GroupByColumn::new("Customers", "Region")];

    let err = pivot(
        &model,
        "Sales",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap_err();

    match err {
        DaxError::Eval(message) => assert!(
            message.contains("ambiguous active relationship path"),
            "unexpected error message: {message}"
        ),
        other => panic!("expected DaxError::Eval, got {other:?}"),
    }
}
