use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
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

fn build_snowflake_model_with_columnar_sales() -> DataModel {
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

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let sales_schema = vec![
        ColumnSchema {
            name: "SaleId".to_string(),
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
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(10.0),
    ]); // A
    sales.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(11.0),
        formula_columnar::Value::Number(5.0),
    ]); // A
    sales.append_row(&[
        formula_columnar::Value::Number(102.0),
        formula_columnar::Value::Number(20.0),
        formula_columnar::Value::Number(7.0),
    ]); // B
    model
        .add_table(Table::from_columnar("Sales", sales.finalize()))
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

fn build_snowflake_model_with_alternate_product_category_relationship() -> DataModel {
    let mut model = DataModel::new();

    let mut categories = Table::new("Categories", vec!["CategoryId", "CategoryName"]);
    categories
        .push_row(vec![1.into(), Value::from("A")])
        .unwrap();
    categories
        .push_row(vec![2.into(), Value::from("B")])
        .unwrap();
    model.add_table(categories).unwrap();

    let mut products = Table::new("Products", vec!["ProductId", "CategoryId", "AltCategoryId"]);
    // Product 10: default category A, alternate category B.
    products.push_row(vec![10.into(), 1.into(), 2.into()]).unwrap();
    // Product 20: default category B, alternate category A.
    products.push_row(vec![20.into(), 2.into(), 1.into()]).unwrap();
    model.add_table(products).unwrap();

    let mut sales = Table::new("Sales", vec!["SaleId", "ProductId", "Amount"]);
    sales
        .push_row(vec![100.into(), 10.into(), 10.0.into()])
        .unwrap(); // default A, alt B
    sales
        .push_row(vec![101.into(), 20.into(), 7.0.into()])
        .unwrap(); // default B, alt A
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

    // Default active relationship.
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

    // Alternate inactive relationship.
    model
        .add_relationship(Relationship {
            name: "Products_Categories_Alt".into(),
            from_table: "Products".into(),
            from_column: "AltCategoryId".into(),
            to_table: "Categories".into(),
            to_column: "CategoryId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: false,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
}

fn build_snowflake_model_with_columnar_sales_and_alternate_product_category_relationship() -> DataModel {
    let mut model = DataModel::new();

    let mut categories = Table::new("Categories", vec!["CategoryId", "CategoryName"]);
    categories
        .push_row(vec![1.into(), Value::from("A")])
        .unwrap();
    categories
        .push_row(vec![2.into(), Value::from("B")])
        .unwrap();
    model.add_table(categories).unwrap();

    let mut products = Table::new("Products", vec!["ProductId", "CategoryId", "AltCategoryId"]);
    // Product 10: default category A, alternate category B.
    products.push_row(vec![10.into(), 1.into(), 2.into()]).unwrap();
    // Product 20: default category B, alternate category A.
    products.push_row(vec![20.into(), 2.into(), 1.into()]).unwrap();
    model.add_table(products).unwrap();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let sales_schema = vec![
        ColumnSchema {
            name: "SaleId".to_string(),
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
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(10.0),
    ]); // default A, alt B
    sales.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(20.0),
        formula_columnar::Value::Number(7.0),
    ]); // default B, alt A
    model
        .add_table(Table::from_columnar("Sales", sales.finalize()))
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

    // Default active relationship.
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

    // Alternate inactive relationship.
    model
        .add_relationship(Relationship {
            name: "Products_Categories_Alt".into(),
            from_table: "Products".into(),
            from_column: "AltCategoryId".into(),
            to_table: "Categories".into(),
            to_column: "CategoryId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: false,
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
fn related_supports_multi_hop_snowflake_navigation_columnar_fact() {
    let mut model = build_snowflake_model_with_columnar_sales();
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
fn related_supports_userelationship_override_across_snowflake_hops() {
    let mut model = build_snowflake_model_with_alternate_product_category_relationship();

    model
        .add_calculated_column("Sales", "DefaultCategory", "RELATED(Categories[CategoryName])")
        .unwrap();
    model
        .add_calculated_column(
            "Sales",
            "AltCategory",
            "CALCULATE(RELATED(Categories[CategoryName]), USERELATIONSHIP(Products[AltCategoryId], Categories[CategoryId]))",
        )
        .unwrap();

    let sales = model.table("Sales").unwrap();
    let values: Vec<(Value, Value)> = (0..sales.row_count())
        .map(|row| {
            (
                sales.value(row, "DefaultCategory").unwrap(),
                sales.value(row, "AltCategory").unwrap(),
            )
        })
        .collect();

    assert_eq!(
        values,
        vec![(Value::from("A"), Value::from("B")), (Value::from("B"), Value::from("A"))]
    );
}

#[test]
fn related_supports_userelationship_override_across_snowflake_hops_columnar_fact() {
    let mut model = build_snowflake_model_with_columnar_sales_and_alternate_product_category_relationship();

    model
        .add_calculated_column("Sales", "DefaultCategory", "RELATED(Categories[CategoryName])")
        .unwrap();
    model
        .add_calculated_column(
            "Sales",
            "AltCategory",
            "CALCULATE(RELATED(Categories[CategoryName]), USERELATIONSHIP(Products[AltCategoryId], Categories[CategoryId]))",
        )
        .unwrap();

    let sales = model.table("Sales").unwrap();
    let values: Vec<(Value, Value)> = (0..sales.row_count())
        .map(|row| {
            (
                sales.value(row, "DefaultCategory").unwrap(),
                sales.value(row, "AltCategory").unwrap(),
            )
        })
        .collect();

    assert_eq!(
        values,
        vec![(Value::from("A"), Value::from("B")), (Value::from("B"), Value::from("A"))]
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
fn pivot_grouping_supports_multi_hop_snowflake_dimensions_columnar_fact() {
    let model = build_snowflake_model_with_columnar_sales();

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
        result.rows,
        vec![
            vec![Value::from("A"), 15.0.into()],
            vec![Value::from("B"), 7.0.into()],
        ]
    );
}

#[test]
fn pivot_grouping_respects_userelationship_override_for_snowflake_dimensions() {
    let model = build_snowflake_model_with_alternate_product_category_relationship();
    let engine = DaxEngine::new();

    let filter = engine
        .apply_calculate_filters(
            &model,
            &FilterContext::empty(),
            &["USERELATIONSHIP(Products[AltCategoryId], Categories[CategoryId])"],
        )
        .unwrap();

    let group_by = vec![GroupByColumn::new("Categories", "CategoryName")];
    let measures = vec![PivotMeasure::new("Total Amount", "SUM(Sales[Amount])").unwrap()];

    let result = pivot(&model, "Sales", &group_by, &measures, &filter).unwrap();

    // Under the alternate relationship, the category mapping is swapped.
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("A"), 7.0.into()],
            vec![Value::from("B"), 10.0.into()],
        ]
    );
}

#[test]
fn pivot_grouping_respects_userelationship_override_for_snowflake_dimensions_columnar_fact() {
    let model = build_snowflake_model_with_columnar_sales_and_alternate_product_category_relationship();
    let engine = DaxEngine::new();

    let filter = engine
        .apply_calculate_filters(
            &model,
            &FilterContext::empty(),
            &["USERELATIONSHIP(Products[AltCategoryId], Categories[CategoryId])"],
        )
        .unwrap();

    let group_by = vec![GroupByColumn::new("Categories", "CategoryName")];
    let measures = vec![PivotMeasure::new("Total Amount", "SUM(Sales[Amount])").unwrap()];

    let result = pivot(&model, "Sales", &group_by, &measures, &filter).unwrap();

    // Under the alternate relationship, the category mapping is swapped.
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("A"), 7.0.into()],
            vec![Value::from("B"), 10.0.into()],
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
fn pivot_crosstab_supports_multi_hop_snowflake_dimensions_columnar_fact() {
    let model = build_snowflake_model_with_columnar_sales();

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
fn summarize_grouping_supports_multi_hop_snowflake_dimensions_columnar_fact() {
    let model = build_snowflake_model_with_columnar_sales();
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

#[test]
fn relatedtable_supports_multi_hop_snowflake_navigation_columnar_fact() {
    let mut model = build_snowflake_model_with_columnar_sales();
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

#[test]
fn relatedtable_supports_userelationship_override_across_snowflake_hops() {
    let mut model = build_snowflake_model_with_alternate_product_category_relationship();

    model
        .add_calculated_column(
            "Categories",
            "Default Total Amount",
            "SUMX(RELATEDTABLE(Sales), Sales[Amount])",
        )
        .unwrap();
    model
        .add_calculated_column(
            "Categories",
            "Alt Total Amount",
            "CALCULATE(SUMX(RELATEDTABLE(Sales), Sales[Amount]), USERELATIONSHIP(Products[AltCategoryId], Categories[CategoryId]))",
        )
        .unwrap();

    let categories = model.table("Categories").unwrap();
    let values: Vec<(Value, Value)> = (0..categories.row_count())
        .map(|row| {
            (
                categories.value(row, "Default Total Amount").unwrap(),
                categories.value(row, "Alt Total Amount").unwrap(),
            )
        })
        .collect();

    // Category A: default is 10, alternate is 7. Category B: default is 7, alternate is 10.
    assert_eq!(values, vec![(10.0.into(), 7.0.into()), (7.0.into(), 10.0.into())]);
}

#[test]
fn relatedtable_supports_userelationship_override_across_snowflake_hops_columnar_fact() {
    let mut model = build_snowflake_model_with_columnar_sales_and_alternate_product_category_relationship();

    model
        .add_calculated_column(
            "Categories",
            "Default Total Amount",
            "SUMX(RELATEDTABLE(Sales), Sales[Amount])",
        )
        .unwrap();
    model
        .add_calculated_column(
            "Categories",
            "Alt Total Amount",
            "CALCULATE(SUMX(RELATEDTABLE(Sales), Sales[Amount]), USERELATIONSHIP(Products[AltCategoryId], Categories[CategoryId]))",
        )
        .unwrap();

    let categories = model.table("Categories").unwrap();
    let values: Vec<(Value, Value)> = (0..categories.row_count())
        .map(|row| {
            (
                categories.value(row, "Default Total Amount").unwrap(),
                categories.value(row, "Alt Total Amount").unwrap(),
            )
        })
        .collect();

    // Category A: default is 10, alternate is 7. Category B: default is 7, alternate is 10.
    assert_eq!(values, vec![(10.0.into(), 7.0.into()), (7.0.into(), 10.0.into())]);
}

#[test]
fn relatedtable_cascades_blank_rows_across_snowflake_hops() {
    // Scenario:
    // - Sales contains a ProductId with no matching Products row.
    // - That creates a virtual blank row in Products.
    // - That blank Products row should be treated as belonging to the blank Categories member,
    //   so navigating Categories(blank) -> RELATEDTABLE(Sales) should include that sales row.
    let mut model = DataModel::new();

    let mut categories = Table::new("Categories", vec!["CategoryId", "CategoryName"]);
    categories.push_row(vec![1.into(), Value::from("A")]).unwrap();
    model.add_table(categories).unwrap();

    let mut products = Table::new("Products", vec!["ProductId", "CategoryId"]);
    products.push_row(vec![10.into(), 1.into()]).unwrap();
    model.add_table(products).unwrap();

    let mut sales = Table::new("Sales", vec!["SaleId", "ProductId", "Amount"]);
    sales.push_row(vec![100.into(), 10.into(), 10.0.into()]).unwrap(); // A
    sales
        .push_row(vec![101.into(), 999.into(), 5.0.into()])
        .unwrap(); // unknown product -> Products blank row -> Categories blank row
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
            enforce_referential_integrity: false,
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

    let mut row_ctx = RowContext::default();
    let categories_blank_row = model.table("Categories").unwrap().row_count();
    row_ctx.push("Categories", categories_blank_row);

    let value = DaxEngine::new()
        .evaluate(
            &model,
            "SUMX(RELATEDTABLE(Sales), Sales[Amount])",
            &FilterContext::empty(),
            &row_ctx,
        )
        .unwrap();

    assert_eq!(value, 5.0.into());
}

#[test]
fn relatedtable_does_not_treat_physical_blank_dimension_rows_as_virtual_blank_members() {
    // Regression: the relationship-generated blank member is distinct from a physical BLANK key
    // row on the dimension side. RELATEDTABLE should only include unmatched rows when navigating
    // from the *virtual* blank member (row_count), not when the current physical row happens to
    // have a BLANK key value.
    let mut model = DataModel::new();

    let mut categories = Table::new("Categories", vec!["CategoryId", "CategoryName"]);
    categories
        .push_row(vec![1.into(), Value::from("A")])
        .unwrap();
    // Physical row whose key is BLANK.
    categories
        .push_row(vec![Value::Blank, Value::from("PhysicalBlank")])
        .unwrap();
    model.add_table(categories).unwrap();

    let mut products = Table::new("Products", vec!["ProductId", "CategoryId"]);
    products.push_row(vec![10.into(), 1.into()]).unwrap();
    model.add_table(products).unwrap();

    let mut sales = Table::new("Sales", vec!["SaleId", "ProductId", "Amount"]);
    sales
        .push_row(vec![100.into(), 10.into(), 10.0.into()])
        .unwrap(); // A
    sales
        .push_row(vec![101.into(), 999.into(), 5.0.into()])
        .unwrap(); // unknown product -> Products blank row -> Categories blank row
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
            enforce_referential_integrity: false,
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

    let mut row_ctx = RowContext::default();
    // Physical blank-key Categories row.
    row_ctx.push("Categories", 1);

    let value = DaxEngine::new()
        .evaluate(
            &model,
            "COUNTROWS(RELATEDTABLE(Sales))",
            &FilterContext::empty(),
            &row_ctx,
        )
        .unwrap();

    assert_eq!(value, 0.into());
}

#[test]
fn relatedtable_cascades_blank_rows_across_snowflake_hops_columnar_fact() {
    // Same scenario as `relatedtable_cascades_blank_rows_across_snowflake_hops`, but with a
    // columnar fact table (no `from_index`, relies on `unmatched_fact_rows`).
    let mut model = DataModel::new();

    let mut categories = Table::new("Categories", vec!["CategoryId", "CategoryName"]);
    categories.push_row(vec![1.into(), Value::from("A")]).unwrap();
    model.add_table(categories).unwrap();

    let mut products = Table::new("Products", vec!["ProductId", "CategoryId"]);
    products.push_row(vec![10.into(), 1.into()]).unwrap();
    model.add_table(products).unwrap();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let sales_schema = vec![
        ColumnSchema {
            name: "SaleId".to_string(),
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
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(10.0),
    ]); // A
    sales.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(999.0),
        formula_columnar::Value::Number(5.0),
    ]); // unknown product -> Products blank row -> Categories blank row
    model
        .add_table(Table::from_columnar("Sales", sales.finalize()))
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
            enforce_referential_integrity: false,
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

    let mut row_ctx = RowContext::default();
    let categories_blank_row = model.table("Categories").unwrap().row_count();
    row_ctx.push("Categories", categories_blank_row);

    let value = DaxEngine::new()
        .evaluate(
            &model,
            "SUMX(RELATEDTABLE(Sales), Sales[Amount])",
            &FilterContext::empty(),
            &row_ctx,
        )
        .unwrap();

    assert_eq!(value, 5.0.into());
}

#[test]
fn snowflake_filter_excludes_unmatched_fact_rows_when_blank_category_filtered_out() {
    // Scenario:
    // - Sales contains an unknown ProductId (no matching Products row).
    // - That creates a virtual blank row in Products, which in turn belongs to the blank
    //   Categories member through the Products -> Categories relationship.
    //
    // When a filter on Categories excludes BLANK (e.g. CategoryName = "A"), the blank Categories
    // member is not visible, so the virtual blank Products member should also be filtered out.
    // Consequently, the unmatched Sales row should not be included in measures grouped by, or
    // filtered by, Categories.
    let mut model = DataModel::new();

    let mut categories = Table::new("Categories", vec!["CategoryId", "CategoryName"]);
    categories.push_row(vec![1.into(), Value::from("A")]).unwrap();
    model.add_table(categories).unwrap();

    let mut products = Table::new("Products", vec!["ProductId", "CategoryId"]);
    products.push_row(vec![10.into(), 1.into()]).unwrap();
    model.add_table(products).unwrap();

    let mut sales = Table::new("Sales", vec!["SaleId", "ProductId", "Amount"]);
    sales.push_row(vec![100.into(), 10.into(), 10.0.into()]).unwrap(); // A
    sales
        .push_row(vec![101.into(), 999.into(), 5.0.into()])
        .unwrap(); // unknown product -> Products blank row -> Categories blank row
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
            enforce_referential_integrity: false,
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

    let filter = FilterContext::empty().with_column_equals("Categories", "CategoryName", "A".into());
    let engine = DaxEngine::new();

    let value = engine
        .evaluate(
            &model,
            "SUM(Sales[Amount])",
            &filter,
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 10.0.into());

    let group_by = vec![GroupByColumn::new("Categories", "CategoryName")];
    let measures = vec![PivotMeasure::new("Total Amount", "SUM(Sales[Amount])").unwrap()];
    let result = pivot(&model, "Sales", &group_by, &measures, &filter).unwrap();

    assert_eq!(result.rows, vec![vec![Value::from("A"), 10.0.into()]]);
}

#[test]
fn snowflake_filter_excludes_unmatched_fact_rows_when_blank_category_filtered_out_columnar_fact() {
    // Same scenario as `snowflake_filter_excludes_unmatched_fact_rows_when_blank_category_filtered_out`,
    // but with a columnar fact table (exercises `unmatched_fact_rows` without a `from_index`).
    let mut model = DataModel::new();

    let mut categories = Table::new("Categories", vec!["CategoryId", "CategoryName"]);
    categories.push_row(vec![1.into(), Value::from("A")]).unwrap();
    model.add_table(categories).unwrap();

    let mut products = Table::new("Products", vec!["ProductId", "CategoryId"]);
    products.push_row(vec![10.into(), 1.into()]).unwrap();
    model.add_table(products).unwrap();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let sales_schema = vec![
        ColumnSchema {
            name: "SaleId".to_string(),
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
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(10.0),
    ]); // A
    sales.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(999.0),
        formula_columnar::Value::Number(5.0),
    ]); // unknown product -> Products blank row -> Categories blank row
    model
        .add_table(Table::from_columnar("Sales", sales.finalize()))
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
            enforce_referential_integrity: false,
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

    let filter = FilterContext::empty().with_column_equals("Categories", "CategoryName", "A".into());
    let engine = DaxEngine::new();

    let value = engine
        .evaluate(
            &model,
            "SUM(Sales[Amount])",
            &filter,
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 10.0.into());

    let group_by = vec![GroupByColumn::new("Categories", "CategoryName")];
    let measures = vec![PivotMeasure::new("Total Amount", "SUM(Sales[Amount])").unwrap()];
    let result = pivot(&model, "Sales", &group_by, &measures, &filter).unwrap();

    assert_eq!(result.rows, vec![vec![Value::from("A"), 10.0.into()]]);
}
