use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, Relationship, RowContext,
    Table, Value,
};
use pretty_assertions::assert_eq;

fn build_model_blank_row_bidirectional() -> DataModel {
    build_model_blank_row_bidirectional_with_cardinality(Cardinality::OneToMany)
}

fn build_model_blank_row_bidirectional_with_cardinality(cardinality: Cardinality) -> DataModel {
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
            cardinality,
            cross_filter_direction: CrossFilterDirection::Both,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
}

fn build_model_blank_row_bidirectional_columnar_fact() -> DataModel {
    build_model_blank_row_bidirectional_columnar_fact_with_cardinality(Cardinality::OneToMany)
}

fn build_model_blank_row_bidirectional_columnar_fact_with_cardinality(
    cardinality: Cardinality,
) -> DataModel {
    let mut model = DataModel::new();

    let mut customers = Table::new("Customers", vec!["CustomerId", "Region"]);
    customers.push_row(vec![1.into(), "East".into()]).unwrap();
    customers.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(customers).unwrap();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };
    let orders_schema = vec![
        ColumnSchema {
            name: "OrderId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut orders = ColumnarTableBuilder::new(orders_schema, options);
    orders.append_row(&[
        formula_columnar::Value::Number(100.0),
        formula_columnar::Value::Number(1.0),
    ]);
    orders.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(999.0),
    ]);
    model
        .add_table(Table::from_columnar("Orders", orders.finalize()))
        .unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders_Customers".into(),
            from_table: "Orders".into(),
            from_column: "CustomerId".into(),
            to_table: "Customers".into(),
            to_column: "CustomerId".into(),
            cardinality,
            cross_filter_direction: CrossFilterDirection::Both,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
}

fn build_model_blank_row_bidirectional_columnar_dim_with_cardinality(
    cardinality: Cardinality,
) -> DataModel {
    let mut model = DataModel::new();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
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
        formula_columnar::Value::String("East".into()),
    ]);
    customers.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String("West".into()),
    ]);
    model
        .add_table(Table::from_columnar("Customers", customers.finalize()))
        .unwrap();

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
            cardinality,
            cross_filter_direction: CrossFilterDirection::Both,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
}

fn build_model_blank_row_bidirectional_columnar_dim_and_fact_with_cardinality(
    cardinality: Cardinality,
) -> DataModel {
    let mut model = DataModel::new();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
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
        formula_columnar::Value::String("East".into()),
    ]);
    customers.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String("West".into()),
    ]);
    model
        .add_table(Table::from_columnar("Customers", customers.finalize()))
        .unwrap();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };
    let orders_schema = vec![
        ColumnSchema {
            name: "OrderId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut orders = ColumnarTableBuilder::new(orders_schema, options);
    orders.append_row(&[
        formula_columnar::Value::Number(100.0),
        formula_columnar::Value::Number(1.0),
    ]);
    orders.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(999.0),
    ]);
    model
        .add_table(Table::from_columnar("Orders", orders.finalize()))
        .unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders_Customers".into(),
            from_table: "Orders".into(),
            from_column: "CustomerId".into(),
            to_table: "Customers".into(),
            to_column: "CustomerId".into(),
            cardinality,
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

#[test]
fn values_virtual_blank_row_respects_bidirectional_filter_context_many_to_many() {
    let model = build_model_blank_row_bidirectional_with_cardinality(Cardinality::ManyToMany);
    let engine = DaxEngine::new();

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
fn values_virtual_blank_row_respects_bidirectional_filter_context_many_to_many_columnar_dim() {
    let model =
        build_model_blank_row_bidirectional_columnar_dim_with_cardinality(Cardinality::ManyToMany);
    let engine = DaxEngine::new();

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
fn values_virtual_blank_row_respects_bidirectional_filter_context_many_to_many_columnar_dim_and_fact() {
    let model = build_model_blank_row_bidirectional_columnar_dim_and_fact_with_cardinality(
        Cardinality::ManyToMany,
    );
    let engine = DaxEngine::new();

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
fn distinctcount_virtual_blank_row_respects_bidirectional_filter_context_many_to_many() {
    let model = build_model_blank_row_bidirectional_with_cardinality(Cardinality::ManyToMany);
    let engine = DaxEngine::new();

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

#[test]
fn summarizecolumns_virtual_blank_row_respects_bidirectional_filter_context() {
    let model = build_model_blank_row_bidirectional();
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn summarizecolumns_virtual_blank_row_respects_bidirectional_filter_context_many_to_many() {
    let model = build_model_blank_row_bidirectional_with_cardinality(Cardinality::ManyToMany);
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn countblank_virtual_blank_row_respects_bidirectional_filter_context() {
    let model = build_model_blank_row_bidirectional();
    let engine = DaxEngine::new();

    // No extra filters: virtual blank member exists -> COUNTBLANK should count it.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTBLANK(Customers[Region])",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );

    // Filter fact to a matched key: virtual blank member is not in context.
    let matched_filter = FilterContext::empty().with_column_equals("Orders", "CustomerId", 1.into());
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTBLANK(Customers[Region])",
                &matched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        0.into()
    );

    // Filter fact to an unmatched key: only the virtual blank member remains.
    let unmatched_filter =
        FilterContext::empty().with_column_equals("Orders", "CustomerId", 999.into());
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTBLANK(Customers[Region])",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn countblank_virtual_blank_row_respects_bidirectional_filter_context_columnar_fact() {
    let model = build_model_blank_row_bidirectional_columnar_fact();
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTBLANK(Customers[Region])",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );

    let matched_filter = FilterContext::empty().with_column_equals("Orders", "CustomerId", 1.into());
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTBLANK(Customers[Region])",
                &matched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        0.into()
    );

    let unmatched_filter =
        FilterContext::empty().with_column_equals("Orders", "CustomerId", 999.into());
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTBLANK(Customers[Region])",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn all_column_virtual_blank_row_respects_bidirectional_filter_context() {
    let model = build_model_blank_row_bidirectional();
    let engine = DaxEngine::new();

    // Unfiltered: East, West, BLANK (unknown member).
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(ALL(Customers[Region]))",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        3.into()
    );

    // Filter fact to a matched key: only East should remain (no BLANK).
    let matched_filter = FilterContext::empty().with_column_equals("Orders", "CustomerId", 1.into());
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(ALL(Customers[Region]))",
                &matched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );

    // Filter fact to an unmatched key: only the virtual blank member remains.
    let unmatched_filter =
        FilterContext::empty().with_column_equals("Orders", "CustomerId", 999.into());
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(ALL(Customers[Region]))",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn all_column_virtual_blank_row_respects_bidirectional_filter_context_columnar_fact() {
    let model = build_model_blank_row_bidirectional_columnar_fact();
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(ALL(Customers[Region]))",
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
                "COUNTROWS(ALL(Customers[Region]))",
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
                "COUNTROWS(ALL(Customers[Region]))",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn values_virtual_blank_row_respects_bidirectional_filter_context_columnar_fact() {
    let model = build_model_blank_row_bidirectional_columnar_fact();
    let engine = DaxEngine::new();

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
fn values_virtual_blank_row_respects_bidirectional_filter_context_columnar_dim() {
    let model =
        build_model_blank_row_bidirectional_columnar_dim_with_cardinality(Cardinality::OneToMany);
    let engine = DaxEngine::new();

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
fn values_virtual_blank_row_respects_bidirectional_filter_context_columnar_dim_and_fact() {
    let model =
        build_model_blank_row_bidirectional_columnar_dim_and_fact_with_cardinality(
            Cardinality::OneToMany,
        );
    let engine = DaxEngine::new();

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
fn distinctcount_virtual_blank_row_respects_bidirectional_filter_context_columnar_fact() {
    let model = build_model_blank_row_bidirectional_columnar_fact();
    let engine = DaxEngine::new();

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

#[test]
fn distinctcount_virtual_blank_row_respects_bidirectional_filter_context_columnar_dim() {
    let model =
        build_model_blank_row_bidirectional_columnar_dim_with_cardinality(Cardinality::OneToMany);
    let engine = DaxEngine::new();

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

#[test]
fn distinctcount_virtual_blank_row_respects_bidirectional_filter_context_columnar_dim_and_fact() {
    let model =
        build_model_blank_row_bidirectional_columnar_dim_and_fact_with_cardinality(
            Cardinality::OneToMany,
        );
    let engine = DaxEngine::new();

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

#[test]
fn summarizecolumns_virtual_blank_row_respects_bidirectional_filter_context_columnar_fact() {
    let model = build_model_blank_row_bidirectional_columnar_fact();
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn summarizecolumns_virtual_blank_row_respects_bidirectional_filter_context_columnar_dim() {
    let model =
        build_model_blank_row_bidirectional_columnar_dim_with_cardinality(Cardinality::OneToMany);
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn summarizecolumns_virtual_blank_row_respects_bidirectional_filter_context_columnar_dim_and_fact() {
    let model =
        build_model_blank_row_bidirectional_columnar_dim_and_fact_with_cardinality(
            Cardinality::OneToMany,
        );
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn values_virtual_blank_row_respects_bidirectional_filter_context_many_to_many_columnar_fact() {
    let model = build_model_blank_row_bidirectional_columnar_fact_with_cardinality(Cardinality::ManyToMany);
    let engine = DaxEngine::new();

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
fn distinctcount_virtual_blank_row_respects_bidirectional_filter_context_many_to_many_columnar_fact() {
    let model = build_model_blank_row_bidirectional_columnar_fact_with_cardinality(Cardinality::ManyToMany);
    let engine = DaxEngine::new();

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

#[test]
fn distinctcount_virtual_blank_row_respects_bidirectional_filter_context_many_to_many_columnar_dim() {
    let model =
        build_model_blank_row_bidirectional_columnar_dim_with_cardinality(Cardinality::ManyToMany);
    let engine = DaxEngine::new();

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

#[test]
fn distinctcount_virtual_blank_row_respects_bidirectional_filter_context_many_to_many_columnar_dim_and_fact(
) {
    let model = build_model_blank_row_bidirectional_columnar_dim_and_fact_with_cardinality(
        Cardinality::ManyToMany,
    );
    let engine = DaxEngine::new();

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

#[test]
fn summarizecolumns_virtual_blank_row_respects_bidirectional_filter_context_many_to_many_columnar_fact() {
    let model =
        build_model_blank_row_bidirectional_columnar_fact_with_cardinality(Cardinality::ManyToMany);
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn summarizecolumns_virtual_blank_row_respects_bidirectional_filter_context_many_to_many_columnar_dim(
) {
    let model =
        build_model_blank_row_bidirectional_columnar_dim_with_cardinality(Cardinality::ManyToMany);
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn summarizecolumns_virtual_blank_row_respects_bidirectional_filter_context_many_to_many_columnar_dim_and_fact(
) {
    let model = build_model_blank_row_bidirectional_columnar_dim_and_fact_with_cardinality(
        Cardinality::ManyToMany,
    );
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
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
                "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
                &unmatched_filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}
