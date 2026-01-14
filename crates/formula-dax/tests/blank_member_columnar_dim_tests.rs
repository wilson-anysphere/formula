use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, Relationship, RowContext,
    Table, Value,
};
use pretty_assertions::assert_eq;

fn build_model_blank_member_columnar_dim() -> DataModel {
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

    let mut orders = Table::new("Orders", vec!["OrderId", "CustomerId", "Amount"]);
    orders.push_row(vec![100.into(), 1.into(), 10.0.into()]).unwrap();
    orders.push_row(vec![101.into(), 999.into(), 7.0.into()]).unwrap();
    model.add_table(orders).unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders_Customers".into(),
            from_table: "Orders".into(),
            from_column: "CustomerId".into(),
            to_table: "Customers".into(),
            to_column: "CustomerId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model.add_measure("Total Sales", "SUM(Orders[Amount])").unwrap();

    model
}

fn build_model_blank_member_columnar_dim_and_fact() -> DataModel {
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
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut orders = ColumnarTableBuilder::new(orders_schema, options);
    orders.append_row(&[
        formula_columnar::Value::Number(100.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    orders.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(999.0),
        formula_columnar::Value::Number(7.0),
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
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model.add_measure("Total Sales", "SUM(Orders[Amount])").unwrap();

    model
}

fn assert_blank_member_semantics(model: &DataModel) {
    let engine = DaxEngine::new();
    let empty = FilterContext::empty();

    // Ensure the relationship-generated blank/unknown member is visible in unfiltered contexts.
    assert_eq!(
        engine
            .evaluate(
                model,
                "COUNTROWS(VALUES(Customers[Region]))",
                &empty,
                &RowContext::default(),
            )
            .unwrap(),
        3.into()
    );
    assert_eq!(
        engine
            .evaluate(
                model,
                "DISTINCTCOUNT(Customers[Region])",
                &empty,
                &RowContext::default(),
            )
            .unwrap(),
        3.into()
    );
    assert_eq!(
        engine
            .evaluate(
                model,
                "DISTINCTCOUNTNOBLANK(Customers[Region])",
                &empty,
                &RowContext::default(),
            )
            .unwrap(),
        2.into()
    );
    assert_eq!(
        engine
            .evaluate(
                model,
                "COUNTBLANK(Customers[Region])",
                &empty,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
    assert_eq!(
        engine
            .evaluate(
                model,
                "COUNTA(Customers[Region])",
                &empty,
                &RowContext::default(),
            )
            .unwrap(),
        2.into()
    );
    assert_eq!(
        engine
            .evaluate(
                model,
                "COUNT(Customers[Region])",
                &empty,
                &RowContext::default(),
            )
            .unwrap(),
        0.into()
    );
    assert_eq!(
        engine
            .evaluate(
                model,
                "COUNTROWS(ALL(Customers[Region]))",
                &empty,
                &RowContext::default(),
            )
            .unwrap(),
        3.into()
    );
    assert_eq!(
        engine
            .evaluate(
                model,
                "COUNTROWS(ALLNOBLANKROW(Customers[Region]))",
                &empty,
                &RowContext::default(),
            )
            .unwrap(),
        2.into()
    );

    // Filtering to the blank member should include unmatched fact rows for measures.
    let blank_region = FilterContext::empty().with_column_equals("Customers", "Region", Value::Blank);
    assert_eq!(
        engine
            .evaluate(
                model,
                "COUNTBLANK(Customers[Region])",
                &blank_region,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
    assert_eq!(
        engine
            .evaluate(
                model,
                "COUNTA(Customers[Region])",
                &blank_region,
                &RowContext::default(),
            )
            .unwrap(),
        0.into()
    );
    assert_eq!(
        model.evaluate_measure("Total Sales", &blank_region).unwrap(),
        7.0.into()
    );

    let non_blank_regions = FilterContext::empty().with_column_in(
        "Customers",
        "Region",
        vec!["East".into(), "West".into()],
    );
    assert_eq!(
        engine
            .evaluate(
                model,
                "COUNTBLANK(Customers[Region])",
                &non_blank_regions,
                &RowContext::default(),
            )
            .unwrap(),
        0.into()
    );
    assert_eq!(
        engine
            .evaluate(
                model,
                "COUNTA(Customers[Region])",
                &non_blank_regions,
                &RowContext::default(),
            )
            .unwrap(),
        2.into()
    );

    // Filtering the dimension to non-blank values should exclude the unmatched fact rows (the
    // relationship-generated blank member itself is filtered out).
    assert_eq!(
        engine
            .evaluate(
                model,
                "CALCULATE(COUNTROWS(Orders), Customers[Region] <> BLANK())",
                &empty,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
    assert_eq!(
        engine
            .evaluate(
                model,
                "CALCULATE([Total Sales], Customers[Region] <> BLANK())",
                &empty,
                &RowContext::default(),
            )
            .unwrap(),
        10.0.into()
    );
}

#[test]
fn blank_member_semantics_work_for_columnar_dimension_tables() {
    let model = build_model_blank_member_columnar_dim();
    assert_blank_member_semantics(&model);
}

#[test]
fn blank_member_semantics_work_for_columnar_dimension_and_fact_tables() {
    let model = build_model_blank_member_columnar_dim_and_fact();
    assert_blank_member_semantics(&model);
}
