use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, Relationship, RowContext,
    Table, Value,
};
use pretty_assertions::assert_eq;

fn build_model_blank_row_single_direction() -> DataModel {
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
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
}

fn build_model_blank_row_single_direction_columnar_fact() -> DataModel {
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
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
}

fn build_model_blank_row_single_direction_columnar_dim() -> DataModel {
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
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
}

fn build_model_blank_row_single_direction_columnar_dim_and_fact() -> DataModel {
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
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
}

fn assert_crossfilter_override_enables_bidirectional_blank_member(model: &DataModel) {
    let engine = DaxEngine::new();

    let expr_values = "CALCULATE(COUNTROWS(VALUES(Customers[Region])), CROSSFILTER(Orders[CustomerId], Customers[CustomerId], \"BOTH\"))";
    let expr_selected =
        "CALCULATE(SELECTEDVALUE(Customers[Region]), CROSSFILTER(Orders[CustomerId], Customers[CustomerId], \"BOTH\"))";

    // Without any extra filters, the unknown member exists due to the unmatched fact key.
    assert_eq!(
        engine
            .evaluate(model, expr_values, &FilterContext::empty(), &RowContext::default())
            .unwrap(),
        3.into()
    );

    // Fact filtered to a matched key: dimension should be restricted to that value (no BLANK).
    let matched_filter = FilterContext::empty().with_column_equals("Orders", "CustomerId", 1.into());
    assert_eq!(
        engine
            .evaluate(model, expr_values, &matched_filter, &RowContext::default())
            .unwrap(),
        1.into()
    );
    assert_eq!(
        engine
            .evaluate(model, expr_selected, &matched_filter, &RowContext::default())
            .unwrap(),
        Value::from("East")
    );

    // Fact filtered to an unmatched key: only the blank/unknown member remains.
    let unmatched_filter =
        FilterContext::empty().with_column_equals("Orders", "CustomerId", 999.into());
    assert_eq!(
        engine
            .evaluate(model, expr_values, &unmatched_filter, &RowContext::default())
            .unwrap(),
        1.into()
    );
    assert_eq!(
        engine
            .evaluate(model, expr_selected, &unmatched_filter, &RowContext::default())
            .unwrap(),
        Value::Blank
    );
}

#[test]
fn crossfilter_override_enables_bidirectional_blank_member_visibility() {
    let model = build_model_blank_row_single_direction();
    assert_crossfilter_override_enables_bidirectional_blank_member(&model);
}

#[test]
fn crossfilter_override_enables_bidirectional_blank_member_visibility_for_columnar_fact() {
    let model = build_model_blank_row_single_direction_columnar_fact();
    assert_crossfilter_override_enables_bidirectional_blank_member(&model);
}

#[test]
fn crossfilter_override_enables_bidirectional_blank_member_visibility_for_columnar_dim() {
    let model = build_model_blank_row_single_direction_columnar_dim();
    assert_crossfilter_override_enables_bidirectional_blank_member(&model);
}

#[test]
fn crossfilter_override_enables_bidirectional_blank_member_visibility_for_columnar_dim_and_fact() {
    let model = build_model_blank_row_single_direction_columnar_dim_and_fact();
    assert_crossfilter_override_enables_bidirectional_blank_member(&model);
}

