use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, FilterContext, Relationship, Table, Value,
};

fn build_model() -> DataModel {
    let mut model = DataModel::new();

    let mut customers = Table::new("Customers", vec!["CustomerId", "Name", "Region"]);
    customers
        .push_row(vec![1.into(), "Alice".into(), "East".into()])
        .unwrap();
    customers
        .push_row(vec![2.into(), "Bob".into(), "West".into()])
        .unwrap();
    customers
        .push_row(vec![3.into(), "Carol".into(), "East".into()])
        .unwrap();
    model.add_table(customers).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "CustomerId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 1.into(), 20.0.into()])
        .unwrap();
    orders
        .push_row(vec![102.into(), 2.into(), 5.0.into()])
        .unwrap();
    orders
        .push_row(vec![103.into(), 3.into(), 8.0.into()])
        .unwrap();
    model.add_table(orders).unwrap();

    // Add a relationship using mismatched casing for tables and columns.
    model
        .add_relationship(Relationship {
            name: "orders_customers".into(),
            from_table: "orders".into(),
            from_column: "customerid".into(),
            to_table: "CUSTOMERS".into(),
            to_column: "CUSTOMERID".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
}

#[test]
fn identifiers_are_case_insensitive_for_measures_columns_filters_and_relationships() {
    let mut model = build_model();

    model.add_measure("Total Sales (lower)", "sum(orders[amount])")
        .unwrap();
    model.add_measure("Total Sales (upper)", "SUM(orders[amount])")
        .unwrap();

    let total = model
        .evaluate_measure("total sales (lower)", &FilterContext::empty())
        .unwrap();
    assert_eq!(total, Value::from(43.0));

    let east_filter =
        FilterContext::empty().with_column_equals("customers", "region", "East".into());

    let east_total = model
        .evaluate_measure("TOTAL SALES (LOWER)", &east_filter)
        .unwrap();
    assert_eq!(east_total, Value::from(38.0));

    let east_total_upper = model
        .evaluate_measure("[total sales (upper)]", &east_filter)
        .unwrap();
    assert_eq!(east_total_upper, Value::from(38.0));

    // LOOKUPVALUE compares table identifiers internally; it should be case-insensitive too.
    model
        .add_measure(
            "Customer 1 Name",
            "LOOKUPVALUE(Customers[Name], customers[customerid], 1)",
        )
        .unwrap();
    let customer_1 = model
        .evaluate_measure("customer 1 name", &FilterContext::empty())
        .unwrap();
    assert_eq!(customer_1, Value::from("Alice"));

    // CROSSFILTER direction matching should be case-insensitive for table/column identifiers.
    model
        .add_measure(
            "Customers With Large Orders",
            "CALCULATE(COUNTROWS(customers), CROSSFILTER(orders[customerid], CUSTOMERS[CUSTOMERID], ONEWAY_LEFTFILTERSRIGHT), orders[amount] > 10)",
        )
        .unwrap();
    let customers_with_large_orders = model
        .evaluate_measure("customers with large orders", &FilterContext::empty())
        .unwrap();
    assert_eq!(customers_with_large_orders, Value::from(1_i64));
}
