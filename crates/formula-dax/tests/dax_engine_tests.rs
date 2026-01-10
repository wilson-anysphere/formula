use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, DaxError, FilterContext, Relationship,
    RowContext, Table, Value,
};
use pretty_assertions::assert_eq;

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
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
}

#[test]
fn relationship_enforces_referential_integrity() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Id"]);
    dim.push_row(vec![1.into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id"]);
    fact.push_row(vec![2.into()]).unwrap();
    model.add_table(fact).unwrap();

    let err = model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Id".into(),
            to_table: "Dim".into(),
            to_column: "Id".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap_err();

    let message = err.to_string();
    assert!(message.contains("referential integrity violation"));
}

#[test]
fn calculated_column_can_use_related() {
    let mut model = build_model();
    model
        .add_calculated_column("Orders", "CustomerName", "RELATED(Customers[Name])")
        .unwrap();

    let orders = model.table("Orders").unwrap();
    let values: Vec<Value> = (0..orders.row_count())
        .map(|row| orders.value(row, "CustomerName").unwrap().clone())
        .collect();

    assert_eq!(
        values,
        vec![
            Value::from("Alice"),
            Value::from("Alice"),
            Value::from("Bob"),
            Value::from("Carol")
        ]
    );
}

#[test]
fn bracket_identifier_resolves_to_column_in_row_context() {
    let mut model = build_model();
    model
        .add_calculated_column("Orders", "Double Amount", "[Amount] * 2")
        .unwrap();

    let orders = model.table("Orders").unwrap();
    let values: Vec<Value> = (0..orders.row_count())
        .map(|row| orders.value(row, "Double Amount").unwrap().clone())
        .collect();

    assert_eq!(
        values,
        vec![20.0.into(), 40.0.into(), 10.0.into(), 16.0.into()]
    );
}

#[test]
fn measure_respects_filter_propagation_across_relationships() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let total = model
        .evaluate_measure("Total Sales", &FilterContext::empty())
        .unwrap();
    assert_eq!(total, Value::from(43.0));

    let east_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    let east_total = model.evaluate_measure("Total Sales", &east_filter).unwrap();
    assert_eq!(east_total, Value::from(38.0));
}

#[test]
fn calculate_overrides_existing_column_filters() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    model
        .add_measure(
            "East Sales",
            "CALCULATE([Total Sales], Customers[Region] = \"East\")",
        )
        .unwrap();

    let west_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "West".into());
    let value = model.evaluate_measure("East Sales", &west_filter).unwrap();
    assert_eq!(value, Value::from(38.0));
}

#[test]
fn relatedtable_supports_iterators() {
    let mut model = build_model();
    model
        .add_calculated_column(
            "Customers",
            "Customer Sales",
            "SUMX(RELATEDTABLE(Orders), Orders[Amount])",
        )
        .unwrap();

    let customers = model.table("Customers").unwrap();
    let values: Vec<Value> = (0..customers.row_count())
        .map(|row| customers.value(row, "Customer Sales").unwrap().clone())
        .collect();

    assert_eq!(values, vec![30.0.into(), 5.0.into(), 8.0.into()]);
}

#[test]
fn calculate_transitions_row_context_to_filter_context_for_measures() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    model
        .add_calculated_column(
            "Customers",
            "Sales via CALCULATE",
            "CALCULATE([Total Sales])",
        )
        .unwrap();

    let customers = model.table("Customers").unwrap();
    let values: Vec<Value> = (0..customers.row_count())
        .map(|row| customers.value(row, "Sales via CALCULATE").unwrap().clone())
        .collect();

    assert_eq!(values, vec![30.0.into(), 5.0.into(), 8.0.into()]);
}

#[test]
fn insert_row_checks_referential_integrity_and_key_uniqueness() {
    let mut model = build_model();

    let err = model
        .insert_row("Orders", vec![104.into(), 999.into(), 1.0.into()])
        .unwrap_err();
    assert!(matches!(
        err,
        DaxError::ReferentialIntegrityViolation { .. }
    ));

    let err = model
        .insert_row(
            "Customers",
            vec![1.into(), "Duplicate".into(), "East".into()],
        )
        .unwrap_err();
    assert!(matches!(err, DaxError::NonUniqueKey { .. }));
}

#[test]
fn calculate_context_transition_respects_existing_filter_context() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let filter = FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    let mut row_ctx = RowContext::default();
    row_ctx.push("Customers", 1); // Bob (West)

    let value = DaxEngine::new()
        .evaluate(&model, "CALCULATE([Total Sales])", &filter, &row_ctx)
        .unwrap();

    assert_eq!(value, Value::Blank);
}

#[test]
fn insert_row_computes_calculated_columns() {
    let mut model = build_model();
    model
        .add_calculated_column("Orders", "CustomerName", "RELATED(Customers[Name])")
        .unwrap();

    model
        .insert_row("Orders", vec![104.into(), 1.into(), 7.0.into()])
        .unwrap();

    let orders = model.table("Orders").unwrap();
    assert_eq!(orders.row_count(), 5);
    assert_eq!(
        orders.value(4, "CustomerName").unwrap(),
        &Value::from("Alice")
    );
}

#[test]
fn insert_row_updates_relationship_key_index() {
    let mut model = build_model();
    model
        .add_calculated_column("Orders", "CustomerName", "RELATED(Customers[Name])")
        .unwrap();

    model
        .insert_row("Customers", vec![4.into(), "Dan".into(), "East".into()])
        .unwrap();

    model
        .insert_row("Orders", vec![104.into(), 4.into(), 12.0.into()])
        .unwrap();

    let orders = model.table("Orders").unwrap();
    assert_eq!(orders.row_count(), 5);
    assert_eq!(
        orders.value(4, "CustomerName").unwrap(),
        &Value::from("Dan")
    );
}

#[test]
fn countrows_counts_relatedtable_rows() {
    let mut model = build_model();
    model
        .add_calculated_column(
            "Customers",
            "Order Count",
            "COUNTROWS(RELATEDTABLE(Orders))",
        )
        .unwrap();

    let customers = model.table("Customers").unwrap();
    let values: Vec<Value> = (0..customers.row_count())
        .map(|row| customers.value(row, "Order Count").unwrap().clone())
        .collect();

    assert_eq!(values, vec![2.into(), 1.into(), 1.into()]);
}

#[test]
fn countx_respects_filter_propagation() {
    let mut model = build_model();
    model
        .add_measure("Order Count", "COUNTX(Orders, Orders[OrderId])")
        .unwrap();

    let east_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    let east_count = model.evaluate_measure("Order Count", &east_filter).unwrap();
    assert_eq!(east_count, 3.into());

    let empty_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "Nowhere".into());
    let empty_count = model
        .evaluate_measure("Order Count", &empty_filter)
        .unwrap();
    assert_eq!(empty_count, 0.into());
}

#[test]
fn maxx_iterates_relatedtable() {
    let mut model = build_model();
    model
        .add_calculated_column(
            "Customers",
            "Max Order Amount",
            "MAXX(RELATEDTABLE(Orders), Orders[Amount])",
        )
        .unwrap();

    let customers = model.table("Customers").unwrap();
    let values: Vec<Value> = (0..customers.row_count())
        .map(|row| customers.value(row, "Max Order Amount").unwrap().clone())
        .collect();

    assert_eq!(values, vec![20.0.into(), 5.0.into(), 8.0.into()]);
}

#[test]
fn max_respects_filter_propagation() {
    let mut model = build_model();
    model
        .add_measure("Max Sale", "MAX(Orders[Amount])")
        .unwrap();

    let west_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "West".into());
    let west_max = model.evaluate_measure("Max Sale", &west_filter).unwrap();
    assert_eq!(west_max, 5.0.into());

    let empty_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "Nowhere".into());
    let empty_max = model.evaluate_measure("Max Sale", &empty_filter).unwrap();
    assert_eq!(empty_max, Value::Blank);
}

#[test]
fn calculate_supports_column_comparisons() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    model
        .add_measure("Big Sales", "CALCULATE([Total Sales], Orders[Amount] > 10)")
        .unwrap();

    let value = model
        .evaluate_measure("Big Sales", &FilterContext::empty())
        .unwrap();
    assert_eq!(value, 20.0.into());

    let east_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    let east_value = model.evaluate_measure("Big Sales", &east_filter).unwrap();
    assert_eq!(east_value, 20.0.into());

    let west_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "West".into());
    let west_value = model.evaluate_measure("Big Sales", &west_filter).unwrap();
    assert_eq!(west_value, Value::Blank);
}

#[test]
fn if_works_in_calculated_columns() {
    let mut model = build_model();
    model
        .add_calculated_column(
            "Customers",
            "IsEast",
            "IF(Customers[Region] = \"East\", 1, 0)",
        )
        .unwrap();

    let customers = model.table("Customers").unwrap();
    let values: Vec<Value> = (0..customers.row_count())
        .map(|row| customers.value(row, "IsEast").unwrap().clone())
        .collect();

    assert_eq!(values, vec![1.into(), 0.into(), 1.into()]);
}

#[test]
fn divide_supports_safe_division_and_blanks() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    model
        .add_measure("Average Sale", "DIVIDE([Total Sales], COUNTROWS(Orders))")
        .unwrap();

    let east_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    let east_avg = model
        .evaluate_measure("Average Sale", &east_filter)
        .unwrap();
    assert_eq!(east_avg, (38.0 / 3.0).into());

    let empty_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "Nowhere".into());
    let empty_avg = model
        .evaluate_measure("Average Sale", &empty_filter)
        .unwrap();
    assert_eq!(empty_avg, Value::Blank);
}

#[test]
fn coalesce_returns_first_non_blank_value() {
    let model = build_model();
    let value = DaxEngine::new()
        .evaluate(
            &model,
            "COALESCE(BLANK(), BLANK(), 7)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 7.into());
}

#[test]
fn selectedvalue_and_hasonevalue_use_filter_context() {
    let mut model = build_model();
    model
        .add_measure("Selected Region", "SELECTEDVALUE(Customers[Region])")
        .unwrap();
    model
        .add_measure(
            "Selected Region (fallback)",
            "SELECTEDVALUE(Customers[Region], \"Multiple\")",
        )
        .unwrap();
    model
        .add_measure("Has One Region", "HASONEVALUE(Customers[Region])")
        .unwrap();
    model
        .add_measure("Region Count", "DISTINCTCOUNT(Customers[Region])")
        .unwrap();

    let east_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    assert_eq!(
        model
            .evaluate_measure("Selected Region", &east_filter)
            .unwrap(),
        Value::from("East")
    );
    assert_eq!(
        model
            .evaluate_measure("Selected Region (fallback)", &east_filter)
            .unwrap(),
        Value::from("East")
    );
    assert_eq!(
        model
            .evaluate_measure("Has One Region", &east_filter)
            .unwrap(),
        Value::from(true)
    );
    assert_eq!(
        model
            .evaluate_measure("Region Count", &east_filter)
            .unwrap(),
        Value::from(1)
    );

    assert_eq!(
        model
            .evaluate_measure("Selected Region", &FilterContext::empty())
            .unwrap(),
        Value::Blank
    );
    assert_eq!(
        model
            .evaluate_measure("Selected Region (fallback)", &FilterContext::empty())
            .unwrap(),
        Value::from("Multiple")
    );
    assert_eq!(
        model
            .evaluate_measure("Has One Region", &FilterContext::empty())
            .unwrap(),
        Value::from(false)
    );
    assert_eq!(
        model
            .evaluate_measure("Region Count", &FilterContext::empty())
            .unwrap(),
        Value::from(2)
    );

    let empty_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "Nowhere".into());
    assert_eq!(
        model
            .evaluate_measure("Selected Region", &empty_filter)
            .unwrap(),
        Value::Blank
    );
    assert_eq!(
        model
            .evaluate_measure("Selected Region (fallback)", &empty_filter)
            .unwrap(),
        Value::from("Multiple")
    );
    assert_eq!(
        model
            .evaluate_measure("Has One Region", &empty_filter)
            .unwrap(),
        Value::from(false)
    );
    assert_eq!(
        model
            .evaluate_measure("Region Count", &empty_filter)
            .unwrap(),
        Value::from(0)
    );
}
