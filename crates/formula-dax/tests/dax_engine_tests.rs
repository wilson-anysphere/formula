mod common;

use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions,
};
use formula_dax::{
    pivot, Cardinality, CrossFilterDirection, DataModel, DaxEngine, DaxError, FilterContext,
    GroupByColumn, PivotMeasure, Relationship, RowContext, Table, Value,
};
use pretty_assertions::assert_eq;
use std::sync::Arc;

use common::{build_model, build_model_bidirectional};

fn build_model_without_relationship() -> DataModel {
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
        .map(|row| orders.value(row, "CustomerName").unwrap())
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
        .map(|row| orders.value(row, "Double Amount").unwrap())
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
fn filter_context_supports_multi_value_column_filters() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let empty_total = model
        .evaluate_measure("Total Sales", &FilterContext::empty())
        .unwrap();

    // Customers[Region] IN {"East","West"} includes all rows in the dimension table, so it should
    // yield the same result as an empty filter context.
    let all_regions_filter = FilterContext::empty().with_column_in(
        "Customers",
        "Region",
        [Value::from("East"), Value::from("West")],
    );
    let all_regions_total = model.evaluate_measure("Total Sales", &all_regions_filter).unwrap();
    assert_eq!(all_regions_total, empty_total);

    let east_filter =
        FilterContext::empty().with_column_in("Customers", "Region", [Value::from("East")]);
    let east_total = model.evaluate_measure("Total Sales", &east_filter).unwrap();
    assert_eq!(east_total, Value::from(38.0));
}

#[test]
fn var_return_works_in_measures() {
    let mut model = build_model();
    model
        .add_measure(
            "Double Total Sales",
            "VAR t = SUM(Orders[Amount]) RETURN t * 2",
        )
        .unwrap();

    let value = model
        .evaluate_measure("Double Total Sales", &FilterContext::empty())
        .unwrap();
    assert_eq!(value, 86.0.into());
}

#[test]
fn var_return_works_in_calculated_columns_in_row_context() {
    let mut model = build_model();
    model
        .add_calculated_column(
            "Orders",
            "Double Amount via Var",
            "VAR x = Orders[Amount] RETURN x * 2",
        )
        .unwrap();

    let orders = model.table("Orders").unwrap();
    let values: Vec<Value> = (0..orders.row_count())
        .map(|row| orders.value(row, "Double Amount via Var").unwrap())
        .collect();
    assert_eq!(values, vec![20.0.into(), 40.0.into(), 10.0.into(), 16.0.into()]);
}

#[test]
fn nested_var_scopes_shadow_outer_bindings() {
    let model = build_model();
    let value = DaxEngine::new()
        .evaluate(
            &model,
            "VAR x = 1 RETURN VAR x = x + 1 RETURN x * 2",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 4.0.into());
}

#[test]
fn var_is_visible_in_calculate_expression_and_filter_arguments() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    model
        .add_measure(
            "East Sales via Var",
            "VAR region = \"East\" RETURN CALCULATE([Total Sales], Customers[Region] = region)",
        )
        .unwrap();

    let value = model
        .evaluate_measure("East Sales via Var", &FilterContext::empty())
        .unwrap();
    assert_eq!(value, 38.0.into());
}

#[test]
fn calculate_can_reference_scalar_var_in_expression_argument() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    model
        .add_measure(
            "Total Sales (constant via var)",
            "VAR t = [Total Sales] RETURN CALCULATE(t, Customers[Region] = \"East\")",
        )
        .unwrap();

    let value = model
        .evaluate_measure("Total Sales (constant via var)", &FilterContext::empty())
        .unwrap();
    assert_eq!(value, 43.0.into());
}

#[test]
fn var_is_visible_in_iterator_body() {
    let model = build_model();
    let value = DaxEngine::new()
        .evaluate(
            &model,
            "VAR factor = 2 RETURN SUMX(Orders, Orders[Amount] * factor)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 86.0.into());
}

#[test]
fn table_vars_can_be_used_as_calculate_filter_arguments() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let value = DaxEngine::new()
        .evaluate(
            &model,
            "VAR t = FILTER(Orders, Orders[Amount] > 10) RETURN CALCULATE([Total Sales], t)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 20.0.into());
}

#[test]
fn table_vars_can_be_used_as_iterator_table_arguments() {
    let model = build_model();
    let value = DaxEngine::new()
        .evaluate(
            &model,
            "VAR t = FILTER(Orders, Orders[Amount] > 10) RETURN SUMX(t, Orders[Amount])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 20.0.into());
}

#[test]
fn vars_shadow_table_names_in_scalar_context() {
    let model = build_model();
    let value = DaxEngine::new()
        .evaluate(
            &model,
            "VAR Orders = 3 RETURN Orders + 1",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 4.0.into());
}

#[test]
fn vars_shadow_table_names_in_table_context() {
    let model = build_model();
    let value = DaxEngine::new()
        .evaluate(
            &model,
            "VAR Orders = FILTER(Orders, Orders[Amount] > 10) RETURN COUNTROWS(Orders)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 1.into());
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
fn calculate_keepfilters_intersects_with_existing_column_filters() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    model
        .add_measure(
            "Override",
            "CALCULATE([Total Sales], Customers[Region] = \"West\")",
        )
        .unwrap();
    model
        .add_measure(
            "Keep",
            "CALCULATE([Total Sales], KEEPFILTERS(Customers[Region] = \"West\"))",
        )
        .unwrap();

    let east_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "East".into());

    // Default CALCULATE filter arguments replace existing filters on the same column.
    assert_eq!(
        model.evaluate_measure("Override", &east_filter).unwrap(),
        5.0.into()
    );

    // KEEPFILTERS forces intersection with the existing filter context.
    assert_eq!(
        model.evaluate_measure("Keep", &east_filter).unwrap(),
        Value::Blank
    );
}

#[test]
fn calculate_keepfilters_on_different_column_applies_both_filters() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    model
        .add_measure(
            "Keep Alice",
            "CALCULATE([Total Sales], KEEPFILTERS(Customers[Name] = \"Alice\"))",
        )
        .unwrap();

    let east_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    assert_eq!(
        model.evaluate_measure("Keep Alice", &east_filter).unwrap(),
        30.0.into()
    );
}

#[test]
fn calculate_keepfilters_preserves_existing_filters_for_boolean_expressions() {
    let mut model = build_model();
    model
        .add_measure(
            "Override Medium",
            "CALCULATE(SUM(Orders[Amount]), Orders[Amount] > 7 && Orders[Amount] < 20)",
        )
        .unwrap();
    model
        .add_measure(
            "Keep Medium",
            "CALCULATE(SUM(Orders[Amount]), KEEPFILTERS(Orders[Amount] > 7 && Orders[Amount] < 20))",
        )
        .unwrap();

    let amount_20 =
        FilterContext::empty().with_column_equals("Orders", "Amount", Value::from(20.0));

    // Default CALCULATE filter arguments replace existing filters on referenced columns.
    assert_eq!(
        model.evaluate_measure("Override Medium", &amount_20).unwrap(),
        18.0.into()
    );

    // KEEPFILTERS forces intersection with the existing filter context.
    assert_eq!(
        model.evaluate_measure("Keep Medium", &amount_20).unwrap(),
        Value::Blank
    );
}

#[test]
fn calculate_keepfilters_on_table_expression_preserves_existing_table_filters() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    model
        .add_measure(
            "Override East Customers",
            "CALCULATE([Total Sales], FILTER(ALL(Customers), Customers[Region] = \"East\"))",
        )
        .unwrap();
    model
        .add_measure(
            "Keep East Customers",
            "CALCULATE([Total Sales], KEEPFILTERS(FILTER(ALL(Customers), Customers[Region] = \"East\")))",
        )
        .unwrap();

    let alice_filter =
        FilterContext::empty().with_column_equals("Customers", "Name", "Alice".into());

    // The table expression targets Customers. Without KEEPFILTERS it replaces the existing
    // Customers[Name] filter, expanding from Alice-only to all East customers.
    assert_eq!(
        model
            .evaluate_measure("Override East Customers", &alice_filter)
            .unwrap(),
        38.0.into()
    );

    // With KEEPFILTERS, the row set from the table expression is intersected with the existing
    // Customers filters, preserving Customers[Name] = \"Alice\".
    assert_eq!(
        model
            .evaluate_measure("Keep East Customers", &alice_filter)
            .unwrap(),
        30.0.into()
    );
}

#[test]
fn calculate_treatas_can_simulate_relationships() {
    let model = build_model_without_relationship();
    let filter = FilterContext::empty().with_column_equals("Customers", "Region", "East".into());

    let value = DaxEngine::new()
        .evaluate(
            &model,
            "CALCULATE(SUM(Orders[Amount]), TREATAS(VALUES(Customers[CustomerId]), Orders[CustomerId]))",
            &filter,
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, 38.0.into());
}

#[test]
fn calculate_treatas_with_empty_source_values_filters_to_empty_set() {
    let model = build_model_without_relationship();
    let filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "Nowhere".into());

    let engine = DaxEngine::new();
    let sum_value = engine
        .evaluate(
            &model,
            "CALCULATE(SUM(Orders[Amount]), TREATAS(VALUES(Customers[CustomerId]), Orders[CustomerId]))",
            &filter,
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(sum_value, Value::Blank);

    let count_value = engine
        .evaluate(
            &model,
            "CALCULATE(COUNTROWS(Orders), TREATAS(VALUES(Customers[CustomerId]), Orders[CustomerId]))",
            &filter,
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(count_value, 0.into());
}

#[test]
fn calculate_keepfilters_with_treatas_intersects_existing_target_filter() {
    let model = build_model_without_relationship();
    let mut filter = FilterContext::empty();
    filter.set_column_equals("Customers", "Region", "East".into());
    filter.set_column_equals("Orders", "CustomerId", 2.into());

    let engine = DaxEngine::new();

    let override_value = engine
        .evaluate(
            &model,
            "CALCULATE(SUM(Orders[Amount]), TREATAS(VALUES(Customers[CustomerId]), Orders[CustomerId]))",
            &filter,
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(override_value, 38.0.into());

    let keep_value = engine
        .evaluate(
            &model,
            "CALCULATE(SUM(Orders[Amount]), KEEPFILTERS(TREATAS(VALUES(Customers[CustomerId]), Orders[CustomerId])))",
            &filter,
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(keep_value, Value::Blank);
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
        .map(|row| customers.value(row, "Customer Sales").unwrap())
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
        .map(|row| customers.value(row, "Sales via CALCULATE").unwrap())
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
        Value::from("Alice")
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
    assert_eq!(orders.value(4, "CustomerName").unwrap(), Value::from("Dan"));
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
        .map(|row| customers.value(row, "Order Count").unwrap())
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
fn count_counta_and_countblank_respect_filter_context() {
    let mut model = DataModel::new();

    let mut t = Table::new("T", vec!["Group", "Value"]);
    t.push_row(vec!["A".into(), 1.into()]).unwrap();
    t.push_row(vec!["A".into(), Value::Blank]).unwrap();
    t.push_row(vec!["B".into(), 2.into()]).unwrap();
    t.push_row(vec!["B".into(), Value::Blank]).unwrap();
    t.push_row(vec!["B".into(), 3.into()]).unwrap();
    model.add_table(t).unwrap();

    model.add_measure("Count", "COUNT(T[Value])").unwrap();
    model.add_measure("CountA", "COUNTA(T[Value])").unwrap();
    model
        .add_measure("CountBlank", "COUNTBLANK(T[Value])")
        .unwrap();

    assert_eq!(
        model.evaluate_measure("Count", &FilterContext::empty()).unwrap(),
        3.into()
    );
    assert_eq!(
        model
            .evaluate_measure("CountA", &FilterContext::empty())
            .unwrap(),
        3.into()
    );
    assert_eq!(
        model
            .evaluate_measure("CountBlank", &FilterContext::empty())
            .unwrap(),
        2.into()
    );

    let group_a = FilterContext::empty().with_column_equals("T", "Group", "A".into());
    assert_eq!(model.evaluate_measure("Count", &group_a).unwrap(), 1.into());
    assert_eq!(model.evaluate_measure("CountA", &group_a).unwrap(), 1.into());
    assert_eq!(
        model.evaluate_measure("CountBlank", &group_a).unwrap(),
        1.into()
    );

    let group_b = FilterContext::empty().with_column_equals("T", "Group", "B".into());
    assert_eq!(model.evaluate_measure("Count", &group_b).unwrap(), 2.into());
    assert_eq!(model.evaluate_measure("CountA", &group_b).unwrap(), 2.into());
    assert_eq!(
        model.evaluate_measure("CountBlank", &group_b).unwrap(),
        1.into()
    );
}

#[test]
fn count_ignores_text_values() {
    let mut model = DataModel::new();

    let mut t = Table::new("T", vec!["Value"]);
    t.push_row(vec![1.into()]).unwrap();
    t.push_row(vec!["oops".into()]).unwrap();
    t.push_row(vec![Value::Blank]).unwrap();
    t.push_row(vec![2.into()]).unwrap();
    t.push_row(vec!["also text".into()]).unwrap();
    t.push_row(vec![Value::Blank]).unwrap();
    model.add_table(t).unwrap();

    model.add_measure("Count", "COUNT(T[Value])").unwrap();
    model.add_measure("CountA", "COUNTA(T[Value])").unwrap();
    model
        .add_measure("CountBlank", "COUNTBLANK(T[Value])")
        .unwrap();

    assert_eq!(
        model.evaluate_measure("Count", &FilterContext::empty()).unwrap(),
        2.into()
    );
    assert_eq!(
        model
            .evaluate_measure("CountA", &FilterContext::empty())
            .unwrap(),
        4.into()
    );
    assert_eq!(
        model
            .evaluate_measure("CountBlank", &FilterContext::empty())
            .unwrap(),
        2.into()
    );
}

#[test]
fn count_functions_work_for_columnar_tables() {
    let mut model = DataModel::new();

    let schema = vec![
        ColumnSchema {
            name: "Group".to_string(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "Value".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut t = ColumnarTableBuilder::new(schema, options);
    t.append_row(&[
        formula_columnar::Value::String(Arc::<str>::from("A")),
        formula_columnar::Value::Number(1.0),
    ]);
    t.append_row(&[
        formula_columnar::Value::String(Arc::<str>::from("A")),
        formula_columnar::Value::Null,
    ]);
    t.append_row(&[
        formula_columnar::Value::String(Arc::<str>::from("B")),
        formula_columnar::Value::Number(2.0),
    ]);
    t.append_row(&[
        formula_columnar::Value::String(Arc::<str>::from("B")),
        formula_columnar::Value::Null,
    ]);
    t.append_row(&[
        formula_columnar::Value::String(Arc::<str>::from("B")),
        formula_columnar::Value::Number(3.0),
    ]);
    model.add_table(Table::from_columnar("T", t.finalize())).unwrap();

    model.add_measure("Count", "COUNT(T[Value])").unwrap();
    model.add_measure("CountA", "COUNTA(T[Value])").unwrap();
    model
        .add_measure("CountBlank", "COUNTBLANK(T[Value])")
        .unwrap();

    assert_eq!(
        model.evaluate_measure("Count", &FilterContext::empty()).unwrap(),
        3.into()
    );
    assert_eq!(
        model
            .evaluate_measure("CountA", &FilterContext::empty())
            .unwrap(),
        3.into()
    );
    assert_eq!(
        model
            .evaluate_measure("CountBlank", &FilterContext::empty())
            .unwrap(),
        2.into()
    );

    let group_a = FilterContext::empty().with_column_equals("T", "Group", "A".into());
    assert_eq!(model.evaluate_measure("Count", &group_a).unwrap(), 1.into());
    assert_eq!(model.evaluate_measure("CountA", &group_a).unwrap(), 1.into());
    assert_eq!(
        model.evaluate_measure("CountBlank", &group_a).unwrap(),
        1.into()
    );

    let group_b = FilterContext::empty().with_column_equals("T", "Group", "B".into());
    assert_eq!(model.evaluate_measure("Count", &group_b).unwrap(), 2.into());
    assert_eq!(model.evaluate_measure("CountA", &group_b).unwrap(), 2.into());
    assert_eq!(
        model.evaluate_measure("CountBlank", &group_b).unwrap(),
        1.into()
    );
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
        .map(|row| customers.value(row, "Max Order Amount").unwrap())
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
fn calculate_supports_compound_boolean_and_filters() {
    let mut model = build_model();
    model
        .add_measure(
            "Medium Sales",
            "CALCULATE(SUM(Orders[Amount]), Orders[Amount] > 7 && Orders[Amount] < 20)",
        )
        .unwrap();

    let value = model
        .evaluate_measure("Medium Sales", &FilterContext::empty())
        .unwrap();
    assert_eq!(value, 18.0.into());
}

#[test]
fn calculate_compound_boolean_filters_respect_relationship_propagation() {
    let mut model = build_model();
    model
        .add_measure(
            "Range Sales",
            "CALCULATE(SUM(Orders[Amount]), Orders[Amount] > 4 && Orders[Amount] < 20)",
        )
        .unwrap();

    assert_eq!(
        model
            .evaluate_measure("Range Sales", &FilterContext::empty())
            .unwrap(),
        23.0.into()
    );

    let east_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    assert_eq!(
        model.evaluate_measure("Range Sales", &east_filter).unwrap(),
        18.0.into()
    );

    let west_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "West".into());
    assert_eq!(
        model.evaluate_measure("Range Sales", &west_filter).unwrap(),
        5.0.into()
    );
}

#[test]
fn calculate_supports_compound_boolean_or_filters() {
    let mut model = build_model();
    model
        .add_measure(
            "Selected Sales",
            "CALCULATE(SUM(Orders[Amount]), Orders[Amount] = 10 || Orders[Amount] = 20)",
        )
        .unwrap();

    let value = model
        .evaluate_measure("Selected Sales", &FilterContext::empty())
        .unwrap();
    assert_eq!(value, 30.0.into());
}

#[test]
fn calculate_supports_not_boolean_filters() {
    let mut model = build_model();
    model
        .add_measure(
            "Not Twenty",
            "CALCULATE(SUM(Orders[Amount]), NOT(Orders[Amount] = 20))",
        )
        .unwrap();

    let value = model
        .evaluate_measure("Not Twenty", &FilterContext::empty())
        .unwrap();
    assert_eq!(value, 23.0.into());
}

#[test]
fn calculate_supports_and_or_function_boolean_filters() {
    let mut model = build_model();
    model
        .add_measure(
            "Medium (AND fn)",
            "CALCULATE(SUM(Orders[Amount]), AND(Orders[Amount] > 7, Orders[Amount] < 20))",
        )
        .unwrap();
    model
        .add_measure(
            "Selected (OR fn)",
            "CALCULATE(SUM(Orders[Amount]), OR(Orders[Amount] = 10, Orders[Amount] = 20))",
        )
        .unwrap();

    assert_eq!(
        model
            .evaluate_measure("Medium (AND fn)", &FilterContext::empty())
            .unwrap(),
        18.0.into()
    );
    assert_eq!(
        model
            .evaluate_measure("Selected (OR fn)", &FilterContext::empty())
            .unwrap(),
        30.0.into()
    );
}

#[test]
fn calculate_boolean_filter_expressions_must_reference_one_table() {
    let model = build_model();
    let err = DaxEngine::new()
        .evaluate(
            &model,
            "CALCULATE(SUM(Orders[Amount]), Orders[Amount] > 7 && Customers[Region] = \"East\")",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap_err();

    match err {
        DaxError::Eval(msg) => {
            assert!(msg.contains("exactly one table"));
            assert!(msg.contains("Orders"));
            assert!(msg.contains("Customers"));
        }
        other => panic!("expected DaxError::Eval, got {other:?}"),
    }
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
        .map(|row| customers.value(row, "IsEast").unwrap())
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
fn switch_supports_simple_and_true_idiom_forms() {
    let model = build_model();
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                "SWITCH(1, 0, \"a\", 1, \"b\", \"c\")",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        Value::from("b")
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "SWITCH(1, 0, \"a\", 2, \"b\", \"c\")",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        Value::from("c")
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "SWITCH(TRUE(), 1=2, \"no\", 2=2, \"yes\")",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        Value::from("yes")
    );
}

#[test]
fn concat_operator_ampersand_concatenates_strings_and_coerces_blank() {
    let model = DataModel::new();
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                r#""A" & "B""#,
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        Value::from("AB")
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                r#""A" & BLANK()"#,
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        Value::from("A")
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                r#"BLANK() & "B""#,
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        Value::from("B")
    );
}

#[test]
fn isblank_returns_true_only_for_blank() {
    let model = DataModel::new();
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                "ISBLANK(BLANK())",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        Value::from(true)
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "ISBLANK(0)",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        Value::from(false)
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "ISBLANK(\"\")",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        Value::from(false)
    );
}

#[test]
fn isblank_enforces_arity() {
    let model = DataModel::new();
    let engine = DaxEngine::new();

    let err = engine
        .evaluate(
            &model,
            "ISBLANK()",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap_err();
    assert!(err.to_string().contains("ISBLANK expects 1 argument"));

    let err = engine
        .evaluate(
            &model,
            "ISBLANK(1, 2)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap_err();
    assert!(err.to_string().contains("ISBLANK expects 1 argument"));
}

#[test]
fn isblank_is_true_for_blank_measures() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let empty_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "Nowhere".into());
    let value = DaxEngine::new()
        .evaluate(
            &model,
            "ISBLANK([Total Sales])",
            &empty_filter,
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(true));
}

#[test]
fn if_isblank_pattern_can_replace_blank_with_zero() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let engine = DaxEngine::new();

    // When the measure is non-blank, IF should return its value.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "IF(ISBLANK([Total Sales]), 0, [Total Sales])",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        43.0.into()
    );

    // When filters yield no rows, [Total Sales] is BLANK and IF should return 0.
    let empty_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "Nowhere".into());
    assert_eq!(
        engine
            .evaluate(
                &model,
                "IF(ISBLANK([Total Sales]), 0, [Total Sales])",
                &empty_filter,
                &RowContext::default(),
            )
            .unwrap(),
        0.into()
    );
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

#[test]
fn values_and_summarize_support_basic_grouping() {
    let model = build_model();
    let engine = DaxEngine::new();

    let regions = engine
        .evaluate(
            &model,
            "COUNTROWS(VALUES(Customers[Region]))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(regions, 2.into());

    let customers = engine
        .evaluate(
            &model,
            "COUNTROWS(SUMMARIZE(Orders, Orders[CustomerId]))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(customers, 3.into());

    let regions = engine
        .evaluate(
            &model,
            "COUNTROWS(SUMMARIZE(Orders, Customers[Region]))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(regions, 2.into());
}

#[test]
fn summarizecolumns_supports_basic_grouping() {
    let model = build_model();
    let engine = DaxEngine::new();

    let summarizecolumns = engine
        .evaluate(
            &model,
            "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region]))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    let distinct = engine
        .evaluate(
            &model,
            "DISTINCTCOUNT(Customers[Region])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(summarizecolumns, distinct);
}

#[test]
fn summarizecolumns_supports_filter_table_arguments() {
    let model = build_model();
    let engine = DaxEngine::new();

    let value = engine
        .evaluate(
            &model,
            "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region], FILTER(Customers, Customers[Region] = \"East\")))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 1.into());
}

#[test]
fn summarizecolumns_filter_args_do_not_perform_context_transition() {
    // Row context should not automatically become filter context for SUMMARIZECOLUMNS.
    // (Context transition is only performed by CALCULATE/CALCULATETABLE or measure evaluation.)
    let mut model = build_model();
    model
        .add_calculated_column(
            "Orders",
            "Customer Groups via SUMMARIZECOLUMNS",
            "COUNTROWS(SUMMARIZECOLUMNS(Orders[CustomerId], Orders[Amount] <> BLANK()))",
        )
        .unwrap();

    let orders = model.table("Orders").unwrap();
    let values: Vec<Value> = (0..orders.row_count())
        .map(|row| orders.value(row, "Customer Groups via SUMMARIZECOLUMNS").unwrap())
        .collect();

    assert_eq!(values, vec![3.into(), 3.into(), 3.into(), 3.into()]);
}

#[test]
fn summarizecolumns_allows_name_expression_pairs_but_does_not_materialize_them_yet() {
    let model = build_model();
    let engine = DaxEngine::new();

    // Client tools commonly emit SUMMARIZECOLUMNS with "Name", expr pairs for measures.
    // The current engine only uses SUMMARIZECOLUMNS for group construction (row set), so the
    // named columns are accepted but not returned as part of the table representation.
    let value = engine
        .evaluate(
            &model,
            "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region], FILTER(Customers, Customers[Region] = \"East\"), \"X\", 1))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 1.into());
}

fn build_summarizecolumns_star_schema_model() -> DataModel {
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
}

#[test]
fn summarizecolumns_star_schema_groups_by_multiple_dimensions() {
    let model = build_summarizecolumns_star_schema_model();

    let engine = DaxEngine::new();
    let groups = engine
        .evaluate(
            &model,
            "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region], Products[Category]))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    // Distinct combinations in Sales are:
    //   (East, A), (East, B), (West, A), (West, B)
    assert_eq!(groups, 4.into());
}

#[test]
fn summarizecolumns_star_schema_respects_filter_table_arguments() {
    let model = build_summarizecolumns_star_schema_model();
    let engine = DaxEngine::new();

    let groups = engine
        .evaluate(
            &model,
            "COUNTROWS(SUMMARIZECOLUMNS(Customers[Region], Products[Category], FILTER(Customers, Customers[Region] = \"East\")))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(groups, 2.into());
}

#[test]
fn calculate_all_can_remove_column_filters() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    model
        .add_measure(
            "All Region Sales",
            "CALCULATE([Total Sales], ALL(Customers[Region]))",
        )
        .unwrap();

    let east_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    let value = model
        .evaluate_measure("All Region Sales", &east_filter)
        .unwrap();
    assert_eq!(value, 43.0.into());
}

#[test]
fn calculate_removefilters_can_remove_column_filters() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    model
        .add_measure(
            "All Region Sales 2",
            "CALCULATE([Total Sales], REMOVEFILTERS(Customers[Region]))",
        )
        .unwrap();

    let east_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    let value = model
        .evaluate_measure("All Region Sales 2", &east_filter)
        .unwrap();
    assert_eq!(value, 43.0.into());
}

#[test]
fn calculate_all_removes_row_context_filters_for_measures() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let engine = DaxEngine::new();
    let orders = model.table("Orders").unwrap();
    for row in 0..orders.row_count() {
        let mut row_ctx = RowContext::default();
        row_ctx.push("Orders", row);
        let value = engine
            .evaluate(
                &model,
                "CALCULATE([Total Sales], ALL(Orders))",
                &FilterContext::empty(),
                &row_ctx,
            )
            .unwrap();
        assert_eq!(value, 43.0.into());
    }
}

#[test]
fn pivot_api_groups_and_evaluates_measures() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let measures = vec![PivotMeasure::new("Total Sales", "[Total Sales]").unwrap()];
    let group_by = vec![GroupByColumn::new("Customers", "Region")];

    let result = pivot(
        &model,
        "Orders",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();
    assert_eq!(
        result.columns,
        vec!["Customers[Region]".to_string(), "Total Sales".to_string()]
    );
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("East"), Value::from(38.0)],
            vec![Value::from("West"), Value::from(5.0)]
        ]
    );
}

#[test]
fn large_synthetic_pivot_is_linearish() {
    let mut model = DataModel::new();
    let mut fact = Table::new("Fact", vec!["Group", "Amount"]);

    let groups = ["A", "B", "C", "D"];
    let mut expected_total = 0.0f64;
    let mut expected_by_group = [0.0f64; 4];

    for i in 0..20_000 {
        let g = groups[i % groups.len()];
        let amount = (i % 100) as f64;
        expected_total += amount;
        expected_by_group[i % groups.len()] += amount;
        fact.push_row(vec![g.into(), amount.into()]).unwrap();
    }
    model.add_table(fact).unwrap();
    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();

    assert_eq!(
        model
            .evaluate_measure("Total", &FilterContext::empty())
            .unwrap(),
        expected_total.into()
    );

    let measures = vec![PivotMeasure::new("Total", "SUM(Fact[Amount])").unwrap()];
    let group_by = vec![GroupByColumn::new("Fact", "Group")];
    let result = pivot(
        &model,
        "Fact",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();
    assert_eq!(result.rows.len(), groups.len());
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("A"), expected_by_group[0].into()],
            vec![Value::from("B"), expected_by_group[1].into()],
            vec![Value::from("C"), expected_by_group[2].into()],
            vec![Value::from("D"), expected_by_group[3].into()],
        ]
    );
}

#[test]
fn columnar_tables_support_measures_and_filter_propagation() {
    let mut model = DataModel::new();

    let customers_schema = vec![
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Name".to_string(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "Region".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut customers = ColumnarTableBuilder::new(customers_schema, options);
    customers.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("Alice")),
        formula_columnar::Value::String(Arc::<str>::from("East")),
    ]);
    customers.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("Bob")),
        formula_columnar::Value::String(Arc::<str>::from("West")),
    ]);
    customers.append_row(&[
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::String(Arc::<str>::from("Carol")),
        formula_columnar::Value::String(Arc::<str>::from("East")),
    ]);
    model
        .add_table(Table::from_columnar("Customers", customers.finalize()))
        .unwrap();

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
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut orders = ColumnarTableBuilder::new(orders_schema, options);
    orders.append_row(&[
        formula_columnar::Value::Number(100.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    orders.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(20.0),
    ]);
    orders.append_row(&[
        formula_columnar::Value::Number(102.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(5.0),
    ]);
    orders.append_row(&[
        formula_columnar::Value::Number(103.0),
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::Number(8.0),
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
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    assert_eq!(
        model
            .evaluate_measure("Total Sales", &FilterContext::empty())
            .unwrap(),
        43.0.into()
    );

    let east_filter =
        FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    assert_eq!(
        model.evaluate_measure("Total Sales", &east_filter).unwrap(),
        38.0.into()
    );
}

#[test]
fn persisted_columnar_table_can_register_calculated_column_definition_without_recomputing() {
    let mut model = DataModel::new();

    // Simulate a persisted model: calculated column values are already present in the stored
    // columnar table, and we only need to register the DAX expression metadata.
    let orders_schema = vec![
        ColumnSchema {
            name: "OrderId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Double Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut orders = ColumnarTableBuilder::new(orders_schema, options);

    let amounts = [10.0, 20.0, 5.0, 8.0];
    for (idx, amount) in amounts.iter().copied().enumerate() {
        orders.append_row(&[
            formula_columnar::Value::Number((100 + idx) as f64),
            formula_columnar::Value::Number(amount),
            formula_columnar::Value::Number(amount * 2.0),
        ]);
    }
    model
        .add_table(Table::from_columnar("Orders", orders.finalize()))
        .unwrap();

    // Negative: registering a definition for a column that doesn't exist should fail.
    let err = model
        .add_calculated_column_definition("Orders", "Missing Column", "Orders[Amount] * 2")
        .unwrap_err();
    assert!(matches!(
        err,
        DaxError::UnknownColumn { table, column }
            if table == "Orders" && column == "Missing Column"
    ));

    model
        .add_calculated_column_definition("Orders", "Double Amount", "Orders[Amount] * 2")
        .unwrap();

    assert!(model.calculated_columns().iter().any(|c| {
        c.table == "Orders" && c.name == "Double Amount" && c.expression == "Orders[Amount] * 2"
    }));

    // Ensure values remain readable from the table and match the persisted values.
    let orders = model.table("Orders").unwrap();
    let values: Vec<Value> = (0..orders.row_count())
        .map(|row| orders.value(row, "Double Amount").unwrap())
        .collect();
    assert_eq!(values, vec![20.0.into(), 40.0.into(), 10.0.into(), 16.0.into()]);

    // Negative: registering the same calculated column definition twice should fail.
    let err = model
        .add_calculated_column_definition("Orders", "Double Amount", "Orders[Amount] * 2")
        .unwrap_err();
    assert!(matches!(
        err,
        DaxError::DuplicateColumn { table, column } if table == "Orders" && column == "Double Amount"
    ));
}

#[test]
fn columnar_tables_support_calculated_columns() {
    let mut model = DataModel::new();

    let schema = vec![ColumnSchema {
        name: "Amount".to_string(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(schema, options);
    fact.append_row(&[formula_columnar::Value::Number(10.0)]);
    fact.append_row(&[formula_columnar::Value::Number(5.0)]);
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

    model
        .add_calculated_column("Fact", "DoubleAmount", "[Amount] * 2")
        .unwrap();

    let fact = model.table("Fact").unwrap();
    assert_eq!(
        fact.columns(),
        &["Amount".to_string(), "DoubleAmount".to_string()]
    );
    assert_eq!(fact.value(0, "DoubleAmount"), Some(20.0.into()));
    assert_eq!(fact.value(1, "DoubleAmount"), Some(10.0.into()));

    let columnar = fact.columnar_table().unwrap();
    assert_eq!(columnar.column_count(), 2);
    assert_eq!(columnar.row_count(), 2);
}

#[test]
fn columnar_tables_support_calculated_columns_materialized_into_storage() {
    let mut model = DataModel::new();

    let customers_schema = vec![
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Name".to_string(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "Region".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut customers = ColumnarTableBuilder::new(customers_schema, options);
    customers.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("Alice")),
        formula_columnar::Value::String(Arc::<str>::from("East")),
    ]);
    customers.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("Bob")),
        formula_columnar::Value::String(Arc::<str>::from("West")),
    ]);
    customers.append_row(&[
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::String(Arc::<str>::from("Carol")),
        formula_columnar::Value::String(Arc::<str>::from("East")),
    ]);
    model
        .add_table(Table::from_columnar("Customers", customers.finalize()))
        .unwrap();

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
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut orders = ColumnarTableBuilder::new(orders_schema, options);
    orders.append_row(&[
        formula_columnar::Value::Number(100.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    orders.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(20.0),
    ]);
    orders.append_row(&[
        formula_columnar::Value::Number(102.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(5.0),
    ]);
    orders.append_row(&[
        formula_columnar::Value::Number(103.0),
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::Number(8.0),
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
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_calculated_column("Orders", "CustomerName", "RELATED(Customers[Name])")
        .unwrap();
    model
        .add_calculated_column("Orders", "Double Amount", "[Amount] * 2")
        .unwrap();

    let orders = model.table("Orders").unwrap();
    assert_eq!(
        orders.columns(),
        &[
            "OrderId".to_string(),
            "CustomerId".to_string(),
            "Amount".to_string(),
            "CustomerName".to_string(),
            "Double Amount".to_string()
        ]
    );

    let columnar = orders.columnar_table().unwrap();
    assert_eq!(columnar.column_count(), 5);
    assert_eq!(columnar.row_count(), 4);

    let names: Vec<Value> = (0..orders.row_count())
        .map(|row| orders.value(row, "CustomerName").unwrap())
        .collect();
    assert_eq!(
        names,
        vec![
            Value::from("Alice"),
            Value::from("Alice"),
            Value::from("Bob"),
            Value::from("Carol")
        ]
    );

    let doubled: Vec<Value> = (0..orders.row_count())
        .map(|row| orders.value(row, "Double Amount").unwrap())
        .collect();
    assert_eq!(
        doubled,
        vec![20.0.into(), 40.0.into(), 10.0.into(), 16.0.into()]
    );
}

#[test]
fn columnar_tables_reject_mixed_type_calculated_columns() {
    let mut model = DataModel::new();

    let schema = vec![ColumnSchema {
        name: "OrderId".to_string(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(schema, options);
    fact.append_row(&[formula_columnar::Value::Number(100.0)]);
    fact.append_row(&[formula_columnar::Value::Number(101.0)]);
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

    let err = model
        .add_calculated_column("Fact", "Mixed", "IF([OrderId] = 100, 1, \"x\")")
        .unwrap_err();
    assert!(matches!(err, DaxError::Type(_)));
}

#[test]
fn measure_in_row_context_performs_implicit_context_transition() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    model
        .add_calculated_column("Customers", "Sales via measure", "[Total Sales]")
        .unwrap();

    let customers = model.table("Customers").unwrap();
    let values: Vec<Value> = (0..customers.row_count())
        .map(|row| customers.value(row, "Sales via measure").unwrap())
        .collect();
    assert_eq!(values, vec![30.0.into(), 5.0.into(), 8.0.into()]);
}

#[test]
fn measure_in_iterator_performs_implicit_context_transition() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let value = DaxEngine::new()
        .evaluate(
            &model,
            "SUMX(Orders, [Total Sales])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 43.0.into());
}

#[test]
fn sumx_values_column_measure_context_transition_filters_only_that_column() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let value = DaxEngine::new()
        .evaluate(
            &model,
            "SUMX(VALUES(Orders[CustomerId]), [Total Sales])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, 43.0.into());
}

#[test]
fn sumx_values_column_iterates_distinct_values() {
    let model = build_model();
    let value = DaxEngine::new()
        .evaluate(
            &model,
            "SUMX(VALUES(Customers[Region]), 1)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, 2.into());
}

#[test]
fn values_column_row_context_disallows_other_columns() {
    let model = build_model();
    let err = DaxEngine::new()
        .evaluate(
            &model,
            "SUMX(VALUES(Orders[CustomerId]), Orders[Amount])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap_err();

    assert!(matches!(err, DaxError::Eval(_)));
    assert!(err
        .to_string()
        .contains("not available in the current row context"));
}

#[test]
fn filter_values_restricts_row_context_columns() {
    let model = build_model();
    let value = DaxEngine::new()
        .evaluate(
            &model,
            "COUNTROWS(FILTER(VALUES(Orders[CustomerId]), Orders[CustomerId] = 1))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 1.into());

    let err = DaxEngine::new()
        .evaluate(
            &model,
            "COUNTROWS(FILTER(VALUES(Orders[CustomerId]), Orders[Amount] > 0))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap_err();
    assert!(matches!(err, DaxError::Eval(_)));
    assert!(err
        .to_string()
        .contains("not available in the current row context"));
}

#[test]
fn allexcept_keeps_only_listed_columns() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    model
        .add_measure(
            "Sales by Region",
            "CALCULATE([Total Sales], ALLEXCEPT(Customers, Customers[Region]))",
        )
        .unwrap();

    let mut filter = FilterContext::empty();
    filter.set_column_equals("Customers", "Region", "East".into());
    filter.set_column_equals("Customers", "Name", "Bob".into());

    assert_eq!(
        model.evaluate_measure("Sales by Region", &filter).unwrap(),
        38.0.into()
    );
}

#[test]
fn calculate_all_and_values_can_reapply_original_column_filters() {
    let mut model = build_model();
    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();
    model
        .add_measure(
            "Sales keeping region",
            "CALCULATE([Total Sales], ALL(Customers), VALUES(Customers[Region]))",
        )
        .unwrap();
    model
        .add_measure(
            "Sales keeping region (reversed)",
            "CALCULATE([Total Sales], VALUES(Customers[Region]), ALL(Customers))",
        )
        .unwrap();

    let mut filter = FilterContext::empty();
    filter.set_column_equals("Customers", "Region", "East".into());
    filter.set_column_equals("Customers", "Name", "Alice".into());

    assert_eq!(
        model
            .evaluate_measure("Sales keeping region", &filter)
            .unwrap(),
        38.0.into()
    );
    assert_eq!(
        model
            .evaluate_measure("Sales keeping region (reversed)", &filter)
            .unwrap(),
        38.0.into()
    );
}

#[test]
fn calculatetable_can_be_used_as_table_filter_argument() {
    let model = build_model();
    let value = DaxEngine::new()
        .evaluate(
            &model,
            "COUNTROWS(CALCULATETABLE(Orders, Customers[Region] = \"East\"))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 3.into());
}

#[test]
fn userelationship_activates_inactive_relationship_and_overrides_active() {
    let mut model = DataModel::new();

    let mut dates = Table::new("Date", vec!["DateKey"]);
    dates.push_row(vec![1.into()]).unwrap();
    dates.push_row(vec![2.into()]).unwrap();
    model.add_table(dates).unwrap();

    let mut sales = Table::new("Sales", vec!["OrderDateKey", "ShipDateKey", "Amount"]);
    sales
        .push_row(vec![1.into(), 2.into(), 10.0.into()])
        .unwrap();
    sales
        .push_row(vec![2.into(), 1.into(), 5.0.into()])
        .unwrap();
    sales
        .push_row(vec![2.into(), 2.into(), 7.0.into()])
        .unwrap();
    model.add_table(sales).unwrap();

    model
        .add_relationship(Relationship {
            name: "Sales_OrderDate".into(),
            from_table: "Sales".into(),
            from_column: "OrderDateKey".into(),
            to_table: "Date".into(),
            to_column: "DateKey".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();
    model
        .add_relationship(Relationship {
            name: "Sales_ShipDate".into(),
            from_table: "Sales".into(),
            from_column: "ShipDateKey".into(),
            to_table: "Date".into(),
            to_column: "DateKey".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: false,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Sales", "SUM(Sales[Amount])").unwrap();
    model
        .add_measure(
            "Sales by ShipDate",
            "CALCULATE([Sales], USERELATIONSHIP(Sales[ShipDateKey], Date[DateKey]))",
        )
        .unwrap();

    let date2_filter = FilterContext::empty().with_column_equals("Date", "DateKey", 2.into());
    assert_eq!(
        model.evaluate_measure("Sales", &date2_filter).unwrap(),
        12.0.into()
    );
    assert_eq!(
        model
            .evaluate_measure("Sales by ShipDate", &date2_filter)
            .unwrap(),
        17.0.into()
    );
}

#[test]
fn related_respects_userelationship_overrides() {
    let mut model = DataModel::new();
    let mut dates = Table::new("Date", vec!["DateKey"]);
    dates.push_row(vec![1.into()]).unwrap();
    dates.push_row(vec![2.into()]).unwrap();
    model.add_table(dates).unwrap();

    let mut sales = Table::new("Sales", vec!["OrderDateKey", "ShipDateKey", "Amount"]);
    sales
        .push_row(vec![1.into(), 2.into(), 10.0.into()])
        .unwrap();
    sales
        .push_row(vec![2.into(), 1.into(), 5.0.into()])
        .unwrap();
    sales
        .push_row(vec![2.into(), 2.into(), 7.0.into()])
        .unwrap();
    model.add_table(sales).unwrap();

    model
        .add_relationship(Relationship {
            name: "Sales_OrderDate".into(),
            from_table: "Sales".into(),
            from_column: "OrderDateKey".into(),
            to_table: "Date".into(),
            to_column: "DateKey".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();
    model
        .add_relationship(Relationship {
            name: "Sales_ShipDate".into(),
            from_table: "Sales".into(),
            from_column: "ShipDateKey".into(),
            to_table: "Date".into(),
            to_column: "DateKey".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: false,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_calculated_column("Sales", "OrderDateViaRelated", "RELATED(Date[DateKey])")
        .unwrap();
    model
        .add_calculated_column(
            "Sales",
            "ShipDateViaRelated",
            "CALCULATE(RELATED(Date[DateKey]), USERELATIONSHIP(Sales[ShipDateKey], Date[DateKey]))",
        )
        .unwrap();

    let sales = model.table("Sales").unwrap();
    let values: Vec<(Value, Value, Value, Value)> = (0..sales.row_count())
        .map(|row| {
            (
                sales.value(row, "OrderDateKey").unwrap(),
                sales.value(row, "ShipDateKey").unwrap(),
                sales.value(row, "OrderDateViaRelated").unwrap(),
                sales.value(row, "ShipDateViaRelated").unwrap(),
            )
        })
        .collect();

    for (order_key, ship_key, order_related, ship_related) in values {
        assert_eq!(order_related, order_key);
        assert_eq!(ship_related, ship_key);
    }
}

#[test]
fn bidirectional_relationship_propagates_filters_to_dimension() {
    let model = build_model_bidirectional();
    let filter = FilterContext::empty().with_column_equals("Orders", "Amount", 20.0.into());

    let value = DaxEngine::new()
        .evaluate(
            &model,
            "COUNTROWS(Customers)",
            &filter,
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, 1.into());
}

#[test]
fn crossfilter_can_override_relationship_direction_inside_calculate() {
    let model = build_model();
    let filter = FilterContext::empty().with_column_equals("Orders", "Amount", 20.0.into());

    let engine = DaxEngine::new();
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(Customers)",
                &filter,
                &RowContext::default(),
            )
            .unwrap(),
        3.into()
    );

    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(Customers), CROSSFILTER(Orders[CustomerId], Customers[CustomerId], BOTH))",
                &filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn crossfilter_none_can_disable_relationship_inside_calculate() {
    let model = build_model();
    let filter = FilterContext::empty().with_column_equals("Customers", "Region", "East".into());
    let engine = DaxEngine::new();

    assert_eq!(
        engine
            .evaluate(
                &model,
                "SUM(Orders[Amount])",
                &filter,
                &RowContext::default(),
            )
            .unwrap(),
        38.0.into()
    );

    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(SUM(Orders[Amount]), CROSSFILTER(Orders[CustomerId], Customers[CustomerId], NONE))",
                &filter,
                &RowContext::default(),
            )
            .unwrap(),
        43.0.into()
    );
}

#[test]
fn count_and_counta_respect_types_and_filters() {
    let mut model = DataModel::new();
    let mut t = Table::new("T", vec!["Col"]);
    t.push_row(vec![1.0.into()]).unwrap();
    t.push_row(vec![Value::Blank]).unwrap();
    t.push_row(vec!["x".into()]).unwrap();
    t.push_row(vec![true.into()]).unwrap();
    model.add_table(t).unwrap();

    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            "COUNT(T[Col])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 1.into());

    let value = engine
        .evaluate(
            &model,
            "COUNTA(T[Col])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 3.into());

    let filter = FilterContext::empty().with_column_equals("T", "Col", "x".into());
    let value = engine
        .evaluate(&model, "COUNT(T[Col])", &filter, &RowContext::default())
        .unwrap();
    assert_eq!(value, 0.into());

    let value = engine
        .evaluate(&model, "COUNTA(T[Col])", &filter, &RowContext::default())
        .unwrap();
    assert_eq!(value, 1.into());
}

#[test]
fn relationship_does_not_filter_facts_when_dimension_is_unfiltered() {
    // When referential integrity is not enforced, tabular models include fact rows whose
    // foreign key has no match in the dimension. Those rows should only be removed when the
    // dimension is filtered (otherwise they contribute to totals and show up under a blank
    // dimension member when grouped).
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Id"]);
    dim.push_row(vec![1.into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Group"]);
    fact.push_row(vec![1.into(), "A".into()]).unwrap();
    fact.push_row(vec![999.into(), "A".into()]).unwrap();
    model.add_table(fact).unwrap();

    model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Id".into(),
            to_table: "Dim".into(),
            to_column: "Id".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model.add_measure("Fact Rows", "COUNTROWS(Fact)").unwrap();

    let filter = FilterContext::empty().with_column_equals("Fact", "Group", "A".into());
    assert_eq!(
        model.evaluate_measure("Fact Rows", &filter).unwrap(),
        2.into()
    );
}

#[test]
fn relationship_blank_dimension_member_includes_unmatched_facts() {
    let mut model = DataModel::new();

    let mut customers = Table::new("Customers", vec!["CustomerId", "Region"]);
    customers.push_row(vec![1.into(), "East".into()]).unwrap();
    customers.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(customers).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "CustomerId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 999.into(), 7.0.into()])
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
        .add_measure("Total Sales", "SUM(Orders[Amount])")
        .unwrap();

    let blank_region =
        FilterContext::empty().with_column_equals("Customers", "Region", Value::Blank);
    assert_eq!(
        model
            .evaluate_measure("Total Sales", &blank_region)
            .unwrap(),
        7.0.into()
    );
}

#[test]
fn values_and_distinctcount_include_virtual_blank_row_for_unmatched_relationship_keys() {
    let mut model = DataModel::new();

    let mut customers = Table::new("Customers", vec!["CustomerId", "Region"]);
    customers.push_row(vec![1.into(), "East".into()]).unwrap();
    customers.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(customers).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "CustomerId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 999.into(), 7.0.into()])
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
            enforce_referential_integrity: false,
        })
        .unwrap();

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
    assert_eq!(
        engine
            .evaluate(
                &model,
                "DISTINCTCOUNTNOBLANK(Customers[Region])",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        2.into()
    );

    let blank_region =
        FilterContext::empty().with_column_equals("Customers", "Region", Value::Blank);
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(VALUES(Customers[Region]))",
                &blank_region,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "DISTINCTCOUNT(Customers[Region])",
                &blank_region,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "DISTINCTCOUNTNOBLANK(Customers[Region])",
                &blank_region,
                &RowContext::default(),
            )
            .unwrap(),
        0.into()
    );
}

#[test]
fn allnoblankrow_excludes_relationship_blank_member() {
    let mut model = DataModel::new();

    let mut customers = Table::new("Customers", vec!["CustomerId", "Region"]);
    customers.push_row(vec![1.into(), "East".into()]).unwrap();
    customers.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(customers).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "CustomerId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 999.into(), 7.0.into()])
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    let engine = DaxEngine::new();
    let empty = FilterContext::empty();

    // `ALL` includes the relationship-generated blank member (unknown customer) when it exists.
    assert_eq!(
        engine
            .evaluate(
                &model,
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
                &model,
                "COUNTROWS(ALLNOBLANKROW(Customers[Region]))",
                &empty,
                &RowContext::default(),
            )
            .unwrap(),
        2.into()
    );

    assert_eq!(
        engine
            .evaluate(
                &model,
                "SUM(Orders[Amount])",
                &empty,
                &RowContext::default(),
            )
            .unwrap(),
        17.0.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(SUM(Orders[Amount]), ALLNOBLANKROW(Customers[Region]))",
                &empty,
                &RowContext::default(),
            )
            .unwrap(),
        10.0.into()
    );
}

#[test]
fn relationship_filter_including_all_real_dimension_rows_can_exclude_blank_member() {
    let mut model = DataModel::new();

    let mut customers = Table::new("Customers", vec!["CustomerId", "Region"]);
    customers.push_row(vec![1.into(), "East".into()]).unwrap();
    customers.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(customers).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "CustomerId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 999.into(), 7.0.into()])
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    let value = DaxEngine::new()
        .evaluate(
            &model,
            "CALCULATE(COUNTROWS(Orders), Customers[Region] <> BLANK())",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, 1.into());
}
