use formula_dax::{
    pivot, Cardinality, CrossFilterDirection, DataModel, DaxEngine, DaxError, FilterContext,
    GroupByColumn, PivotMeasure, Relationship, RowContext, Table, Value,
};

fn build_model() -> DataModel {
    build_model_with_relationship_active(true)
}

fn build_model_with_relationship_active(is_active: bool) -> DataModel {
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
            is_active,
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

#[test]
fn identifiers_are_case_insensitive_for_unicode_names() {
    // Use a German sharp S (ß) to ensure we handle Unicode-aware case folding for identifiers.
    //
    // In particular, `ß` uppercases to `SS`, so `'Straße'` should be addressable as `'STRASSE'`.
    let mut model = DataModel::new();
    let mut table = Table::new("Straße", vec!["Maß"]);
    table.push_row(vec![1.0.into()]).unwrap();
    table.push_row(vec![2.0.into()]).unwrap();
    model.add_table(table).unwrap();

    model.add_measure("Total", "SUM('STRASSE'[MASS])").unwrap();
    let total = model
        .evaluate_measure("[TOTAL]", &FilterContext::empty())
        .unwrap();
    assert_eq!(total, Value::from(3.0));

    // Unquoted Unicode identifiers should also be accepted when they match identifier-like rules.
    // (This aligns with the pivot/model display helpers that render `Straße[Maß]` without quotes.)
    model.add_measure("Total Unquoted", "SUM(Straße[Maß])")
        .unwrap();
    let total_unquoted = model
        .evaluate_measure("total unquoted", &FilterContext::empty())
        .unwrap();
    assert_eq!(total_unquoted, Value::from(3.0));
}

#[test]
fn var_names_are_case_insensitive_for_unicode_names() {
    let model = DataModel::new();
    let engine = DaxEngine::new();

    // Variables use the same case-insensitive identifier matching rules as tables/columns.
    let value = engine
        .evaluate(
            &model,
            "VAR Straße = 1 RETURN STRASSE",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(1.0));
}

#[test]
fn pivot_resolves_unicode_identifiers_case_insensitively_and_preserves_model_casing() {
    let mut model = DataModel::new();
    let mut table = Table::new("Straße", vec!["Kategorie", "Maß"]);
    table.push_row(vec!["A".into(), 1.0.into()]).unwrap();
    table.push_row(vec!["A".into(), 2.0.into()]).unwrap();
    table.push_row(vec!["B".into(), 4.0.into()]).unwrap();
    model.add_table(table).unwrap();

    model.add_measure("Total", "SUM('Straße'[Maß])").unwrap();

    let group_by = vec![GroupByColumn::new("STRASSE", "kategorie")];
    let measures = vec![PivotMeasure::new("Total", "[TOTAL]").unwrap()];

    let result = pivot(
        &model,
        // Use ASCII table identifier; the model stores `Straße`.
        "strasse",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(result.columns, vec!["Straße[Kategorie]", "Total"]);
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("A"), Value::from(3.0)],
            vec![Value::from("B"), Value::from(4.0)],
        ]
    );
}

#[test]
fn unicode_identifiers_work_in_filter_context() {
    let mut model = DataModel::new();
    let mut table = Table::new("Straße", vec!["Maß"]);
    table.push_row(vec![1.0.into()]).unwrap();
    table.push_row(vec![2.0.into()]).unwrap();
    table.push_row(vec![3.0.into()]).unwrap();
    model.add_table(table).unwrap();

    model.add_measure("Total", "SUM('Straße'[Maß])").unwrap();

    let total = model
        .evaluate_measure("[TOTAL]", &FilterContext::empty())
        .unwrap();
    assert_eq!(total, Value::from(6.0));

    // Filter using ASCII identifiers that casefold to the Unicode table/column names.
    let filter = FilterContext::empty().with_column_equals("strasse", "mass", 2.0.into());
    let filtered = model.evaluate_measure("[TOTAL]", &filter).unwrap();
    assert_eq!(filtered, Value::from(2.0));

    let filter =
        FilterContext::empty().with_column_in("STRASSE", "MASS", [1.0.into(), 3.0.into()]);
    let filtered = model.evaluate_measure("[TOTAL]", &filter).unwrap();
    assert_eq!(filtered, Value::from(4.0));
}

#[test]
fn unicode_relationships_are_case_insensitive_for_filters_and_pivots() {
    let mut model = DataModel::new();

    let mut streets = Table::new("Straße", vec!["StraßenId", "Region"]);
    streets.push_row(vec![1.into(), "East".into()]).unwrap();
    streets.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(streets).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "StraßenId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 1.into(), 20.0.into()])
        .unwrap();
    orders
        .push_row(vec![102.into(), 2.into(), 5.0.into()])
        .unwrap();
    model.add_table(orders).unwrap();

    // Add relationship using mixed casing and ASCII-only identifiers that casefold to the Unicode
    // model names (`Straße` -> `STRASSE`, `StraßenId` -> `STRASSENID`).
    model
        .add_relationship(Relationship {
            name: "Orders->Straße".into(),
            from_table: "orders".into(),
            from_column: "straßenid".into(),
            to_table: "STRASSE".into(),
            to_column: "STRASSENID".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Total", "SUM(orders[amount])").unwrap();

    let east_filter =
        FilterContext::empty().with_column_equals("strasse", "region", "East".into());
    let east_total = model.evaluate_measure("[TOTAL]", &east_filter).unwrap();
    assert_eq!(east_total, Value::from(30.0));

    // Pivot should also traverse relationships with Unicode identifiers and preserve the model's
    // original casing in the output headers.
    let group_by = vec![GroupByColumn::new("STRASSE", "region")];
    let measures = vec![PivotMeasure::new("Total", "[TOTAL]").unwrap()];
    let result = pivot(
        &model,
        "orders",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();
    assert_eq!(result.columns, vec!["Straße[Region]", "Total"]);
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("East"), Value::from(30.0)],
            vec![Value::from("West"), Value::from(5.0)],
        ]
    );
}

#[test]
fn userelationship_resolves_unicode_relationship_case_insensitively() {
    let mut model = DataModel::new();

    let mut streets = Table::new("Straße", vec!["StraßenId", "Region"]);
    streets.push_row(vec![1.into(), "East".into()]).unwrap();
    streets.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(streets).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "StraßenId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 1.into(), 20.0.into()])
        .unwrap();
    orders
        .push_row(vec![102.into(), 2.into(), 5.0.into()])
        .unwrap();
    model.add_table(orders).unwrap();

    // Start with an inactive relationship so USERELATIONSHIP is required for propagation.
    model
        .add_relationship(Relationship {
            name: "Orders->Straße (inactive)".into(),
            from_table: "orders".into(),
            from_column: "straßenid".into(),
            to_table: "STRASSE".into(),
            to_column: "STRASSENID".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: false,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Orders Total", "SUM(orders[amount])").unwrap();
    model
        .add_measure(
            "Orders East (inactive)",
            "CALCULATE([Orders Total], STRASSE[REGION] = \"East\")",
        )
        .unwrap();
    model
        .add_measure(
            "Orders East (userelationship)",
            "CALCULATE([Orders Total], USERELATIONSHIP(orders[straßenid], STRASSE[STRASSENID]), STRASSE[REGION] = \"East\")",
        )
        .unwrap();

    let inactive = model
        .evaluate_measure("orders east (inactive)", &FilterContext::empty())
        .unwrap();
    assert_eq!(
        inactive,
        Value::from(35.0),
        "inactive relationship should not propagate the STRASSE[Region] filter"
    );

    let with_userelationship = model
        .evaluate_measure("[ORDERS EAST (USERELATIONSHIP)]", &FilterContext::empty())
        .unwrap();
    assert_eq!(with_userelationship, Value::from(30.0));
}

#[test]
fn crossfilter_resolves_unicode_relationship_case_insensitively() {
    let mut model = DataModel::new();

    let mut streets = Table::new("Straße", vec!["StraßenId", "Region"]);
    streets.push_row(vec![1.into(), "East".into()]).unwrap();
    streets.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(streets).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "StraßenId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 1.into(), 20.0.into()])
        .unwrap();
    orders
        .push_row(vec![102.into(), 2.into(), 5.0.into()])
        .unwrap();
    model.add_table(orders).unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders->Straße".into(),
            from_table: "orders".into(),
            from_column: "straßenid".into(),
            to_table: "STRASSE".into(),
            to_column: "STRASSENID".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Orders Total", "SUM(orders[amount])").unwrap();
    model
        .add_measure(
            "Orders East (default)",
            "CALCULATE([Orders Total], STRASSE[REGION] = \"East\")",
        )
        .unwrap();
    model
        .add_measure(
            "Orders East (crossfilter none)",
            "CALCULATE([Orders Total], CROSSFILTER(orders[straßenid], STRASSE[STRASSENID], NONE), STRASSE[REGION] = \"East\")",
        )
        .unwrap();

    let with_default = model
        .evaluate_measure("orders east (default)", &FilterContext::empty())
        .unwrap();
    assert_eq!(with_default, Value::from(30.0));

    let disabled = model
        .evaluate_measure("orders east (crossfilter none)", &FilterContext::empty())
        .unwrap();
    assert_eq!(
        disabled,
        Value::from(35.0),
        "CROSSFILTER(..., NONE) should disable the relationship even when identifiers differ only by Unicode case folding"
    );
}

#[test]
fn lookupvalue_resolves_unicode_identifiers_case_insensitively() {
    let mut model = DataModel::new();
    let mut streets = Table::new("Straße", vec!["StraßenId", "Region"]);
    streets.push_row(vec![1.into(), "East".into()]).unwrap();
    streets.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(streets).unwrap();

    model
        .add_measure(
            "Region 1",
            "LOOKUPVALUE(STRASSE[REGION], STRASSE[STRASSENID], 1)",
        )
        .unwrap();
    let value = model
        .evaluate_measure("REGION 1", &FilterContext::empty())
        .unwrap();
    assert_eq!(value, Value::from("East"));
}

#[test]
fn treatas_resolves_unicode_identifiers_case_insensitively() {
    // Ensure TREATAS resolves both source and target identifiers case-insensitively, including
    // Unicode-aware case folding (ß -> SS).
    let mut model = DataModel::new();

    let mut streets = Table::new("Straße", vec!["StraßenId", "Region"]);
    streets.push_row(vec![1.into(), "East".into()]).unwrap();
    streets.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(streets).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "StraßenId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 1.into(), 20.0.into()])
        .unwrap();
    orders
        .push_row(vec![102.into(), 2.into(), 5.0.into()])
        .unwrap();
    model.add_table(orders).unwrap();

    model.add_measure("Orders Total", "SUM(orders[amount])").unwrap();
    model
        .add_measure(
            "Orders East (treatas)",
            "CALCULATE([Orders Total], TREATAS(VALUES(STRASSE[STRASSENID]), Orders[straßenid]))",
        )
        .unwrap();

    let east_filter =
        FilterContext::empty().with_column_equals("STRASSE", "REGION", "East".into());
    let value = model
        .evaluate_measure("[ORDERS EAST (TREATAS)]", &east_filter)
        .unwrap();
    assert_eq!(value, Value::from(30.0));
}

#[test]
fn row_constructor_in_resolves_unicode_identifiers_case_insensitively() {
    // Ensure `(Table[Col1], Table[Col2]) IN {...}` filters resolve Unicode identifiers
    // case-insensitively (ß -> SS), and can be used for relationship propagation.
    let mut model = DataModel::new();

    let mut streets = Table::new("Straße", vec!["StraßenId", "Region"]);
    streets.push_row(vec![1.into(), "East".into()]).unwrap();
    streets.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(streets).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "StraßenId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 1.into(), 20.0.into()])
        .unwrap();
    orders
        .push_row(vec![102.into(), 2.into(), 5.0.into()])
        .unwrap();
    model.add_table(orders).unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders->Straße".into(),
            from_table: "orders".into(),
            from_column: "straßenid".into(),
            to_table: "STRASSE".into(),
            to_column: "STRASSENID".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Orders Total", "SUM(orders[amount])").unwrap();
    model
        .add_measure(
            "Orders East (row in)",
            "CALCULATE([Orders Total], (STRASSE[STRASSENID], STRASSE[REGION]) IN {(1, \"East\")})",
        )
        .unwrap();

    let value = model
        .evaluate_measure("orders east (row in)", &FilterContext::empty())
        .unwrap();
    assert_eq!(value, Value::from(30.0));
}

#[test]
fn related_resolves_unicode_identifiers_case_insensitively() {
    let mut model = DataModel::new();
    let mut streets = Table::new("Straße", vec!["StraßenId", "Region"]);
    streets.push_row(vec![1.into(), "East".into()]).unwrap();
    streets.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(streets).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "StraßenId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 1.into(), 20.0.into()])
        .unwrap();
    orders
        .push_row(vec![102.into(), 2.into(), 5.0.into()])
        .unwrap();
    model.add_table(orders).unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders->Straße".into(),
            from_table: "orders".into(),
            from_column: "straßenid".into(),
            to_table: "STRASSE".into(),
            to_column: "STRASSENID".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            r#"SUMX(orders, IF(RELATED(STRASSE[REGION]) = "East", orders[amount], 0))"#,
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(30.0));
}

#[test]
fn relatedtable_resolves_unicode_identifiers_case_insensitively() {
    let mut model = DataModel::new();
    let mut streets = Table::new("Straße", vec!["StraßenId", "Region"]);
    streets.push_row(vec![1.into(), "East".into()]).unwrap();
    streets.push_row(vec![2.into(), "West".into()]).unwrap();
    model.add_table(streets).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "StraßenId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 1.into(), 20.0.into()])
        .unwrap();
    orders
        .push_row(vec![102.into(), 2.into(), 5.0.into()])
        .unwrap();
    model.add_table(orders).unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders->Straße".into(),
            from_table: "orders".into(),
            from_column: "straßenid".into(),
            to_table: "STRASSE".into(),
            to_column: "STRASSENID".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            "SUMX(STRASSE, COUNTROWS(RELATEDTABLE(orders)))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(3_i64));
}

#[test]
fn add_table_rejects_duplicate_table_names_case_insensitively_for_unicode() {
    // `ß` uppercases to `SS`, so these two table names collide under case-insensitive matching.
    let mut model = DataModel::new();
    model
        .add_table(Table::new("Straße", vec!["Id"]))
        .expect("first table insert");
    let err = model
        .add_table(Table::new("STRASSE", vec!["Id"]))
        .unwrap_err();
    assert!(matches!(
        err,
        DaxError::DuplicateTable { table } if table == "STRASSE"
    ));
}

#[test]
fn add_table_rejects_duplicate_column_names_case_insensitively_for_unicode() {
    // `Maß` uppercases to `MASS`, so these columns collide under case-insensitive matching.
    let mut model = DataModel::new();
    let table = Table::new("T", vec!["Maß", "MASS"]);
    let err = model.add_table(table).unwrap_err();
    assert!(matches!(
        err,
        DaxError::DuplicateColumn { table, column } if table == "T" && column == "MASS"
    ));
}

#[test]
fn duplicate_measure_names_are_rejected_case_insensitively_for_unicode() {
    // `Maß` uppercases to `MASS`, so these measure names collide under case-insensitive matching.
    let mut model = DataModel::new();
    model.add_measure("Maß", "1").unwrap();
    let err = model.add_measure("MASS", "2").unwrap_err();
    assert!(matches!(err, DaxError::DuplicateMeasure { .. }));

    // Measure lookup should also be case-insensitive for Unicode names.
    let value = model
        .evaluate_measure("[MASS]", &FilterContext::empty())
        .unwrap();
    assert_eq!(value, Value::from(1.0));
}

#[test]
fn add_table_rejects_duplicate_column_names_case_insensitively() {
    let mut model = DataModel::new();
    let table = Table::new("T", vec!["Col", "col"]);
    let err = model.add_table(table).unwrap_err();
    assert!(matches!(
        err,
        DaxError::DuplicateColumn { table, column } if table == "T" && column == "col"
    ));
}

#[test]
fn pivot_resolves_identifiers_case_insensitively_and_uses_model_casing_for_headers() {
    let mut model = build_model();
    model.add_measure("Total", "SUM(Orders[Amount])").unwrap();

    let group_by = vec![GroupByColumn::new("customers", "region")];
    let measures = vec![PivotMeasure::new("Total", "[TOTAL]").unwrap()];

    let result = pivot(
        &model,
        "orders",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();

    // The pivot output should preserve the model's original casing, even when callers pass
    // mismatched identifier casing.
    assert_eq!(result.columns, vec!["Customers[Region]", "Total"]);
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("East"), Value::from(38.0)],
            vec![Value::from("West"), Value::from(5.0)],
        ]
    );
}

#[test]
fn add_table_rejects_duplicate_table_names_case_insensitively() {
    let mut model = DataModel::new();
    model.add_table(Table::new("Orders", vec!["Id"])).unwrap();
    let err = model
        .add_table(Table::new("orders", vec!["Id"]))
        .unwrap_err();
    assert!(matches!(
        err,
        DaxError::DuplicateTable { table } if table == "orders"
    ));
}

#[test]
fn duplicate_measure_names_are_rejected_case_insensitively() {
    let mut model = DataModel::new();
    model.add_measure("Total", "1").unwrap();
    // Also ensure bracketed measure names normalize to the same identifier.
    let err = model.add_measure("[total]", "2").unwrap_err();
    assert!(matches!(err, DaxError::DuplicateMeasure { .. }));
}

#[test]
fn dax_engine_resolves_mixed_case_table_and_column_refs() {
    let model = build_model();
    let engine = DaxEngine::new();

    let value = engine
        .evaluate(
            &model,
            "SUM(orders[amount])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(43.0));
}

#[test]
fn userelationship_resolves_relationship_case_insensitively() {
    // Start with an inactive relationship, so USERELATIONSHIP is required for propagation.
    let mut model = build_model_with_relationship_active(false);

    model.add_measure(
        "Orders East (inactive)",
        "CALCULATE(SUM(orders[amount]), customers[REGION] = \"East\")",
    )
    .unwrap();

    model.add_measure(
        "Orders East (userelationship)",
        "CALCULATE(SUM(orders[amount]), USERELATIONSHIP(orders[customerid], CUSTOMERS[CUSTOMERID]), customers[REGION] = \"East\")",
    )
    .unwrap();

    let no_rel = model
        .evaluate_measure("orders east (inactive)", &FilterContext::empty())
        .unwrap();
    assert_eq!(no_rel, Value::from(43.0));

    let with_rel = model
        .evaluate_measure("ORDERS EAST (USERELATIONSHIP)", &FilterContext::empty())
        .unwrap();
    assert_eq!(with_rel, Value::from(38.0));
}
