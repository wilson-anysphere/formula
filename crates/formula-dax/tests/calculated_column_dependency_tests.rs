use formula_dax::{DataModel, DaxError, Table, Value};

#[test]
fn insert_row_computes_dependent_calculated_columns() {
    let mut model = DataModel::new();

    let mut t = Table::new("T", vec!["a0"]);
    t.push_row(vec![1.into()]).unwrap();
    model.add_table(t).unwrap();

    model
        .add_calculated_column("T", "A", "[a0] + 1")
        .unwrap();
    model
        .add_calculated_column("T", "B", "[A] + 1")
        .unwrap();

    model.insert_row("T", vec![10.into()]).unwrap();

    let t = model.table("T").unwrap();
    assert_eq!(t.row_count(), 2);
    assert_eq!(t.value(1, "A").unwrap(), Value::from(11.0));
    assert_eq!(t.value(1, "B").unwrap(), Value::from(12.0));
}

#[test]
fn insert_row_computes_calculated_columns_in_dependency_order() {
    let mut model = DataModel::new();

    // Simulate a persisted model where calculated column values are already stored in the table,
    // but definitions can be registered in any order.
    //
    // Base column: a0
    // Calculated columns:
    //   A = [a0] + 1
    //   B = [A] + 1
    let mut t = Table::new("T", vec!["a0", "A", "B"]);
    t.push_row(vec![1.into(), 2.into(), 3.into()]).unwrap();
    model.add_table(t).unwrap();

    // Register definitions out-of-order on purpose: B before A.
    model.add_calculated_column_definition("T", "B", "[A] + 1")
        .unwrap();
    model.add_calculated_column_definition("T", "A", "[a0] + 1")
        .unwrap();

    model.insert_row("T", vec![10.into()]).unwrap();

    let t = model.table("T").unwrap();
    assert_eq!(t.row_count(), 2);
    assert_eq!(t.value(1, "A").unwrap(), Value::from(11.0));
    assert_eq!(t.value(1, "B").unwrap(), Value::from(12.0));
}

#[test]
fn calculated_column_dependency_cycle_errors() {
    let mut model = DataModel::new();

    let mut t = Table::new("T", vec!["a0", "A", "B"]);
    t.push_row(vec![1.into(), Value::Blank, Value::Blank]).unwrap();
    model.add_table(t).unwrap();

    model
        .add_calculated_column_definition("T", "A", "[B] + 1")
        .unwrap();
    let err = model
        .add_calculated_column_definition("T", "B", "[A] + 1")
        .unwrap_err();

    assert!(matches!(err, DaxError::Eval(_)));
    let msg = err.to_string();
    assert!(msg.contains("dependency cycle"), "unexpected error: {msg}");
    assert!(msg.contains("T"), "unexpected error: {msg}");
}

#[test]
fn insert_row_maps_values_around_calculated_columns() {
    let mut model = DataModel::new();

    // Simulate a persisted table where calculated columns are not physically last.
    // Schema order: a0, A (calc), b0, B (calc)
    let mut t = Table::new("T", vec!["a0", "A", "b0", "B"]);
    t.push_row(vec![1.into(), 2.into(), 20.into(), 22.into()]).unwrap();
    model.add_table(t).unwrap();

    // Register definitions out-of-order to also exercise topo ordering.
    model.add_calculated_column_definition("T", "B", "[A] + [b0]")
        .unwrap();
    model.add_calculated_column_definition("T", "A", "[a0] + 1")
        .unwrap();

    // Provide only non-calculated column values in schema order (a0, b0).
    model.insert_row("T", vec![10.into(), 20.into()]).unwrap();

    let t = model.table("T").unwrap();
    assert_eq!(t.row_count(), 2);
    assert_eq!(t.value(1, "a0").unwrap(), Value::from(10.0));
    assert_eq!(t.value(1, "b0").unwrap(), Value::from(20.0));
    assert_eq!(t.value(1, "A").unwrap(), Value::from(11.0));
    assert_eq!(t.value(1, "B").unwrap(), Value::from(31.0));
}

#[test]
fn insert_row_rolls_back_on_calculated_column_error() {
    let mut model = DataModel::new();

    let mut t = Table::new("T", vec!["a0", "A"]);
    t.push_row(vec![1.into(), Value::Blank]).unwrap();
    model.add_table(t).unwrap();

    // Register a calculated column definition that will fail at evaluation time.
    model
        .add_calculated_column_definition("T", "A", "BOGUS()")
        .unwrap();

    let err = model.insert_row("T", vec![10.into()]).unwrap_err();
    assert!(matches!(err, DaxError::Eval(_)));

    // Ensure the failed insert didn't leave a partially-inserted row behind.
    let t = model.table("T").unwrap();
    assert_eq!(t.row_count(), 1);
}

#[test]
fn calculated_column_dependencies_traverse_var_bindings_and_body() {
    let mut model = DataModel::new();

    // Persisted model: calculated column storage already exists, but definitions can be registered
    // in any order.
    let mut t = Table::new("T", vec!["a0", "A", "B", "C"]);
    t.push_row(vec![1.into(), Value::Blank, Value::Blank, Value::Blank])
        .unwrap();
    model.add_table(t).unwrap();

    // Register out-of-order on purpose: C depends on both A (via VAR binding) and B (via RETURN).
    model
        .add_calculated_column_definition("T", "C", "VAR x = [A] RETURN x + [B]")
        .unwrap();
    model
        .add_calculated_column_definition("T", "A", "[a0] + 1")
        .unwrap();
    model
        .add_calculated_column_definition("T", "B", "[a0] + 2")
        .unwrap();

    model.insert_row("T", vec![10.into()]).unwrap();

    let t = model.table("T").unwrap();
    assert_eq!(t.row_count(), 2);
    assert_eq!(t.value(1, "A").unwrap(), Value::from(11.0));
    assert_eq!(t.value(1, "B").unwrap(), Value::from(12.0));
    assert_eq!(t.value(1, "C").unwrap(), Value::from(23.0));
}

#[test]
fn calculated_column_dependencies_resolve_unicode_identifiers_case_insensitively() {
    // Use a German sharp S (ß) to ensure Unicode-aware case folding (ß -> SS) is applied when:
    // - registering calculated column definitions (`add_calculated_column_definition`)
    // - resolving bracket identifiers like `[MASS]` to same-table columns in row context
    // - computing dependency order between calculated columns.
    let mut model = DataModel::new();

    // Persisted model: calculated column storage already exists, but definitions can be registered
    // in any order.
    let mut t = Table::new("Straße", vec!["Maß", "Maß2", "Quad"]);
    t.push_row(vec![1.into(), Value::Blank, Value::Blank]).unwrap();
    model.add_table(t).unwrap();

    // Register out-of-order on purpose: Quad depends on Maß2; Maß2 depends on Maß.
    //
    // Use ASCII-only identifiers (`STRASSE`, `MASS`, `MASS2`) to ensure Unicode-aware folding is
    // used during lookup.
    model
        .add_calculated_column_definition("STRASSE", "Quad", "[MASS2] * 2")
        .unwrap();
    model
        .add_calculated_column_definition("strasse", "MASS2", "[mass] + 1")
        .unwrap();

    model.insert_row("strasse", vec![10.into()]).unwrap();

    let t = model.table("STRASSE").unwrap();
    assert_eq!(t.value(1, "Maß2").unwrap(), Value::from(11.0));
    assert_eq!(t.value(1, "Quad").unwrap(), Value::from(22.0));
}

#[test]
fn add_calculated_column_resolves_unicode_identifiers_case_insensitively() {
    let mut model = DataModel::new();

    let mut t = Table::new("Straße", vec!["Maß"]);
    t.push_row(vec![1.into()]).unwrap();
    t.push_row(vec![2.into()]).unwrap();
    model.add_table(t).unwrap();

    // Refer to the table via ASCII-only spelling and refer to the source column via `[MASS]`.
    model
        .add_calculated_column("STRASSE", "Maß2", "[MASS] + 1")
        .unwrap();

    let t = model.table("strasse").unwrap();
    assert_eq!(t.row_count(), 2);
    assert_eq!(t.value(0, "Maß2").unwrap(), Value::from(2.0));
    assert_eq!(t.value(1, "MASS2").unwrap(), Value::from(3.0));
}
