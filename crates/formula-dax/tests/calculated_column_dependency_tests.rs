use formula_dax::{DataModel, DaxError, Table, Value};

#[test]
fn insert_row_computes_dependent_calculated_columns() {
    let mut model = DataModel::new();

    let mut t = Table::new("T", vec!["a"]);
    t.push_row(vec![1.into()]).unwrap();
    model.add_table(t).unwrap();

    model
        .add_calculated_column("T", "A", "[a] + 1")
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
    // Base column: a
    // Calculated columns:
    //   A = [a] + 1
    //   B = [A] + 1
    let mut t = Table::new("T", vec!["a", "A", "B"]);
    t.push_row(vec![1.into(), 2.into(), 3.into()]).unwrap();
    model.add_table(t).unwrap();

    // Register definitions out-of-order on purpose: B before A.
    model.add_calculated_column_definition("T", "B", "[A] + 1")
        .unwrap();
    model.add_calculated_column_definition("T", "A", "[a] + 1")
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

    let mut t = Table::new("T", vec!["a", "A", "B"]);
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
