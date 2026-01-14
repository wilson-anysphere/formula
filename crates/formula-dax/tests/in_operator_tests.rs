mod common;

use common::build_model;
use formula_dax::{DataModel, DaxEngine, FilterContext, RowContext, Table, Value};
use pretty_assertions::assert_eq;

#[test]
fn in_operator_scalar_table_constructor() {
    let model = build_model();
    let engine = DaxEngine::new();
    let filter = FilterContext::empty();
    let row_ctx = RowContext::default();

    let value = engine
        .evaluate(&model, "1 IN {1,2}", &filter, &row_ctx)
        .unwrap();
    assert_eq!(value, Value::from(true));

    let value = engine
        .evaluate(&model, "3 IN {1,2}", &filter, &row_ctx)
        .unwrap();
    assert_eq!(value, Value::from(false));
}

#[test]
fn in_operator_calculate_column_filter() {
    let model = build_model();
    let engine = DaxEngine::new();
    let filter = FilterContext::empty();
    let row_ctx = RowContext::default();

    let value = engine
        .evaluate(
            &model,
            "CALCULATE(COUNTROWS(Orders), Orders[CustomerId] IN {1,3})",
            &filter,
            &row_ctx,
        )
        .unwrap();
    assert_eq!(value, Value::from(3i64));
}

#[test]
fn in_operator_scalar_table_expression_rhs() {
    let model = build_model();
    let engine = DaxEngine::new();
    let filter = FilterContext::empty();
    let row_ctx = RowContext::default();

    // Physical one-column table expressions like VALUES(column) should work.
    let value = engine
        .evaluate(&model, "1 IN VALUES(Orders[CustomerId])", &filter, &row_ctx)
        .unwrap();
    assert_eq!(value, Value::from(true));

    let value = engine
        .evaluate(&model, "4 IN VALUES(Orders[CustomerId])", &filter, &row_ctx)
        .unwrap();
    assert_eq!(value, Value::from(false));

    // Virtual one-column tables should also work (e.g. SUMMARIZE).
    let value = engine
        .evaluate(&model, "1 IN SUMMARIZE(Orders, Orders[CustomerId])", &filter, &row_ctx)
        .unwrap();
    assert_eq!(value, Value::from(true));

    let value = engine
        .evaluate(&model, "4 IN SUMMARIZE(Orders, Orders[CustomerId])", &filter, &row_ctx)
        .unwrap();
    assert_eq!(value, Value::from(false));
}

#[test]
fn in_operator_calculate_filter_rhs_table_expression() {
    let model = build_model();
    let engine = DaxEngine::new();
    let filter = FilterContext::empty();
    let row_ctx = RowContext::default();

    // IN inside CALCULATE can use a one-column table expression, not only a table constructor.
    let value = engine
        .evaluate(
            &model,
            "CALCULATE(COUNTROWS(Orders), Orders[CustomerId] IN SUMMARIZE(FILTER(Customers, Customers[Region] = \"East\"), Customers[CustomerId]))",
            &filter,
            &row_ctx,
        )
        .unwrap();
    assert_eq!(value, Value::from(3i64));
}

#[test]
fn in_operator_calculate_filter_rhs_physical_mask_table_expression() {
    // Regression test: ensure CALCULATE's `col IN <table expr>` filter handling supports
    // `TableResult::PhysicalMask` without materializing a large `Vec<usize>`.
    //
    // `FILTER` produces `PhysicalMask` for small tables (dense representation threshold is 0 when
    // row_count < 64), so this is deterministic.
    let mut model = DataModel::new();
    let mut t = Table::new("T", vec!["Id"]);
    for id in 1..=5i64 {
        t.push_row(vec![id.into()]).unwrap();
    }
    model.add_table(t).unwrap();

    let engine = DaxEngine::new();
    let filter = FilterContext::empty();
    let row_ctx = RowContext::default();

    let value = engine
        .evaluate(
            &model,
            "CALCULATE(COUNTROWS(T), T[Id] IN FILTER(T, T[Id] <= 3))",
            &filter,
            &row_ctx,
        )
        .unwrap();
    assert_eq!(value, Value::from(3i64));
}

#[test]
fn in_operator_table_expression_requires_one_column() {
    let model = build_model();
    let engine = DaxEngine::new();
    let filter = FilterContext::empty();
    let row_ctx = RowContext::default();

    let err = engine
        .evaluate(&model, "1 IN Orders", &filter, &row_ctx)
        .unwrap_err();
    let message = err.to_string();
    assert!(message.contains("one-column"));
}

#[test]
fn in_operator_calculate_filter_rhs_physical_mask() {
    // Regression coverage: `IN` should support RHS table expressions that evaluate to a
    // `TableResult::PhysicalMask` (dense row set). This exercises the `IN` HashSet extraction path
    // for bitmap-backed physical tables.
    let mut model = DataModel::new();

    let mut keys = Table::new("Keys", vec!["Value"]);
    for i in 1..=256 {
        keys.push_row(vec![i.into()]).unwrap();
    }
    model.add_table(keys).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "CustomerId"]);
    orders.push_row(vec![100.into(), 1.into()]).unwrap();
    orders.push_row(vec![101.into(), 250.into()]).unwrap();
    model.add_table(orders).unwrap();

    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            "CALCULATE(COUNTROWS(Orders), Orders[CustomerId] IN FILTER(Keys, Keys[Value] <= 200))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    // Only the CustomerId=1 row should match the IN filter.
    assert_eq!(value, Value::from(1i64));
}

#[test]
fn in_operator_calculate_filter_rhs_physical_mask_same_table() {
    // Also cover the case where the RHS filter is evaluated over the same physical table as the
    // LHS column. This should not introduce dependency cycles and still needs to handle the
    // PhysicalMask representation.
    let mut model = DataModel::new();
    let mut t = Table::new("T", vec!["Value"]);
    for i in 1..=10 {
        t.push_row(vec![i.into()]).unwrap();
    }
    model.add_table(t).unwrap();

    let value = DaxEngine::new()
        .evaluate(
            &model,
            "CALCULATE(COUNTROWS(T), T[Value] IN FILTER(T, T[Value] > 5))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from(5i64));
}

#[test]
fn in_operator_scalar_rhs_physical_mask() {
    // Similar to `in_operator_calculate_filter_rhs_physical_mask`, but exercises the scalar
    // evaluation path for `IN` directly (not through CALCULATE filter parsing).
    let mut model = DataModel::new();

    let mut keys = Table::new("Keys", vec!["Value"]);
    for i in 1..=256 {
        keys.push_row(vec![i.into()]).unwrap();
    }
    model.add_table(keys).unwrap();

    let engine = DaxEngine::new();
    assert_eq!(
        engine
            .evaluate(
                &model,
                "1 IN FILTER(Keys, Keys[Value] <= 200)",
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
                "250 IN FILTER(Keys, Keys[Value] <= 200)",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        Value::from(false)
    );
}
