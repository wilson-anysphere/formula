use formula_dax::{DataModel, DaxEngine, DaxError, FilterContext, RowContext, Table, Value};
use pretty_assertions::assert_eq;

#[test]
fn earlier_supports_common_rank_pattern() {
    let mut model = DataModel::new();
    let mut t = Table::new("T", vec!["x"]);
    t.push_row(vec![1.into()]).unwrap();
    t.push_row(vec![2.into()]).unwrap();
    t.push_row(vec![3.into()]).unwrap();
    model.add_table(t).unwrap();

    model
        .add_calculated_column(
            "T",
            "Rank",
            "COUNTROWS(FILTER(T, T[x] < EARLIER(T[x])))",
        )
        .unwrap();

    let t = model.table("T").unwrap();
    let ranks: Vec<Value> = (0..t.row_count())
        .map(|row| t.value(row, "Rank").unwrap())
        .collect();
    assert_eq!(ranks, vec![0.into(), 1.into(), 2.into()]);
}

#[test]
fn earlier_errors_without_nested_row_context() {
    let mut model = DataModel::new();
    let mut t = Table::new("T", vec!["x"]);
    t.push_row(vec![1.into()]).unwrap();
    t.push_row(vec![2.into()]).unwrap();
    t.push_row(vec![3.into()]).unwrap();
    model.add_table(t).unwrap();

    let err = model
        .add_calculated_column("T", "Bad", "EARLIER(T[x])")
        .unwrap_err();
    assert!(matches!(err, DaxError::Eval(_)));
    assert!(err.to_string().contains("EARLIER"));
}

#[test]
fn earlier_respects_restricted_values_row_context_visibility() {
    let mut model = DataModel::new();
    let mut t = Table::new("T", vec!["x", "y"]);
    t.push_row(vec![1.into(), 10.into()]).unwrap();
    t.push_row(vec![2.into(), 20.into()]).unwrap();
    model.add_table(t).unwrap();

    let err = DaxEngine::new()
        .evaluate(
            &model,
            "SUMX(VALUES(T[x]), SUMX(T, EARLIER(T[y])))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap_err();
    assert!(matches!(err, DaxError::Eval(_)));
    assert!(err
        .to_string()
        .contains("not available in the current row context"));
}
