use formula_dax::{DataModel, DaxEngine, FilterContext, RowContext, Table, Value};
use pretty_assertions::assert_eq;

#[test]
fn bracket_identifiers_support_escaped_closing_brackets_in_names() {
    let mut model = DataModel::new();
    let mut table = Table::new("T", vec!["Amount]USD"]);
    table.push_row(vec![1.0.into()]).unwrap();
    table.push_row(vec![2.0.into()]).unwrap();
    model.add_table(table).unwrap();

    model
        .add_measure("Total]USD", "SUM(T[Amount]]USD])")
        .unwrap();

    let engine = DaxEngine::new();
    let v = engine
        .evaluate(
            &model,
            "SUM(T[Amount]]USD])",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(v, Value::from(3.0));

    let v = engine
        .evaluate(
            &model,
            "[Total]]USD]",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(v, Value::from(3.0));
}

