use formula_dax::{DataModel, DaxEngine, DaxError, FilterContext, RowContext, Table, Value};
use pretty_assertions::assert_eq;

#[test]
fn lookupvalue_simple_table_lookup_returns_expected_value() {
    let mut model = DataModel::new();
    let mut dim = Table::new("Dim", vec!["Id", "Name"]);
    dim.push_row(vec![1.into(), "Alice".into()]).unwrap();
    dim.push_row(vec![2.into(), "Bob".into()]).unwrap();
    model.add_table(dim).unwrap();

    let engine = DaxEngine::new();
    let value = engine
        .evaluate(
            &model,
            "LOOKUPVALUE(Dim[Name], Dim[Id], 2)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(value, Value::from("Bob"));
}

#[test]
fn lookupvalue_no_match_returns_blank_or_alternate() {
    let mut model = DataModel::new();
    let mut dim = Table::new("Dim", vec!["Id", "Name"]);
    dim.push_row(vec![1.into(), "Alice".into()]).unwrap();
    model.add_table(dim).unwrap();

    let engine = DaxEngine::new();

    let blank = engine
        .evaluate(
            &model,
            "LOOKUPVALUE(Dim[Name], Dim[Id], 999)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(blank, Value::Blank);

    let alt = engine
        .evaluate(
            &model,
            "LOOKUPVALUE(Dim[Name], Dim[Id], 999, \"Unknown\")",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(alt, Value::from("Unknown"));
}

#[test]
fn lookupvalue_duplicate_matches_return_error_or_value_when_identical() {
    let engine = DaxEngine::new();

    // Duplicate key with different result values: should error.
    let mut model = DataModel::new();
    let mut dim = Table::new("Dim", vec!["Id", "Name"]);
    dim.push_row(vec![1.into(), "Alice".into()]).unwrap();
    dim.push_row(vec![1.into(), "Alicia".into()]).unwrap();
    model.add_table(dim).unwrap();

    let err = engine
        .evaluate(
            &model,
            "LOOKUPVALUE(Dim[Name], Dim[Id], 1)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap_err();
    assert!(matches!(err, DaxError::Eval(_)));

    // Duplicate key with identical result values: should return the value.
    let mut model = DataModel::new();
    let mut dim = Table::new("Dim", vec!["Id", "Name"]);
    dim.push_row(vec![1.into(), "Alice".into()]).unwrap();
    dim.push_row(vec![1.into(), "Alice".into()]).unwrap();
    model.add_table(dim).unwrap();

    let value = engine
        .evaluate(
            &model,
            "LOOKUPVALUE(Dim[Name], Dim[Id], 1)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, Value::from("Alice"));
}

