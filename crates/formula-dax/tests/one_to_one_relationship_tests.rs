use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, DaxError, FilterContext, Relationship,
    RowContext, Table, Value,
};

use pretty_assertions::assert_eq;

#[test]
fn one_to_one_relationship_supports_related() {
    let mut model = DataModel::new();

    let mut people = Table::new("People", vec!["PersonId", "Name"]);
    people
        .push_row(vec![1.into(), "Alice".into()])
        .unwrap();
    people.push_row(vec![2.into(), "Bob".into()]).unwrap();
    model.add_table(people).unwrap();

    let mut details = Table::new("Details", vec!["PersonId", "Age"]);
    details.push_row(vec![1.into(), 30.into()]).unwrap();
    details.push_row(vec![2.into(), 40.into()]).unwrap();
    model.add_table(details).unwrap();

    model
        .add_relationship(Relationship {
            name: "People_Details".into(),
            from_table: "People".into(),
            from_column: "PersonId".into(),
            to_table: "Details".into(),
            to_column: "PersonId".into(),
            cardinality: Cardinality::OneToOne,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_calculated_column("People", "Age", "RELATED(Details[Age])")
        .unwrap();

    let people = model.table("People").unwrap();
    let values: Vec<Value> = (0..people.row_count())
        .map(|row| people.value(row, "Age").unwrap())
        .collect();
    assert_eq!(values, vec![30.into(), 40.into()]);
}

#[test]
fn one_to_one_relationship_rejects_duplicate_from_keys() {
    let mut model = DataModel::new();

    let mut from = Table::new("From", vec!["Id"]);
    from.push_row(vec![1.into()]).unwrap();
    from.push_row(vec![1.into()]).unwrap();
    model.add_table(from).unwrap();

    let mut to = Table::new("To", vec!["Id"]);
    to.push_row(vec![1.into()]).unwrap();
    to.push_row(vec![2.into()]).unwrap();
    model.add_table(to).unwrap();

    let err = model
        .add_relationship(Relationship {
            name: "From_To".into(),
            from_table: "From".into(),
            from_column: "Id".into(),
            to_table: "To".into(),
            to_column: "Id".into(),
            cardinality: Cardinality::OneToOne,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap_err();

    match err {
        DaxError::NonUniqueKey { table, column, value } => {
            assert_eq!(table, "From");
            assert_eq!(column, "Id");
            assert_eq!(value, 1.into());
        }
        other => panic!("expected NonUniqueKey, got {other:?}"),
    }
}

#[test]
fn one_to_one_bidirectional_relationship_propagates_filters_both_directions() {
    let mut model = DataModel::new();

    let mut a = Table::new("A", vec!["Id", "Group"]);
    a.push_row(vec![1.into(), "G1".into()]).unwrap();
    a.push_row(vec![2.into(), "G2".into()]).unwrap();
    model.add_table(a).unwrap();

    let mut b = Table::new("B", vec!["Id", "Flag"]);
    b.push_row(vec![1.into(), "X".into()]).unwrap();
    b.push_row(vec![2.into(), "Y".into()]).unwrap();
    model.add_table(b).unwrap();

    model
        .add_relationship(Relationship {
            name: "A_B".into(),
            from_table: "A".into(),
            from_column: "Id".into(),
            to_table: "B".into(),
            to_column: "Id".into(),
            cardinality: Cardinality::OneToOne,
            cross_filter_direction: CrossFilterDirection::Both,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    let engine = DaxEngine::new();

    let filter_a = FilterContext::empty().with_column_equals("A", "Group", "G1".into());
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(B)", &filter_a, &RowContext::default())
            .unwrap(),
        1.into()
    );

    let filter_b = FilterContext::empty().with_column_equals("B", "Flag", "Y".into());
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(A)", &filter_b, &RowContext::default())
            .unwrap(),
        1.into()
    );
}

#[test]
fn one_to_one_insert_row_rejects_duplicate_from_keys() {
    let mut model = DataModel::new();

    let mut from = Table::new("From", vec!["Id", "Label"]);
    from.push_row(vec![1.into(), "One".into()]).unwrap();
    from.push_row(vec![2.into(), "Two".into()]).unwrap();
    model.add_table(from).unwrap();

    let mut to = Table::new("To", vec!["Id", "Value"]);
    to.push_row(vec![1.into(), "A".into()]).unwrap();
    to.push_row(vec![2.into(), "B".into()]).unwrap();
    model.add_table(to).unwrap();

    model
        .add_relationship(Relationship {
            name: "From_To".into(),
            from_table: "From".into(),
            from_column: "Id".into(),
            to_table: "To".into(),
            to_column: "Id".into(),
            cardinality: Cardinality::OneToOne,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    let err = model
        .insert_row("From", vec![1.into(), "Duplicate".into()])
        .unwrap_err();
    assert!(matches!(err, DaxError::NonUniqueKey { .. }));

    assert_eq!(model.table("From").unwrap().row_count(), 2);
}

