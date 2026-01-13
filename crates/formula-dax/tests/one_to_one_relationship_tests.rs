use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, DaxError, FilterContext, Relationship,
    RowContext, Table, Value,
};
use pretty_assertions::assert_eq;

#[test]
fn can_add_one_to_one_relationship_with_unique_keys() {
    let mut model = DataModel::new();

    let mut people = Table::new("People", vec!["PersonId", "Name"]);
    people.push_row(vec![1.into(), "Alice".into()]).unwrap();
    people.push_row(vec![2.into(), "Bob".into()]).unwrap();
    model.add_table(people).unwrap();

    let mut passports = Table::new("Passports", vec!["PersonId", "PassportNo"]);
    passports.push_row(vec![1.into(), "P1".into()]).unwrap();
    passports.push_row(vec![2.into(), "P2".into()]).unwrap();
    model.add_table(passports).unwrap();

    model
        .add_relationship(Relationship {
            name: "Passports_People".into(),
            from_table: "Passports".into(),
            from_column: "PersonId".into(),
            to_table: "People".into(),
            to_column: "PersonId".into(),
            cardinality: Cardinality::OneToOne,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();
}

#[test]
fn one_to_one_relationship_rejects_non_unique_to_side_keys() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Id"]);
    dim.push_row(vec![1.into()]).unwrap();
    dim.push_row(vec![1.into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id"]);
    fact.push_row(vec![1.into()]).unwrap();
    model.add_table(fact).unwrap();

    let err = model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Id".into(),
            to_table: "Dim".into(),
            to_column: "Id".into(),
            cardinality: Cardinality::OneToOne,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap_err();

    assert!(matches!(err, DaxError::NonUniqueKey { .. }));
}

#[test]
fn one_to_one_relationship_supports_related() {
    let mut model = DataModel::new();

    let mut people = Table::new("People", vec!["PersonId", "Name"]);
    people.push_row(vec![1.into(), "Alice".into()]).unwrap();
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
fn one_to_one_filter_propagation_works() {
    let mut model = DataModel::new();

    let mut people = Table::new("People", vec!["PersonId", "Name", "Region"]);
    people
        .push_row(vec![1.into(), "Alice".into(), "East".into()])
        .unwrap();
    people
        .push_row(vec![2.into(), "Bob".into(), "West".into()])
        .unwrap();
    people
        .push_row(vec![3.into(), "Carol".into(), "East".into()])
        .unwrap();
    model.add_table(people).unwrap();

    let mut passports = Table::new("Passports", vec!["PersonId", "PassportNo"]);
    passports.push_row(vec![1.into(), "P1".into()]).unwrap();
    passports.push_row(vec![2.into(), "P2".into()]).unwrap();
    model.add_table(passports).unwrap();

    model
        .add_relationship(Relationship {
            name: "Passports_People".into(),
            from_table: "Passports".into(),
            from_column: "PersonId".into(),
            to_table: "People".into(),
            to_column: "PersonId".into(),
            cardinality: Cardinality::OneToOne,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_measure("Passport Count", "COUNTROWS(Passports)")
        .unwrap();

    let all = model
        .evaluate_measure("Passport Count", &FilterContext::empty())
        .unwrap();
    assert_eq!(all, Value::from(2i64));

    let east_filter =
        FilterContext::empty().with_column_equals("People", "Region", "East".into());
    let east = model.evaluate_measure("Passport Count", &east_filter).unwrap();
    assert_eq!(east, Value::from(1i64));
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

#[test]
fn one_to_one_insert_row_rejects_duplicate_to_keys() {
    let mut model = DataModel::new();

    let mut from = Table::new("From", vec!["Id"]);
    from.push_row(vec![1.into()]).unwrap();
    from.push_row(vec![2.into()]).unwrap();
    model.add_table(from).unwrap();

    let mut to = Table::new("To", vec!["Id", "Label"]);
    to.push_row(vec![1.into(), "One".into()]).unwrap();
    to.push_row(vec![2.into(), "Two".into()]).unwrap();
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
        .insert_row("To", vec![1.into(), "Duplicate".into()])
        .unwrap_err();
    assert!(matches!(err, DaxError::NonUniqueKey { .. }));

    assert_eq!(model.table("To").unwrap().row_count(), 2);
}

#[test]
fn one_to_one_related_returns_expected_values() {
    let mut model = DataModel::new();

    let mut people = Table::new("People", vec!["PersonId", "Name"]);
    people.push_row(vec![1.into(), "Alice".into()]).unwrap();
    people.push_row(vec![2.into(), "Bob".into()]).unwrap();
    model.add_table(people).unwrap();

    let mut passports = Table::new("Passports", vec!["PersonId", "PassportNo"]);
    passports.push_row(vec![1.into(), "P1".into()]).unwrap();
    passports.push_row(vec![2.into(), "P2".into()]).unwrap();
    model.add_table(passports).unwrap();

    model
        .add_relationship(Relationship {
            name: "Passports_People".into(),
            from_table: "Passports".into(),
            from_column: "PersonId".into(),
            to_table: "People".into(),
            to_column: "PersonId".into(),
            cardinality: Cardinality::OneToOne,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_calculated_column("Passports", "PersonName", "RELATED(People[Name])")
        .unwrap();

    let passports = model.table("Passports").unwrap();
    let names: Vec<Value> = (0..passports.row_count())
        .map(|row| passports.value(row, "PersonName").unwrap())
        .collect();

    assert_eq!(names, vec![Value::from("Alice"), Value::from("Bob")]);
}

#[test]
fn one_to_one_relatedtable_returns_expected_rows() {
    let mut model = DataModel::new();

    let mut people = Table::new("People", vec!["PersonId", "Name"]);
    people.push_row(vec![1.into(), "Alice".into()]).unwrap();
    people.push_row(vec![2.into(), "Bob".into()]).unwrap();
    people.push_row(vec![3.into(), "Carol".into()]).unwrap();
    model.add_table(people).unwrap();

    let mut passports = Table::new("Passports", vec!["PersonId", "PassportNo"]);
    passports.push_row(vec![1.into(), "P1".into()]).unwrap();
    passports.push_row(vec![2.into(), "P2".into()]).unwrap();
    model.add_table(passports).unwrap();

    model
        .add_relationship(Relationship {
            name: "Passports_People".into(),
            from_table: "Passports".into(),
            from_column: "PersonId".into(),
            to_table: "People".into(),
            to_column: "PersonId".into(),
            cardinality: Cardinality::OneToOne,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_calculated_column(
            "People",
            "Passport Row Count",
            "COUNTROWS(RELATEDTABLE(Passports))",
        )
        .unwrap();

    let people = model.table("People").unwrap();
    let counts: Vec<Value> = (0..people.row_count())
        .map(|row| people.value(row, "Passport Row Count").unwrap())
        .collect();
    assert_eq!(counts, vec![1.into(), 1.into(), 0.into()]);
}

#[test]
fn one_to_one_relationship_enforces_referential_integrity_on_add() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Id"]);
    dim.push_row(vec![1.into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id"]);
    // Non-BLANK key that does not exist in Dim.
    fact.push_row(vec![2.into()]).unwrap();
    model.add_table(fact).unwrap();

    let err = model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Id".into(),
            to_table: "Dim".into(),
            to_column: "Id".into(),
            cardinality: Cardinality::OneToOne,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap_err();

    assert!(matches!(err, DaxError::ReferentialIntegrityViolation { .. }));
}

#[test]
fn one_to_one_insert_row_enforces_referential_integrity() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Id"]);
    dim.push_row(vec![1.into()]).unwrap();
    dim.push_row(vec![2.into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id"]);
    fact.push_row(vec![1.into()]).unwrap();
    model.add_table(fact).unwrap();

    model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Id".into(),
            to_table: "Dim".into(),
            to_column: "Id".into(),
            cardinality: Cardinality::OneToOne,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    let err = model.insert_row("Fact", vec![999.into()]).unwrap_err();
    assert!(matches!(
        err,
        DaxError::ReferentialIntegrityViolation { .. }
    ));

    assert_eq!(model.table("Fact").unwrap().row_count(), 1);
}

#[test]
fn one_to_one_single_direction_does_not_propagate_filters_from_from_side_to_to_side() {
    let mut model = DataModel::new();

    let mut people = Table::new("People", vec!["PersonId"]);
    people.push_row(vec![1.into()]).unwrap();
    people.push_row(vec![2.into()]).unwrap();
    model.add_table(people).unwrap();

    let mut passports = Table::new("Passports", vec!["PersonId", "PassportNo"]);
    passports.push_row(vec![1.into(), "P1".into()]).unwrap();
    passports.push_row(vec![2.into(), "P2".into()]).unwrap();
    model.add_table(passports).unwrap();

    model
        .add_relationship(Relationship {
            name: "Passports_People".into(),
            from_table: "Passports".into(),
            from_column: "PersonId".into(),
            to_table: "People".into(),
            to_column: "PersonId".into(),
            cardinality: Cardinality::OneToOne,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    let filter =
        FilterContext::empty().with_column_equals("Passports", "PassportNo", "P1".into());
    let value = DaxEngine::new()
        .evaluate(&model, "COUNTROWS(People)", &filter, &RowContext::default())
        .unwrap();

    // The default single-direction relationship only propagates People -> Passports.
    // Filtering the fact-side (Passports) should not restrict the dimension-side (People).
    assert_eq!(value, 2.into());
}
