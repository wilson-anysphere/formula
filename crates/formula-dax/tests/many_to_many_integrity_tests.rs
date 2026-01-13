use formula_dax::{Cardinality, CrossFilterDirection, DataModel, DaxError, Relationship, Table};

#[test]
fn add_relationship_enforces_referential_integrity_for_many_to_many_when_enabled() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key"]);
    dim.push_row(vec![1.into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Key"]);
    fact.push_row(vec![2.into()]).unwrap();
    model.add_table(fact).unwrap();

    let err = model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Key".into(),
            to_table: "Dim".into(),
            to_column: "Key".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap_err();

    assert!(matches!(err, DaxError::ReferentialIntegrityViolation { .. }));
}

#[test]
fn add_relationship_does_not_require_unique_to_side_keys_for_many_to_many() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key"]);
    dim.push_row(vec![1.into()]).unwrap();
    dim.push_row(vec![1.into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Key"]);
    fact.push_row(vec![1.into()]).unwrap();
    model.add_table(fact).unwrap();

    model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Key".into(),
            to_table: "Dim".into(),
            to_column: "Key".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();
}

#[test]
fn add_relationship_still_rejects_non_unique_keys_for_one_to_many() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key"]);
    dim.push_row(vec![1.into()]).unwrap();
    dim.push_row(vec![1.into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Key"]);
    fact.push_row(vec![1.into()]).unwrap();
    model.add_table(fact).unwrap();

    let err = model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Key".into(),
            to_table: "Dim".into(),
            to_column: "Key".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap_err();

    assert!(matches!(err, DaxError::NonUniqueKey { .. }));
}

#[test]
fn insert_row_respects_referential_integrity_under_many_to_many() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key"]);
    dim.push_row(vec![1.into()]).unwrap();
    model.add_table(dim).unwrap();

    let fact = Table::new("Fact", vec!["Key"]);
    model.add_table(fact).unwrap();

    model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Key".into(),
            to_table: "Dim".into(),
            to_column: "Key".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    let err = model.insert_row("Fact", vec![2.into()]).unwrap_err();
    assert!(matches!(err, DaxError::ReferentialIntegrityViolation { .. }));
}

