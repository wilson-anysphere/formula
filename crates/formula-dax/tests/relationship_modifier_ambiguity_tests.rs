use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, Relationship, Table,
};

#[test]
fn userelationship_errors_when_multiple_relationships_match_same_column_pair() {
    let mut model = DataModel::new();
    model.add_table(Table::new("Dim", vec!["Id"])).unwrap();
    model.add_table(Table::new("Fact", vec!["DimId"])).unwrap();

    model
        .add_relationship(Relationship {
            name: "Rel1".to_string(),
            from_table: "Fact".to_string(),
            from_column: "DimId".to_string(),
            to_table: "Dim".to_string(),
            to_column: "Id".to_string(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();
    model
        .add_relationship(Relationship {
            name: "Rel2".to_string(),
            from_table: "Fact".to_string(),
            from_column: "DimId".to_string(),
            to_table: "Dim".to_string(),
            to_column: "Id".to_string(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: false,
            enforce_referential_integrity: false,
        })
        .unwrap();

    let engine = DaxEngine::new();
    let err = engine
        .apply_calculate_filters(
            &model,
            &FilterContext::empty(),
            &["USERELATIONSHIP(Fact[DimId], Dim[Id])"],
        )
        .unwrap_err();
    let msg = err.to_string();
    assert!(
        msg.contains("multiple relationships found"),
        "unexpected error: {msg}"
    );
    assert!(msg.contains("Rel1"), "unexpected error: {msg}");
    assert!(msg.contains("Rel2"), "unexpected error: {msg}");
}

#[test]
fn crossfilter_errors_when_multiple_relationships_match_same_column_pair() {
    let mut model = DataModel::new();
    model.add_table(Table::new("Dim", vec!["Id"])).unwrap();
    model.add_table(Table::new("Fact", vec!["DimId"])).unwrap();

    model
        .add_relationship(Relationship {
            name: "Rel1".to_string(),
            from_table: "Fact".to_string(),
            from_column: "DimId".to_string(),
            to_table: "Dim".to_string(),
            to_column: "Id".to_string(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();
    model
        .add_relationship(Relationship {
            name: "Rel2".to_string(),
            from_table: "Fact".to_string(),
            from_column: "DimId".to_string(),
            to_table: "Dim".to_string(),
            to_column: "Id".to_string(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: false,
            enforce_referential_integrity: false,
        })
        .unwrap();

    let engine = DaxEngine::new();
    let err = engine
        .apply_calculate_filters(
            &model,
            &FilterContext::empty(),
            &["CROSSFILTER(Fact[DimId], Dim[Id], NONE)"],
        )
        .unwrap_err();
    let msg = err.to_string();
    assert!(
        msg.contains("multiple relationships found"),
        "unexpected error: {msg}"
    );
    assert!(msg.contains("Rel1"), "unexpected error: {msg}");
    assert!(msg.contains("Rel2"), "unexpected error: {msg}");
}

