use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, FilterContext, Relationship, Table,
};

#[test]
fn many_to_many_userelationship_activates_inactive_relationship_and_overrides_active() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["IdA", "IdB"]);
    dim.push_row(vec![1.into(), 10.into()]).unwrap();
    dim.push_row(vec![2.into(), 20.into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["KeyA", "KeyB", "Amount"]);
    // Fact rows intentionally cross the (IdA, IdB) pairs in Dim so that filtering by Dim[IdB]
    // produces different results depending on which relationship is active.
    fact.push_row(vec![1.into(), 20.into(), 5.0.into()]).unwrap();
    fact.push_row(vec![2.into(), 10.into(), 7.0.into()]).unwrap();
    model.add_table(fact).unwrap();

    // Relationship A (active): Fact[KeyA] -> Dim[IdA]
    model
        .add_relationship(Relationship {
            name: "Fact_Dim_A".into(),
            from_table: "Fact".into(),
            from_column: "KeyA".into(),
            to_table: "Dim".into(),
            to_column: "IdA".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    // Relationship B (inactive): Fact[KeyB] -> Dim[IdB]
    model
        .add_relationship(Relationship {
            name: "Fact_Dim_B".into(),
            from_table: "Fact".into(),
            from_column: "KeyB".into(),
            to_table: "Dim".into(),
            to_column: "IdB".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: false,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
    model
        .add_measure(
            "Total via B",
            "CALCULATE([Total], USERELATIONSHIP(Fact[KeyB], Dim[IdB]))",
        )
        .unwrap();

    // With filter Dim[IdB] = 10:
    // - [Total] should use active relationship A and return 5 (IdA=1 -> Fact KeyA=1 row).
    // - [Total via B] should activate relationship B and return 7 (Fact KeyB=10 row).
    let idb_10_filter = FilterContext::empty().with_column_equals("Dim", "IdB", 10.into());
    assert_eq!(model.evaluate_measure("Total", &idb_10_filter).unwrap(), 5.0.into());
    assert_eq!(
        model
            .evaluate_measure("Total via B", &idb_10_filter)
            .unwrap(),
        7.0.into()
    );

    // With filter Dim[IdB] = 20:
    // - [Total] == 7 (IdA=2 -> Fact KeyA=2 row)
    // - [Total via B] == 5 (Fact KeyB=20 row)
    let idb_20_filter = FilterContext::empty().with_column_equals("Dim", "IdB", 20.into());
    assert_eq!(model.evaluate_measure("Total", &idb_20_filter).unwrap(), 7.0.into());
    assert_eq!(
        model
            .evaluate_measure("Total via B", &idb_20_filter)
            .unwrap(),
        5.0.into()
    );
}

