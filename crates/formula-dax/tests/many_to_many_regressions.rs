use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, FilterContext, Relationship, Table, Value,
};
use pretty_assertions::assert_eq;

#[test]
fn insert_row_updates_m2m_to_index() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "Old".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "Amount"]);
    fact.push_row(vec![1.into(), 1.into(), 10.0.into()]).unwrap();
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

    model.add_measure("Total Amount", "SUM(Fact[Amount])").unwrap();

    // Before the insert, Dim[Attr] = "New" has no matching row, so the fact row is filtered out.
    let new_attr_filter = FilterContext::empty().with_column_equals("Dim", "Attr", "New".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &new_attr_filter).unwrap(),
        Value::Blank
    );

    // Insert a new Dim row that reuses Key=1 (valid for ManyToMany) but has a different attribute.
    model
        .insert_row("Dim", vec![1.into(), "New".into()])
        .unwrap();

    // Filtering to the newly-inserted attribute should still keep the fact row (Key=1) allowed.
    assert_eq!(
        model.evaluate_measure("Total Amount", &new_attr_filter).unwrap(),
        10.0.into()
    );

    // `RELATED` is ambiguous under ManyToMany when the key matches multiple Dim rows.
    // The engine should surface an error rather than choosing a row silently.
    let err = model
        .add_calculated_column("Fact", "Related Attr", "RELATED(Dim[Attr])")
        .unwrap_err();
    let msg = err.to_string().to_ascii_lowercase();
    assert!(
        msg.contains("ambig") || msg.contains("multiple") || msg.contains("more than one"),
        "unexpected RELATED error with duplicate keys: {err}"
    );
}

#[test]
fn userelationship_override_works_with_m2m() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["KeyA", "KeyB", "Attr"]);
    dim.push_row(vec![1.into(), 10.into(), "A".into()]).unwrap();
    dim.push_row(vec![2.into(), 20.into(), "B".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "KeyA", "KeyB", "Amount"]);
    // Each fact row "crosses" the keys so the active vs. inactive relationship produces
    // different results under the same Dim filter.
    fact.push_row(vec![1.into(), 1.into(), 20.into(), 100.0.into()])
        .unwrap();
    fact.push_row(vec![2.into(), 2.into(), 10.into(), 200.0.into()])
        .unwrap();
    model.add_table(fact).unwrap();

    model
        .add_relationship(Relationship {
            name: "Fact_Dim_KeyA".into(),
            from_table: "Fact".into(),
            from_column: "KeyA".into(),
            to_table: "Dim".into(),
            to_column: "KeyA".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();
    model
        .add_relationship(Relationship {
            name: "Fact_Dim_KeyB".into(),
            from_table: "Fact".into(),
            from_column: "KeyB".into(),
            to_table: "Dim".into(),
            to_column: "KeyB".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: false,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
    model
        .add_measure(
            "Total via KeyB",
            "CALCULATE([Total], USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
        )
        .unwrap();

    let filter_a = FilterContext::empty().with_column_equals("Dim", "Attr", "A".into());
    assert_eq!(model.evaluate_measure("Total", &filter_a).unwrap(), 100.0.into());
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &filter_a).unwrap(),
        200.0.into()
    );

    // Ensure the override disables the default active relationship for the table pair rather
    // than applying both relationships simultaneously (which would intersect and remove all
    // fact rows in this setup).
    let filter_b = FilterContext::empty().with_column_equals("Dim", "Attr", "B".into());
    assert_eq!(model.evaluate_measure("Total", &filter_b).unwrap(), 200.0.into());
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &filter_b).unwrap(),
        100.0.into()
    );
}

#[test]
fn blank_foreign_keys_in_m2m_flow_to_blank_dimension_member_when_allowed() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "Amount"]);
    fact.push_row(vec![1.into(), 1.into(), 10.0.into()]).unwrap();
    fact.push_row(vec![2.into(), Value::Blank, 7.0.into()])
        .unwrap();
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

    model.add_measure("Total Amount", "SUM(Fact[Amount])").unwrap();

    // Filtering the dimension to BLANK should include facts whose FK is BLANK via the implicit
    // blank dimension member.
    let blank_attr = FilterContext::empty().with_column_equals("Dim", "Attr", Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total Amount", &blank_attr).unwrap(),
        7.0.into()
    );

    // Filtering to a non-blank dimension value should exclude those fact rows.
    let attr_a = FilterContext::empty().with_column_equals("Dim", "Attr", "A".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &attr_a).unwrap(),
        10.0.into()
    );
}

#[test]
fn insert_row_updates_m2m_from_index() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "Amount"]);
    fact.push_row(vec![1.into(), 1.into(), 10.0.into()]).unwrap();
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

    model.add_measure("Total Amount", "SUM(Fact[Amount])").unwrap();

    let a_filter = FilterContext::empty().with_column_equals("Dim", "Attr", "A".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &a_filter).unwrap(),
        10.0.into()
    );

    // Insert a new fact row after the relationship is defined and ensure propagation picks it up.
    model
        .insert_row("Fact", vec![2.into(), 1.into(), 5.0.into()])
        .unwrap();
    assert_eq!(
        model.evaluate_measure("Total Amount", &a_filter).unwrap(),
        15.0.into()
    );
}

#[test]
fn insert_row_can_resolve_unmatched_facts_and_updates_blank_member() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    // Start with one matched and one unmatched fact key.
    let mut fact = Table::new("Fact", vec!["Id", "Key", "Amount"]);
    fact.push_row(vec![1.into(), 1.into(), 10.0.into()]).unwrap();
    fact.push_row(vec![2.into(), 999.into(), 7.0.into()]).unwrap();
    model.add_table(fact).unwrap();

    // Allow unmatched facts so the virtual blank member is materialized.
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    model.add_measure("Total Amount", "SUM(Fact[Amount])").unwrap();

    let blank_attr = FilterContext::empty().with_column_equals("Dim", "Attr", Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total Amount", &blank_attr).unwrap(),
        7.0.into()
    );

    // Insert a Dim row for the previously-unmatched key. This should move the fact row out of the
    // virtual blank member and under the new Dim row.
    model
        .insert_row("Dim", vec![999.into(), "New".into()])
        .unwrap();

    assert_eq!(
        model.evaluate_measure("Total Amount", &blank_attr).unwrap(),
        Value::Blank
    );

    let new_attr = FilterContext::empty().with_column_equals("Dim", "Attr", "New".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &new_attr).unwrap(),
        7.0.into()
    );
}

#[test]
fn related_respects_userelationship_overrides_with_m2m() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["KeyA", "KeyB", "Attr"]);
    dim.push_row(vec![1.into(), 10.into(), "RowA".into()]).unwrap();
    dim.push_row(vec![2.into(), 20.into(), "RowB".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "KeyA", "KeyB"]);
    // Cross the keys so the active vs. USERELATIONSHIP-overridden relationship produces
    // different RELATED values.
    fact.push_row(vec![100.into(), 1.into(), 20.into()]).unwrap();
    fact.push_row(vec![101.into(), 2.into(), 10.into()]).unwrap();
    model.add_table(fact).unwrap();

    model
        .add_relationship(Relationship {
            name: "Fact_Dim_KeyA".into(),
            from_table: "Fact".into(),
            from_column: "KeyA".into(),
            to_table: "Dim".into(),
            to_column: "KeyA".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();
    model
        .add_relationship(Relationship {
            name: "Fact_Dim_KeyB".into(),
            from_table: "Fact".into(),
            from_column: "KeyB".into(),
            to_table: "Dim".into(),
            to_column: "KeyB".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: false,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_calculated_column("Fact", "Attr via active", "RELATED(Dim[Attr])")
        .unwrap();
    model
        .add_calculated_column(
            "Fact",
            "Attr via KeyB",
            "CALCULATE(RELATED(Dim[Attr]), USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
        )
        .unwrap();

    let fact = model.table("Fact").unwrap();
    assert_eq!(fact.value(0, "Attr via active").unwrap(), "RowA".into());
    assert_eq!(fact.value(0, "Attr via KeyB").unwrap(), "RowB".into());
    assert_eq!(fact.value(1, "Attr via active").unwrap(), "RowB".into());
    assert_eq!(fact.value(1, "Attr via KeyB").unwrap(), "RowA".into());
}

#[test]
fn insert_row_updates_inactive_m2m_indexes_used_by_userelationship() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["KeyA", "KeyB", "Attr"]);
    dim.push_row(vec![1.into(), 10.into(), "A".into()]).unwrap();
    dim.push_row(vec![2.into(), 20.into(), "B".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "KeyA", "KeyB", "Amount"]);
    fact.push_row(vec![1.into(), 1.into(), 10.into(), 5.0.into()])
        .unwrap();
    fact.push_row(vec![2.into(), 2.into(), 20.into(), 7.0.into()])
        .unwrap();
    model.add_table(fact).unwrap();

    // Active relationship on KeyA.
    model
        .add_relationship(Relationship {
            name: "Fact_Dim_KeyA".into(),
            from_table: "Fact".into(),
            from_column: "KeyA".into(),
            to_table: "Dim".into(),
            to_column: "KeyA".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    // Inactive relationship on KeyB, only enabled via USERELATIONSHIP.
    model
        .add_relationship(Relationship {
            name: "Fact_Dim_KeyB".into(),
            from_table: "Fact".into(),
            from_column: "KeyB".into(),
            to_table: "Dim".into(),
            to_column: "KeyB".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: false,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
    model
        .add_measure(
            "Total via KeyB",
            "CALCULATE([Total], USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
        )
        .unwrap();

    // Touch the inactive relationship so any lazy index construction happens before inserts.
    let a_filter = FilterContext::empty().with_column_equals("Dim", "Attr", "A".into());
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &a_filter).unwrap(),
        5.0.into()
    );

    // Insert a new Dim row and a new Fact row that only match via the inactive relationship.
    model
        .insert_row("Dim", vec![3.into(), 30.into(), "C".into()])
        .unwrap();

    let c_filter = FilterContext::empty().with_column_equals("Dim", "Attr", "C".into());
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &c_filter).unwrap(),
        Value::Blank
    );

    model
        .insert_row("Fact", vec![3.into(), 1.into(), 30.into(), 11.0.into()])
        .unwrap();

    // The inserted row should be visible through USERELATIONSHIP, meaning the inactive
    // relationship indexes were incrementally updated.
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &c_filter).unwrap(),
        11.0.into()
    );

    // The default active relationship should not accidentally include it under the same Dim filter.
    assert_eq!(model.evaluate_measure("Total", &c_filter).unwrap(), Value::Blank);
}
