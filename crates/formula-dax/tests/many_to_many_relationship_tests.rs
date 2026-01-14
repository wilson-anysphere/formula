use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, DaxError, FilterContext, Relationship,
    RowContext, Table, Value,
};

#[test]
fn many_to_many_relationship_can_be_added_with_duplicate_to_keys() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![1.into(), "B".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key"]);
    fact.push_row(vec![10.into(), 1.into()]).unwrap();
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
fn many_to_many_insert_row_allows_duplicate_to_keys_and_updates_index() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "Amount"]);
    fact.push_row(vec![10.into(), 1.into(), 5.0.into()]).unwrap();
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

    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();

    // Insert a duplicate key on the to-side after the relationship exists.
    model
        .insert_row("Dim", vec![1.into(), "B".into()])
        .unwrap();

    let filter = FilterContext::empty().with_column_equals("Dim", "Attr", "B".into());
    assert_eq!(model.evaluate_measure("Total", &filter).unwrap(), 5.0.into());
}

#[test]
fn many_to_many_filter_propagates_from_dimension_to_fact() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![1.into(), "B".into()]).unwrap();
    dim.push_row(vec![2.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "Amount"]);
    fact.push_row(vec![100.into(), 1.into(), 10.0.into()]).unwrap();
    fact.push_row(vec![101.into(), 1.into(), 20.0.into()]).unwrap();
    fact.push_row(vec![102.into(), 2.into(), 5.0.into()]).unwrap();
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

    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();

    let filter = FilterContext::empty().with_column_equals("Dim", "Attr", "B".into());
    assert_eq!(model.evaluate_measure("Total", &filter).unwrap(), 30.0.into());
}

#[test]
fn many_to_many_bidirectional_propagation_preserves_duplicate_dimension_rows() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![1.into(), "B".into()]).unwrap();
    dim.push_row(vec![2.into(), "C".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key"]);
    fact.push_row(vec![100.into(), 1.into()]).unwrap();
    fact.push_row(vec![101.into(), 1.into()]).unwrap();
    fact.push_row(vec![102.into(), 2.into()]).unwrap();
    model.add_table(fact).unwrap();

    model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Key".into(),
            to_table: "Dim".into(),
            to_column: "Key".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Both,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_measure("Attr Count", "DISTINCTCOUNT(Dim[Attr])")
        .unwrap();

    let filter = FilterContext::empty().with_column_equals("Fact", "Key", 1.into());
    assert_eq!(
        model.evaluate_measure("Attr Count", &filter).unwrap(),
        2.into()
    );
}

#[test]
fn many_to_many_virtual_blank_row_includes_unmatched_facts() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Id", "Region"]);
    dim.push_row(vec![1.into(), "East".into()]).unwrap();
    dim.push_row(vec![1.into(), "West".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Group", "Amount"]);
    fact.push_row(vec![1.into(), "A".into(), 10.0.into()]).unwrap();
    fact.push_row(vec![999.into(), "A".into(), 7.0.into()]).unwrap();
    model.add_table(fact).unwrap();

    model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Id".into(),
            to_table: "Dim".into(),
            to_column: "Id".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();

    // When the dimension side is unfiltered, unmatched fact rows should still contribute.
    let fact_group = FilterContext::empty().with_column_equals("Fact", "Group", "A".into());
    assert_eq!(model.evaluate_measure("Total", &fact_group).unwrap(), 17.0.into());

    // Filtering the dimension to BLANK selects only unmatched facts.
    let blank_region = FilterContext::empty().with_column_equals("Dim", "Region", Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total", &blank_region).unwrap(),
        7.0.into()
    );
}

#[test]
fn related_errors_on_ambiguous_many_to_many_lookup() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![1.into(), "B".into()]).unwrap();
    dim.push_row(vec![2.into(), "C".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key"]);
    fact.push_row(vec![10.into(), 1.into()]).unwrap();
    fact.push_row(vec![11.into(), 2.into()]).unwrap();
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

    let engine = DaxEngine::new();

    let unique_ctx = {
        let mut ctx = RowContext::default();
        ctx.push("Fact", 1);
        ctx
    };
    assert_eq!(
        engine
            .evaluate(
                &model,
                "RELATED(Dim[Attr])",
                &FilterContext::empty(),
                &unique_ctx
            )
            .unwrap(),
        "C".into()
    );

    let ambiguous_ctx = {
        let mut ctx = RowContext::default();
        ctx.push("Fact", 0);
        ctx
    };
    let err = engine
        .evaluate(
            &model,
            "RELATED(Dim[Attr])",
            &FilterContext::empty(),
            &ambiguous_ctx,
        )
        .unwrap_err();
    match err {
        DaxError::Eval(msg) => assert!(
            msg.to_ascii_lowercase().contains("ambiguous"),
            "unexpected error message: {msg}"
        ),
        other => panic!("expected DaxError::Eval, got {other:?}"),
    }
}

#[test]
fn relatedtable_works_when_dimension_keys_are_not_unique() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![1.into(), "B".into()]).unwrap();
    dim.push_row(vec![2.into(), "C".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key"]);
    fact.push_row(vec![10.into(), 1.into()]).unwrap();
    fact.push_row(vec![11.into(), 1.into()]).unwrap();
    fact.push_row(vec![12.into(), 2.into()]).unwrap();
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

    let engine = DaxEngine::new();

    let mut ctx = RowContext::default();
    ctx.push("Dim", 0);
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(RELATEDTABLE(Fact))",
                &FilterContext::empty(),
                &ctx
            )
            .unwrap(),
        2.into()
    );
}

#[test]
fn many_to_many_crossfilter_can_enable_bidirectional_filtering_inside_calculate() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![1.into(), "B".into()]).unwrap();
    dim.push_row(vec![2.into(), "C".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key"]);
    fact.push_row(vec![10.into(), 1.into()]).unwrap();
    fact.push_row(vec![11.into(), 2.into()]).unwrap();
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

    let engine = DaxEngine::new();

    // Without CROSSFILTER(BOTH), filters on Fact do not flow to Dim (relationship is single-direction).
    let no_crossfilter = engine
        .evaluate(
            &model,
            "CALCULATE(DISTINCTCOUNT(Dim[Attr]), Fact[Key] = 1)",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(no_crossfilter, 3.into());

    // With CROSSFILTER(BOTH), Fact filters propagate to Dim; for M2M, all Dim rows for a visible key remain.
    let with_crossfilter = engine
        .evaluate(
            &model,
            "CALCULATE(DISTINCTCOUNT(Dim[Attr]), Fact[Key] = 1, CROSSFILTER(Fact[Key], Dim[Key], BOTH))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(with_crossfilter, 2.into());
}

#[test]
fn many_to_many_crossfilter_none_disables_relationship_propagation() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![1.into(), "B".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "Amount"]);
    fact.push_row(vec![10.into(), 1.into(), 5.0.into()]).unwrap();
    fact.push_row(vec![11.into(), 1.into(), 7.0.into()]).unwrap();
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

    let engine = DaxEngine::new();

    // With CROSSFILTER(NONE), Dim filters should not propagate to Fact.
    let value = engine
        .evaluate(
            &model,
            "CALCULATE(SUM(Fact[Amount]), Dim[Attr] = \"B\", CROSSFILTER(Fact[Key], Dim[Key], NONE))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(value, 12.0.into());
}

#[test]
fn many_to_many_userelationship_can_override_active_relationship() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["KeyA", "KeyB", "Attr"]);
    dim.push_row(vec![1.into(), 10.into(), "A".into()]).unwrap();
    dim.push_row(vec![1.into(), 11.into(), "B".into()]).unwrap();
    dim.push_row(vec![2.into(), 20.into(), "C".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "KeyA", "KeyB", "Amount"]);
    fact.push_row(vec![100.into(), 1.into(), 10.into(), 5.0.into()])
        .unwrap();
    fact.push_row(vec![101.into(), 1.into(), 11.into(), 7.0.into()])
        .unwrap();
    fact.push_row(vec![102.into(), 2.into(), 20.into(), 3.0.into()])
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

    // Inactive relationship on KeyB.
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

    let engine = DaxEngine::new();

    // With the active KeyA relationship, Dim[Attr] = "B" implies KeyA = 1, which includes both KeyA=1 fact rows.
    let active_rel = engine
        .evaluate(
            &model,
            "CALCULATE(SUM(Fact[Amount]), Dim[Attr] = \"B\")",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(active_rel, 12.0.into());

    // USERELATIONSHIP activates the inactive relationship and overrides the active one for the same table pair.
    let userel = engine
        .evaluate(
            &model,
            "CALCULATE(SUM(Fact[Amount]), Dim[Attr] = \"B\", USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
            &FilterContext::empty(),
            &RowContext::default(),
        )
        .unwrap();
    assert_eq!(userel, 7.0.into());
}
