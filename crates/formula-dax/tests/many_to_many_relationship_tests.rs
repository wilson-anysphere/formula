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

