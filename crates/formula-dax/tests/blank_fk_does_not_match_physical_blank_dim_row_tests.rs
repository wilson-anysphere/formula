use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, Relationship, RowContext,
    Table, Value,
};
use pretty_assertions::assert_eq;

#[test]
fn blank_fk_does_not_match_physical_blank_dim_row() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    // A physical BLANK key row exists on the dimension side.
    dim.push_row(vec![Value::Blank, "Phys".into()]).unwrap();
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Key", "Amount"]);
    // Fact contains a BLANK FK row.
    fact.push_row(vec![Value::Blank, 10.into()]).unwrap();
    // Include at least one non-BLANK row so filtering to BLANK actually restricts the fact table.
    fact.push_row(vec![1.into(), 20.into()]).unwrap();
    model.add_table(fact).unwrap();

    model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Key".into(),
            to_table: "Dim".into(),
            to_column: "Key".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Both,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    let engine = DaxEngine::new();
    let filter = FilterContext::empty().with_column_equals("Fact", "Key", Value::Blank);

    // When the fact FK is BLANK, it should map to the virtual blank/unknown member on the
    // dimension side. A physical Dim[Key] = BLANK() row must *not* become visible via relationship
    // propagation.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(VALUES(Dim[Attr]))",
                &filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "SELECTEDVALUE(Dim[Attr])",
                &filter,
                &RowContext::default(),
            )
            .unwrap(),
        Value::Blank
    );
}

