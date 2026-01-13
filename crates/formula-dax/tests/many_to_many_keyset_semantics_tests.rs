use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, Relationship, RowContext,
    Table, Value,
};
use pretty_assertions::assert_eq;

fn build_m2m_model(cross_filter_direction: CrossFilterDirection) -> DataModel {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![1.into(), "B".into()]).unwrap();
    dim.push_row(vec![2.into(), "C".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "Amount"]);
    fact.push_row(vec![10.into(), 1.into(), 5.0.into()]).unwrap();
    fact.push_row(vec![11.into(), 2.into(), 7.0.into()]).unwrap();
    model.add_table(fact).unwrap();

    model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Key".into(),
            to_table: "Dim".into(),
            to_column: "Key".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
}

#[test]
fn single_direction_propagation_uses_key_set_with_duplicate_dim_keys() {
    let model = build_m2m_model(CrossFilterDirection::Single);
    let engine = DaxEngine::new();

    // Keep only Dim[Key] = 1 via an attribute filter that matches a *single row*.
    // M2M propagation should still use the distinct set of visible key values (here: {1}),
    // not the identity of the surviving Dim row.
    let filter = FilterContext::empty().with_column_equals("Dim", "Attr", "A".into());

    assert_eq!(
        engine
            .evaluate(
                &model,
                "SUM(Fact[Amount])",
                &filter,
                &RowContext::default()
            )
            .unwrap(),
        5.0.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(Fact)",
                &filter,
                &RowContext::default()
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn bidirectional_propagation_exposes_all_dim_rows_for_a_key() {
    let model = build_m2m_model(CrossFilterDirection::Both);
    let engine = DaxEngine::new();

    // Filter from-side down to key 1, and ensure bidirectional propagation makes *all*
    // Dim rows with Key=1 visible (Attr A and B).
    let filter = FilterContext::empty().with_column_equals("Fact", "Key", 1.into());

    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(VALUES(Dim[Attr]))",
                &filter,
                &RowContext::default()
            )
            .unwrap(),
        2.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "DISTINCTCOUNT(Dim[Attr])",
                &filter,
                &RowContext::default()
            )
            .unwrap(),
        2.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "SELECTEDVALUE(Dim[Attr], \"Multiple\")",
                &filter,
                &RowContext::default()
            )
            .unwrap(),
        Value::from("Multiple")
    );
}

#[test]
fn filtering_duplicate_dim_rows_down_to_one_key_still_has_one_value() {
    let model = build_m2m_model(CrossFilterDirection::Single);
    let engine = DaxEngine::new();

    // Dim[Key]=1 matches two rows, but there is still a single distinct key value.
    let filter = FilterContext::empty().with_column_equals("Dim", "Key", 1.into());
    assert_eq!(
        engine
            .evaluate(
                &model,
                "HASONEVALUE(Dim[Key])",
                &filter,
                &RowContext::default()
            )
            .unwrap(),
        true.into()
    );
}

