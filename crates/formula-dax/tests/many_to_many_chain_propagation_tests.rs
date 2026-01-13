use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, Relationship,
    RowContext, Table,
};
use pretty_assertions::assert_eq;

fn build_m2m_chain_model(cross_filter_direction: CrossFilterDirection) -> DataModel {
    let mut model = DataModel::new();

    let mut a = Table::new("A", vec!["Key", "AAttr"]);
    a.push_row(vec![1.into(), "a1".into()]).unwrap();
    a.push_row(vec![2.into(), "a2".into()]).unwrap();
    model.add_table(a).unwrap();

    let mut b = Table::new("B", vec!["Key", "BAttr"]);
    b.push_row(vec![1.into(), "b1".into()]).unwrap();
    b.push_row(vec![1.into(), "b1b".into()]).unwrap();
    b.push_row(vec![2.into(), "b2".into()]).unwrap();
    model.add_table(b).unwrap();

    let mut c = Table::new("C", vec!["Key", "CAttr"]);
    c.push_row(vec![1.into(), "c1".into()]).unwrap();
    c.push_row(vec![2.into(), "c2".into()]).unwrap();
    c.push_row(vec![2.into(), "c2b".into()]).unwrap();
    model.add_table(c).unwrap();

    model
        .add_relationship(Relationship {
            name: "A_B".into(),
            from_table: "A".into(),
            from_column: "Key".into(),
            to_table: "B".into(),
            to_column: "Key".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_relationship(Relationship {
            name: "B_C".into(),
            from_table: "B".into(),
            from_column: "Key".into(),
            to_table: "C".into(),
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
fn single_direction_chain_propagation() {
    let model = build_m2m_chain_model(CrossFilterDirection::Single);
    let engine = DaxEngine::new();

    let filter = FilterContext::empty().with_column_equals("C", "CAttr", "c2".into());

    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(B)", &filter, &RowContext::default(),)
            .unwrap(),
        1.into()
    );
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(A)", &filter, &RowContext::default(),)
            .unwrap(),
        1.into()
    );
}

#[test]
fn bidirectional_chain_propagation() {
    let model = build_m2m_chain_model(CrossFilterDirection::Both);
    let engine = DaxEngine::new();

    let filter = FilterContext::empty().with_column_equals("A", "AAttr", "a1".into());

    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(B)", &filter, &RowContext::default(),)
            .unwrap(),
        2.into()
    );
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(C)", &filter, &RowContext::default(),)
            .unwrap(),
        1.into()
    );
}

#[test]
fn bidirectional_middle_table_filter_propagates_both_ways() {
    let model = build_m2m_chain_model(CrossFilterDirection::Both);
    let engine = DaxEngine::new();

    let filter = FilterContext::empty().with_column_equals("B", "BAttr", "b1b".into());

    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(A)", &filter, &RowContext::default(),)
            .unwrap(),
        1.into()
    );
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(C)", &filter, &RowContext::default(),)
            .unwrap(),
        1.into()
    );
}
