use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, Relationship,
    RowContext, Table,
};
use pretty_assertions::assert_eq;
use std::sync::Arc;

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

fn build_m2m_chain_model_columnar(cross_filter_direction: CrossFilterDirection) -> DataModel {
    let mut model = DataModel::new();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };

    let a_schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "AAttr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut a_builder = ColumnarTableBuilder::new(a_schema, options);
    a_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("a1")),
    ]);
    a_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("a2")),
    ]);
    model
        .add_table(Table::from_columnar("A", a_builder.finalize()))
        .unwrap();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };
    let b_schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "BAttr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut b_builder = ColumnarTableBuilder::new(b_schema, options);
    b_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("b1")),
    ]);
    b_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("b1b")),
    ]);
    b_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("b2")),
    ]);
    model
        .add_table(Table::from_columnar("B", b_builder.finalize()))
        .unwrap();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };
    let c_schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "CAttr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut c_builder = ColumnarTableBuilder::new(c_schema, options);
    c_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("c1")),
    ]);
    c_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("c2")),
    ]);
    c_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("c2b")),
    ]);
    model
        .add_table(Table::from_columnar("C", c_builder.finalize()))
        .unwrap();

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

fn build_m2m_chain_model_three_hops(cross_filter_direction: CrossFilterDirection) -> DataModel {
    let mut model = DataModel::new();

    let mut a = Table::new("A", vec!["Key", "AAttr"]);
    a.push_row(vec![1.into(), "a1".into()]).unwrap();
    a.push_row(vec![2.into(), "a2".into()]).unwrap();
    model.add_table(a).unwrap();

    let mut b = Table::new("B", vec!["Key", "BAttr"]);
    b.push_row(vec![1.into(), "b1".into()]).unwrap();
    b.push_row(vec![1.into(), "b1b".into()]).unwrap();
    b.push_row(vec![2.into(), "b2".into()]).unwrap();
    b.push_row(vec![2.into(), "b2b".into()]).unwrap();
    model.add_table(b).unwrap();

    let mut c = Table::new("C", vec!["Key", "CAttr"]);
    c.push_row(vec![1.into(), "c1".into()]).unwrap();
    c.push_row(vec![2.into(), "c2".into()]).unwrap();
    c.push_row(vec![2.into(), "c2b".into()]).unwrap();
    model.add_table(c).unwrap();

    let mut d = Table::new("D", vec!["Key", "DAttr"]);
    d.push_row(vec![1.into(), "d1".into()]).unwrap();
    d.push_row(vec![2.into(), "d2".into()]).unwrap();
    d.push_row(vec![2.into(), "d2b".into()]).unwrap();
    model.add_table(d).unwrap();

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
        .add_relationship(Relationship {
            name: "C_D".into(),
            from_table: "C".into(),
            from_column: "Key".into(),
            to_table: "D".into(),
            to_column: "Key".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
}

fn build_m2m_chain_model_three_hops_columnar(
    cross_filter_direction: CrossFilterDirection,
) -> DataModel {
    let mut model = DataModel::new();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };
    let a_schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "AAttr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut a_builder = ColumnarTableBuilder::new(a_schema, options);
    a_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("a1")),
    ]);
    a_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("a2")),
    ]);
    model
        .add_table(Table::from_columnar("A", a_builder.finalize()))
        .unwrap();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };
    let b_schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "BAttr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut b_builder = ColumnarTableBuilder::new(b_schema, options);
    b_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("b1")),
    ]);
    b_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("b1b")),
    ]);
    b_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("b2")),
    ]);
    b_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("b2b")),
    ]);
    model
        .add_table(Table::from_columnar("B", b_builder.finalize()))
        .unwrap();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };
    let c_schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "CAttr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut c_builder = ColumnarTableBuilder::new(c_schema, options);
    c_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("c1")),
    ]);
    c_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("c2")),
    ]);
    c_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("c2b")),
    ]);
    model
        .add_table(Table::from_columnar("C", c_builder.finalize()))
        .unwrap();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };
    let d_schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "DAttr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut d_builder = ColumnarTableBuilder::new(d_schema, options);
    d_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("d1")),
    ]);
    d_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("d2")),
    ]);
    d_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("d2b")),
    ]);
    model
        .add_table(Table::from_columnar("D", d_builder.finalize()))
        .unwrap();

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
        .add_relationship(Relationship {
            name: "C_D".into(),
            from_table: "C".into(),
            from_column: "Key".into(),
            to_table: "D".into(),
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

#[test]
fn single_direction_chain_propagation_columnar() {
    let model = build_m2m_chain_model_columnar(CrossFilterDirection::Single);
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
fn bidirectional_chain_propagation_columnar() {
    let model = build_m2m_chain_model_columnar(CrossFilterDirection::Both);
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
fn bidirectional_middle_table_filter_propagates_both_ways_columnar() {
    let model = build_m2m_chain_model_columnar(CrossFilterDirection::Both);
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

#[test]
fn three_hop_single_direction_chain_propagates_to_a() {
    let model = build_m2m_chain_model_three_hops(CrossFilterDirection::Single);
    let engine = DaxEngine::new();

    // This requires more than one propagation pass because the relationship insertion order is
    // A_B, B_C, C_D: filtering D should restrict C first, then B, then finally A.
    let filter = FilterContext::empty().with_column_equals("D", "DAttr", "d2".into());

    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(C)", &filter, &RowContext::default(),)
            .unwrap(),
        2.into()
    );
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(B)", &filter, &RowContext::default(),)
            .unwrap(),
        2.into()
    );
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(A)", &filter, &RowContext::default(),)
            .unwrap(),
        1.into()
    );
}

#[test]
fn three_hop_single_direction_chain_propagates_to_a_columnar() {
    let model = build_m2m_chain_model_three_hops_columnar(CrossFilterDirection::Single);
    let engine = DaxEngine::new();

    let filter = FilterContext::empty().with_column_equals("D", "DAttr", "d2".into());

    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(C)", &filter, &RowContext::default(),)
            .unwrap(),
        2.into()
    );
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(B)", &filter, &RowContext::default(),)
            .unwrap(),
        2.into()
    );
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(A)", &filter, &RowContext::default(),)
            .unwrap(),
        1.into()
    );
}

#[test]
fn single_direction_does_not_propagate_from_a_to_c() {
    let model = build_m2m_chain_model(CrossFilterDirection::Single);
    let engine = DaxEngine::new();

    let filter = FilterContext::empty().with_column_equals("A", "AAttr", "a1".into());

    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(B)", &filter, &RowContext::default(),)
            .unwrap(),
        3.into()
    );
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(C)", &filter, &RowContext::default(),)
            .unwrap(),
        3.into()
    );
}

#[test]
fn single_direction_does_not_propagate_from_a_to_c_columnar() {
    let model = build_m2m_chain_model_columnar(CrossFilterDirection::Single);
    let engine = DaxEngine::new();

    let filter = FilterContext::empty().with_column_equals("A", "AAttr", "a1".into());

    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(B)", &filter, &RowContext::default(),)
            .unwrap(),
        3.into()
    );
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(C)", &filter, &RowContext::default(),)
            .unwrap(),
        3.into()
    );
}

#[test]
fn single_direction_middle_table_filter_propagates_only_to_a() {
    let model = build_m2m_chain_model(CrossFilterDirection::Single);
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
        3.into()
    );
}

#[test]
fn single_direction_middle_table_filter_propagates_only_to_a_columnar() {
    let model = build_m2m_chain_model_columnar(CrossFilterDirection::Single);
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
        3.into()
    );
}

#[test]
fn bidirectional_conflicting_filters_converge_to_empty() {
    let model = build_m2m_chain_model(CrossFilterDirection::Both);
    let engine = DaxEngine::new();

    let filter = FilterContext::empty()
        .with_column_equals("A", "AAttr", "a1".into())
        .with_column_equals("C", "CAttr", "c2".into());

    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(B)", &filter, &RowContext::default(),)
            .unwrap(),
        0.into()
    );
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(A)", &filter, &RowContext::default(),)
            .unwrap(),
        0.into()
    );
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(C)", &filter, &RowContext::default(),)
            .unwrap(),
        0.into()
    );
}

#[test]
fn bidirectional_conflicting_filters_converge_to_empty_columnar() {
    let model = build_m2m_chain_model_columnar(CrossFilterDirection::Both);
    let engine = DaxEngine::new();

    let filter = FilterContext::empty()
        .with_column_equals("A", "AAttr", "a1".into())
        .with_column_equals("C", "CAttr", "c2".into());

    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(B)", &filter, &RowContext::default(),)
            .unwrap(),
        0.into()
    );
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(A)", &filter, &RowContext::default(),)
            .unwrap(),
        0.into()
    );
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(C)", &filter, &RowContext::default(),)
            .unwrap(),
        0.into()
    );
}
