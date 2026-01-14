use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
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

#[test]
fn blank_fk_does_not_match_physical_blank_dim_row_with_crossfilter_override() {
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

    // Single-direction relationship; enable bidirectional filtering via CROSSFILTER in the query.
    model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Key".into(),
            to_table: "Dim".into(),
            to_column: "Key".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    let engine = DaxEngine::new();
    let filter = FilterContext::empty().with_column_equals("Fact", "Key", Value::Blank);

    // With bidirectional filtering forced via CROSSFILTER(BOTH), a fact-side BLANK FK must not
    // make a physical Dim[Key] = BLANK() row visible; only the virtual blank/unknown member should
    // contribute.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(VALUES(Dim[Attr])), CROSSFILTER(Fact[Key], Dim[Key], \"BOTH\"))",
                &filter,
                &RowContext::default(),
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn blank_fk_does_not_match_physical_blank_dim_row_for_columnar_fact() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    // A physical BLANK key row exists on the dimension side.
    dim.push_row(vec![Value::Blank, "Phys".into()]).unwrap();
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    // Columnar fact table (ensures `relationship.from_index == None`, exercising the columnar
    // `Direction::ToOne` propagation path).
    let schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };
    let mut fact = ColumnarTableBuilder::new(schema, options);
    // Two BLANK FK rows. This ensures the allowed fact row set is "dense" relative to the
    // `row_count/64` heuristic in `propagate_filter(Direction::ToOne)` so we exercise the
    // BitVec-scanning code path (no large `Vec<usize>` allocation).
    fact.append_row(&[formula_columnar::Value::Null, formula_columnar::Value::Number(10.0)]);
    fact.append_row(&[formula_columnar::Value::Null, formula_columnar::Value::Number(11.0)]);
    // Many non-BLANK rows so filtering to BLANK actually restricts the fact table.
    for _ in 0..62 {
        fact.append_row(&[
            formula_columnar::Value::Number(1.0),
            formula_columnar::Value::Number(20.0),
        ]);
    }
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

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

#[test]
fn blank_fk_does_not_match_physical_blank_dim_row_for_columnar_fact_sparse_path() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    // A physical BLANK key row exists on the dimension side.
    dim.push_row(vec![Value::Blank, "Phys".into()]).unwrap();
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    // Columnar fact table sized so `row_count/64 == 1` and `visible_count == 1`, which forces the
    // sparse path that materializes `Vec<usize>` for `distinct_values_filtered`.
    let schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };
    let mut fact = ColumnarTableBuilder::new(schema, options);
    // One BLANK FK row.
    fact.append_row(&[formula_columnar::Value::Null, formula_columnar::Value::Number(10.0)]);
    // Many non-BLANK rows so filtering to BLANK actually restricts the fact table.
    for _ in 0..63 {
        fact.append_row(&[
            formula_columnar::Value::Number(1.0),
            formula_columnar::Value::Number(20.0),
        ]);
    }
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

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

#[test]
fn blank_fk_does_not_match_physical_blank_dim_row_for_columnar_dim() {
    let mut model = DataModel::new();

    // Columnar dimension table (ensures `relationship.to_index` is built from columnar storage).
    let dim_schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Attr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, options);
    // A physical BLANK key row exists on the dimension side.
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("Phys".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

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

#[test]
fn blank_fk_does_not_match_physical_blank_dim_row_for_columnar_dim_and_fact() {
    let mut model = DataModel::new();

    let dim_schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Attr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, options);
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("Phys".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, options);
    fact.append_row(&[formula_columnar::Value::Null, formula_columnar::Value::Number(10.0)]);
    fact.append_row(&[formula_columnar::Value::Number(1.0), formula_columnar::Value::Number(20.0)]);
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

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
