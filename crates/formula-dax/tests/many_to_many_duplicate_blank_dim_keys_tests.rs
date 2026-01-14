use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions,
};
use formula_dax::{
    pivot, Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, GroupByColumn,
    PivotMeasure, Relationship, Table, Value,
};
use pretty_assertions::assert_eq;

fn options() -> TableOptions {
    TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    }
}

#[test]
fn many_to_many_duplicate_physical_blank_dim_keys_do_not_join() {
    // Regression guard: physical BLANK keys on the dimension side do not participate in
    // relationship joins, and fact-side BLANK/invalid foreign keys map to the *virtual* blank
    // member instead.
    //
    // This test also ensures duplicate physical BLANK dimension keys do not force many-to-many
    // expansion logic or cause duplicated groups.
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![Value::Blank, "Phys1".into()]).unwrap();
    dim.push_row(vec![Value::Blank, "Phys2".into()]).unwrap();
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Key", "Amount"]);
    // BLANK FK.
    fact.push_row(vec![Value::Blank, 10.0.into()]).unwrap();
    // Matched FK.
    fact.push_row(vec![1.into(), 20.0.into()]).unwrap();
    // Unmatched non-BLANK FK.
    fact.push_row(vec![999.into(), 30.0.into()]).unwrap();
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
        .add_measure("Total Amount", "SUM(Fact[Amount])")
        .unwrap();

    // Filtering to a physical BLANK key row's attribute should not match any fact rows with BLANK
    // foreign keys (or unmatched keys).
    let phys1 = FilterContext::empty().with_column_equals("Dim", "Attr", "Phys1".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &phys1).unwrap(),
        Value::Blank
    );

    let measures = vec![PivotMeasure::new("Total Amount", "[Total Amount]").unwrap()];
    let group_by = vec![GroupByColumn::new("Dim", "Attr")];

    // Under no filters, the pivot should include:
    // - the matched dimension attribute (A)
    // - a single BLANK group for the virtual blank member (includes BLANK + unmatched facts)
    // It must not include the physical BLANK-key dimension rows (Phys1/Phys2).
    let result = pivot(
        &model,
        "Fact",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("A"), 20.0.into()],
            vec![Value::Blank, 40.0.into()],
        ]
    );

    // Excluding BLANK explicitly should remove the virtual blank member and therefore exclude the
    // unmatched/BLANK fact rows.
    let non_blank_filter = DaxEngine::new()
        .apply_calculate_filters(&model, &FilterContext::empty(), &["Dim[Attr] <> BLANK()"])
        .unwrap();
    let result = pivot(&model, "Fact", &group_by, &measures, &non_blank_filter).unwrap();
    assert_eq!(result.rows, vec![vec![Value::from("A"), 20.0.into()]]);
}

#[test]
fn many_to_many_duplicate_physical_blank_dim_keys_do_not_join_for_columnar_dim() {
    // Same regression as `many_to_many_duplicate_physical_blank_dim_keys_do_not_join`, but with a
    // columnar-backed dimension table so the relationship uses `ToIndex::KeySet`.
    let mut model = DataModel::new();

    let schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Attr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut dim = ColumnarTableBuilder::new(schema, options());
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("Phys1".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("Phys2".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let mut fact = Table::new("Fact", vec!["Key", "Amount"]);
    fact.push_row(vec![Value::Blank, 10.0.into()]).unwrap();
    fact.push_row(vec![1.into(), 20.0.into()]).unwrap();
    fact.push_row(vec![999.into(), 30.0.into()]).unwrap();
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
        .add_measure("Total Amount", "SUM(Fact[Amount])")
        .unwrap();

    let phys1 = FilterContext::empty().with_column_equals("Dim", "Attr", "Phys1".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &phys1).unwrap(),
        Value::Blank
    );

    let measures = vec![PivotMeasure::new("Total Amount", "[Total Amount]").unwrap()];
    let group_by = vec![GroupByColumn::new("Dim", "Attr")];
    let result = pivot(
        &model,
        "Fact",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("A"), 20.0.into()],
            vec![Value::Blank, 40.0.into()]
        ]
    );

    let non_blank_filter = DaxEngine::new()
        .apply_calculate_filters(&model, &FilterContext::empty(), &["Dim[Attr] <> BLANK()"])
        .unwrap();
    let result = pivot(&model, "Fact", &group_by, &measures, &non_blank_filter).unwrap();
    assert_eq!(result.rows, vec![vec![Value::from("A"), 20.0.into()]]);
}

#[test]
fn many_to_many_duplicate_physical_blank_dim_keys_do_not_join_for_columnar_dim_and_fact() {
    // Same regression as `many_to_many_duplicate_physical_blank_dim_keys_do_not_join`, but with
    // both the dimension and fact tables stored columnar.
    let mut model = DataModel::new();

    let schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Attr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut dim = ColumnarTableBuilder::new(schema, options());
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("Phys1".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("Phys2".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

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
    let mut fact = ColumnarTableBuilder::new(schema, options());
    fact.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::Number(10.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(20.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(999.0),
        formula_columnar::Value::Number(30.0),
    ]);
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
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
        .add_measure("Total Amount", "SUM(Fact[Amount])")
        .unwrap();

    let phys1 = FilterContext::empty().with_column_equals("Dim", "Attr", "Phys1".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &phys1).unwrap(),
        Value::Blank
    );

    let measures = vec![PivotMeasure::new("Total Amount", "[Total Amount]").unwrap()];
    let group_by = vec![GroupByColumn::new("Dim", "Attr")];
    let result = pivot(
        &model,
        "Fact",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();
    assert_eq!(
        result.rows,
        vec![
            vec![Value::from("A"), 20.0.into()],
            vec![Value::Blank, 40.0.into()]
        ]
    );

    let non_blank_filter = DaxEngine::new()
        .apply_calculate_filters(&model, &FilterContext::empty(), &["Dim[Attr] <> BLANK()"])
        .unwrap();
    let result = pivot(&model, "Fact", &group_by, &measures, &non_blank_filter).unwrap();
    assert_eq!(result.rows, vec![vec![Value::from("A"), 20.0.into()]]);
}
