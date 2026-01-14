use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, Relationship, RowContext,
    Table, Value,
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
fn userelationship_override_works_with_m2m_for_columnar_dim() {
    let mut model = DataModel::new();

    let dim_schema = vec![
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Attr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::String("A".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(20.0),
        formula_columnar::Value::String("B".into()),
    ]);
    // Physical BLANK key row: should not pick up unmatched facts under USERELATIONSHIP.
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::Null,
        formula_columnar::Value::String("PhysicalBlank".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "KeyA", "KeyB", "Amount"]);
    // Cross keys so the active vs. USERELATIONSHIP-overridden relationship produces different totals.
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

    let filter_b = FilterContext::empty().with_column_equals("Dim", "Attr", "B".into());
    assert_eq!(model.evaluate_measure("Total", &filter_b).unwrap(), 200.0.into());
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &filter_b).unwrap(),
        100.0.into()
    );
}

#[test]
fn userelationship_override_works_with_m2m_for_columnar_dim_and_fact() {
    // Same regression as `userelationship_override_works_with_m2m`, but with both tables
    // columnar-backed.
    let mut model = DataModel::new();

    let dim_schema = vec![
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Attr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::String("A".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(20.0),
        formula_columnar::Value::String("B".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let fact_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, fact_options);
    // Cross keys so the active vs. USERELATIONSHIP-overridden relationship produces different totals.
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(20.0),
        formula_columnar::Value::Number(100.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(200.0),
    ]);
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

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

    let filter_b = FilterContext::empty().with_column_equals("Dim", "Attr", "B".into());
    assert_eq!(model.evaluate_measure("Total", &filter_b).unwrap(), 200.0.into());
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &filter_b).unwrap(),
        100.0.into()
    );
}

#[test]
fn userelationship_override_sees_virtual_blank_member_for_columnar_dim() {
    // Regression coverage: even when the *dimension* table is columnar-backed, an inactive
    // ManyToMany relationship with RI disabled should still surface unmatched fact keys via the
    // relationship-generated blank/unknown member under USERELATIONSHIP.
    let mut model = DataModel::new();

    let dim_schema = vec![
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Attr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::String("A".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(20.0),
        formula_columnar::Value::String("B".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "KeyA", "KeyB", "Amount"]);
    // One row matched on KeyB.
    fact.push_row(vec![1.into(), 1.into(), 10.into(), 5.0.into()])
        .unwrap();
    // One row unmatched on KeyB (should belong to blank member for relationship B).
    fact.push_row(vec![2.into(), 1.into(), 999.into(), 7.0.into()])
        .unwrap();
    model.add_table(fact).unwrap();

    // Active relationship A (RI enforced).
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

    // Inactive relationship B (RI disabled) so unmatched keys map to the virtual blank member.
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
    model
        .add_measure(
            "Total via KeyB",
            "CALCULATE([Total], USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
        )
        .unwrap();

    let blank_attr = FilterContext::empty().with_column_equals("Dim", "Attr", Value::Blank);
    assert_eq!(model.evaluate_measure("Total", &blank_attr).unwrap(), Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &blank_attr).unwrap(),
        7.0.into()
    );

    // Ensure the physical BLANK-key dimension row does *not* behave like the relationship-generated
    // blank member.
    let physical_blank =
        FilterContext::empty().with_column_equals("Dim", "Attr", "PhysicalBlank".into());
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &physical_blank).unwrap(),
        Value::Blank
    );
}

#[test]
fn userelationship_override_sees_virtual_blank_member_for_columnar_dim_and_fact() {
    // Same regression as `userelationship_override_sees_virtual_blank_member_for_columnar_dim`, but
    // with both tables columnar-backed.
    let mut model = DataModel::new();

    let dim_schema = vec![
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Attr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let dim_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, dim_options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::String("A".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(20.0),
        formula_columnar::Value::String("B".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::Null,
        formula_columnar::Value::String("PhysicalBlank".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let fact_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, fact_options);
    // One row matched on KeyB.
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(5.0),
    ]);
    // One row unmatched on KeyB (should belong to blank member for relationship B).
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(999.0),
        formula_columnar::Value::Number(7.0),
    ]);
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

    // Active relationship A (RI enforced).
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

    // Inactive relationship B (RI disabled) so unmatched keys map to the virtual blank member.
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
    model
        .add_measure(
            "Total via KeyB",
            "CALCULATE([Total], USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
        )
        .unwrap();

    let blank_attr = FilterContext::empty().with_column_equals("Dim", "Attr", Value::Blank);
    assert_eq!(model.evaluate_measure("Total", &blank_attr).unwrap(), Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &blank_attr).unwrap(),
        7.0.into()
    );

    let physical_blank =
        FilterContext::empty().with_column_equals("Dim", "Attr", "PhysicalBlank".into());
    assert_eq!(model.evaluate_measure("Total", &physical_blank).unwrap(), Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &physical_blank).unwrap(),
        Value::Blank
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
fn blank_foreign_keys_do_not_match_physical_blank_dimension_keys() {
    // Regression: Tabular's relationship-generated blank member is distinct from a *physical*
    // BLANK key on the dimension side. Fact rows with BLANK foreign keys should belong to the
    // virtual blank member, not match a physical Dim row whose key is BLANK.
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    // Physical dimension row whose key is BLANK.
    dim.push_row(vec![Value::Blank, "PhysicalBlank".into()])
        .unwrap();
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

    // The virtual blank member should include fact rows whose FK is BLANK.
    let blank_attr = FilterContext::empty().with_column_equals("Dim", "Attr", Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total Amount", &blank_attr).unwrap(),
        7.0.into()
    );

    // Filtering to the *physical* BLANK key row should NOT match those fact rows.
    let physical_blank = FilterContext::empty()
        .with_column_equals("Dim", "Attr", "PhysicalBlank".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &physical_blank).unwrap(),
        Value::Blank
    );
}

#[test]
fn blank_foreign_keys_do_not_match_physical_blank_dimension_keys_for_columnar_dim() {
    // Same regression as `blank_foreign_keys_do_not_match_physical_blank_dimension_keys`, but with
    // a columnar-backed dimension table.
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
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    // Physical dimension row whose key is BLANK.
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("PhysicalBlank".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

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

    // The virtual blank member should include fact rows whose FK is BLANK.
    let blank_attr = FilterContext::empty().with_column_equals("Dim", "Attr", Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total Amount", &blank_attr).unwrap(),
        7.0.into()
    );

    // Filtering to the *physical* BLANK key row should NOT match those fact rows.
    let physical_blank =
        FilterContext::empty().with_column_equals("Dim", "Attr", "PhysicalBlank".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &physical_blank).unwrap(),
        Value::Blank
    );
}

#[test]
fn blank_foreign_keys_do_not_match_physical_blank_dimension_keys_for_columnar_fact() {
    // Same regression as `blank_foreign_keys_do_not_match_physical_blank_dimension_keys`, but with
    // a columnar-backed fact table.
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    // Physical dimension row whose key is BLANK.
    dim.push_row(vec![Value::Blank, "PhysicalBlank".into()])
        .unwrap();
    model.add_table(dim).unwrap();

    let schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
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
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(schema, options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Null,
        formula_columnar::Value::Number(7.0),
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
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Total Amount", "SUM(Fact[Amount])").unwrap();

    // The virtual blank member should include fact rows whose FK is BLANK.
    let blank_attr = FilterContext::empty().with_column_equals("Dim", "Attr", Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total Amount", &blank_attr).unwrap(),
        7.0.into()
    );

    // Filtering to the *physical* BLANK key row should NOT match those fact rows.
    let physical_blank =
        FilterContext::empty().with_column_equals("Dim", "Attr", "PhysicalBlank".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &physical_blank).unwrap(),
        Value::Blank
    );
}

#[test]
fn blank_foreign_keys_do_not_match_physical_blank_dimension_keys_for_columnar_dim_and_fact() {
    // Same regression as `blank_foreign_keys_do_not_match_physical_blank_dimension_keys`, but with
    // *both* dimension and fact tables columnar-backed.
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
    let dim_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, dim_options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("PhysicalBlank".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let fact_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, fact_options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Null,
        formula_columnar::Value::Number(7.0),
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
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Total Amount", "SUM(Fact[Amount])").unwrap();

    let blank_attr = FilterContext::empty().with_column_equals("Dim", "Attr", Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total Amount", &blank_attr).unwrap(),
        7.0.into()
    );

    let physical_blank =
        FilterContext::empty().with_column_equals("Dim", "Attr", "PhysicalBlank".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &physical_blank).unwrap(),
        Value::Blank
    );
}

#[test]
fn related_does_not_match_blank_facts_to_physical_blank_dimension_keys() {
    // Regression: BLANK fact keys map to the *virtual* blank dimension member, not a physical
    // dimension row whose key is BLANK. This should apply to `RELATED` navigation as well.
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![Value::Blank, "PhysicalBlank".into()])
        .unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key"]);
    fact.push_row(vec![1.into(), 1.into()]).unwrap();
    fact.push_row(vec![2.into(), Value::Blank]).unwrap();
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

    model
        .add_calculated_column("Fact", "Related Attr", "RELATED(Dim[Attr])")
        .unwrap();

    let fact = model.table("Fact").unwrap();
    assert_eq!(fact.value(0, "Related Attr").unwrap(), "A".into());
    assert_eq!(fact.value(1, "Related Attr").unwrap(), Value::Blank);
}

#[test]
fn related_does_not_match_blank_facts_to_physical_blank_dimension_keys_for_columnar_dim() {
    // Same regression as `related_does_not_match_blank_facts_to_physical_blank_dimension_keys`, but
    // with a columnar-backed dimension table.
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
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("PhysicalBlank".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key"]);
    fact.push_row(vec![1.into(), 1.into()]).unwrap();
    fact.push_row(vec![2.into(), Value::Blank]).unwrap();
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

    model
        .add_calculated_column("Fact", "Related Attr", "RELATED(Dim[Attr])")
        .unwrap();

    let fact = model.table("Fact").unwrap();
    assert_eq!(fact.value(0, "Related Attr").unwrap(), "A".into());
    assert_eq!(fact.value(1, "Related Attr").unwrap(), Value::Blank);
}

#[test]
fn related_does_not_match_blank_facts_to_physical_blank_dimension_keys_for_columnar_fact() {
    // Same regression as `related_does_not_match_blank_facts_to_physical_blank_dimension_keys`, but
    // with a columnar-backed fact table.
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![Value::Blank, "PhysicalBlank".into()])
        .unwrap();
    model.add_table(dim).unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Null,
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
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_calculated_column("Fact", "Related Attr", "RELATED(Dim[Attr])")
        .unwrap();

    let fact = model.table("Fact").unwrap();
    assert_eq!(fact.value(0, "Related Attr").unwrap(), "A".into());
    assert_eq!(fact.value(1, "Related Attr").unwrap(), Value::Blank);
}

#[test]
fn related_does_not_match_blank_facts_to_physical_blank_dimension_keys_for_columnar_dim_and_fact() {
    // Same regression as `related_does_not_match_blank_facts_to_physical_blank_dimension_keys`, but
    // with both dimension and fact tables columnar-backed.
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
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("PhysicalBlank".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Null,
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
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_calculated_column("Fact", "Related Attr", "RELATED(Dim[Attr])")
        .unwrap();

    let fact = model.table("Fact").unwrap();
    assert_eq!(fact.value(0, "Related Attr").unwrap(), "A".into());
    assert_eq!(fact.value(1, "Related Attr").unwrap(), Value::Blank);
}

#[test]
fn blank_foreign_keys_do_not_match_physical_blank_dimension_keys_with_bidirectional_filtering() {
    // Similar to `blank_foreign_keys_do_not_match_physical_blank_dimension_keys`, but with
    // bidirectional filtering enabled.
    //
    // This exercises `propagate_filter(Direction::ToOne)`: filtering the fact table to BLANK()
    // should not make a *physical* BLANK key row on the dimension side visible.
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    // Physical dimension row whose key is BLANK.
    dim.push_row(vec![Value::Blank, "PhysicalBlank".into()])
        .unwrap();
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
            cross_filter_direction: CrossFilterDirection::Both,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    let engine = DaxEngine::new();

    let filter = FilterContext::empty().with_column_equals("Fact", "Key", Value::Blank);
    let result = engine
        .evaluate(
            &model,
            r#"COUNTROWS(FILTER(VALUES(Dim[Attr]), Dim[Attr] = "PhysicalBlank"))"#,
            &filter,
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(result, 0.into());
}

#[test]
fn blank_foreign_keys_do_not_match_physical_blank_dimension_keys_with_bidirectional_filtering_for_columnar_dim(
) {
    // Same regression as `blank_foreign_keys_do_not_match_physical_blank_dimension_keys_with_bidirectional_filtering`,
    // but with a columnar-backed dimension table.
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
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    // Physical dimension row whose key is BLANK.
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("PhysicalBlank".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

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
            cross_filter_direction: CrossFilterDirection::Both,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    let engine = DaxEngine::new();

    let filter = FilterContext::empty().with_column_equals("Fact", "Key", Value::Blank);
    let result = engine
        .evaluate(
            &model,
            r#"COUNTROWS(FILTER(VALUES(Dim[Attr]), Dim[Attr] = "PhysicalBlank"))"#,
            &filter,
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(result, 0.into());
}

#[test]
fn blank_foreign_keys_do_not_match_physical_blank_dimension_keys_with_bidirectional_filtering_for_columnar_fact(
) {
    // Same regression as `blank_foreign_keys_do_not_match_physical_blank_dimension_keys_with_bidirectional_filtering`,
    // but with a columnar-backed fact table.
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    // Physical dimension row whose key is BLANK.
    dim.push_row(vec![Value::Blank, "PhysicalBlank".into()])
        .unwrap();
    model.add_table(dim).unwrap();

    let schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
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
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(schema, options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Null,
        formula_columnar::Value::Number(7.0),
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
            cross_filter_direction: CrossFilterDirection::Both,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    let engine = DaxEngine::new();

    let filter = FilterContext::empty().with_column_equals("Fact", "Key", Value::Blank);
    let result = engine
        .evaluate(
            &model,
            r#"COUNTROWS(FILTER(VALUES(Dim[Attr]), Dim[Attr] = "PhysicalBlank"))"#,
            &filter,
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(result, 0.into());
}

#[test]
fn blank_foreign_keys_do_not_match_physical_blank_dimension_keys_with_bidirectional_filtering_for_columnar_dim_and_fact(
) {
    // Same regression as `blank_foreign_keys_do_not_match_physical_blank_dimension_keys_with_bidirectional_filtering`,
    // but with *both* dimension and fact tables columnar-backed.
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
    let dim_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, dim_options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("PhysicalBlank".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let fact_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, fact_options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Null,
        formula_columnar::Value::Number(7.0),
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
            cross_filter_direction: CrossFilterDirection::Both,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    let engine = DaxEngine::new();

    let filter = FilterContext::empty().with_column_equals("Fact", "Key", Value::Blank);
    let result = engine
        .evaluate(
            &model,
            r#"COUNTROWS(FILTER(VALUES(Dim[Attr]), Dim[Attr] = "PhysicalBlank"))"#,
            &filter,
            &RowContext::default(),
        )
        .unwrap();

    assert_eq!(result, 0.into());
}

#[test]
fn relatedtable_from_virtual_blank_dimension_member_includes_unmatched_facts_m2m() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "Amount"]);
    fact.push_row(vec![1.into(), 1.into(), 10.0.into()]).unwrap();
    fact.push_row(vec![2.into(), 999.into(), 7.0.into()]).unwrap();
    fact.push_row(vec![3.into(), Value::Blank, 5.0.into()])
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    let engine = DaxEngine::new();
    let blank_row = model.table("Dim").unwrap().row_count();

    let mut ctx = RowContext::default();
    ctx.push("Dim", blank_row);

    // The "virtual blank" member should expose fact rows whose FK is BLANK or has no match.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "SUMX(RELATEDTABLE(Fact), Fact[Amount])",
                &FilterContext::empty(),
                &ctx,
            )
            .unwrap(),
        12.0.into()
    );
}

#[test]
fn relatedtable_does_not_match_blank_facts_to_physical_blank_dimension_keys() {
    // Same regression as `blank_foreign_keys_do_not_match_physical_blank_dimension_keys`, but for
    // row-context navigation via RELATEDTABLE.
    //
    // Facts with BLANK foreign keys belong to the relationship-generated blank member, not a
    // physical Dim row whose key is BLANK.
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![Value::Blank, "PhysicalBlank".into()])
        .unwrap();
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

    model
        .add_calculated_column("Dim", "Related Fact Count", "COUNTROWS(RELATEDTABLE(Fact))")
        .unwrap();

    let dim = model.table("Dim").unwrap();
    assert_eq!(dim.value(0, "Related Fact Count").unwrap(), 1.into());
    assert_eq!(dim.value(1, "Related Fact Count").unwrap(), 0.into());
}

#[test]
fn relatedtable_does_not_match_blank_facts_to_physical_blank_dimension_keys_for_columnar_dim() {
    // Same regression as `relatedtable_does_not_match_blank_facts_to_physical_blank_dimension_keys`,
    // but with a columnar-backed dimension table.
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
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("PhysicalBlank".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key"]);
    fact.push_row(vec![1.into(), 1.into()]).unwrap();
    fact.push_row(vec![2.into(), Value::Blank]).unwrap();
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

    model
        .add_calculated_column("Dim", "Related Fact Count", "COUNTROWS(RELATEDTABLE(Fact))")
        .unwrap();

    let dim = model.table("Dim").unwrap();
    assert_eq!(dim.value(0, "Related Fact Count").unwrap(), 1.into());
    assert_eq!(dim.value(1, "Related Fact Count").unwrap(), 0.into());
}

#[test]
fn relatedtable_does_not_match_blank_facts_to_physical_blank_dimension_keys_for_columnar_fact() {
    // Same regression as `relatedtable_does_not_match_blank_facts_to_physical_blank_dimension_keys`,
    // but with a columnar-backed fact table.
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![Value::Blank, "PhysicalBlank".into()])
        .unwrap();
    model.add_table(dim).unwrap();

    let schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(schema, options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Null,
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
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_calculated_column("Dim", "Related Fact Count", "COUNTROWS(RELATEDTABLE(Fact))")
        .unwrap();

    let dim = model.table("Dim").unwrap();
    assert_eq!(dim.value(0, "Related Fact Count").unwrap(), 1.into());
    assert_eq!(dim.value(1, "Related Fact Count").unwrap(), 0.into());
}

#[test]
fn relatedtable_does_not_match_blank_facts_to_physical_blank_dimension_keys_for_columnar_dim_and_fact(
) {
    // Same regression as `relatedtable_does_not_match_blank_facts_to_physical_blank_dimension_keys`,
    // but with *both* dimension and fact tables columnar-backed.
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
    let dim_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, dim_options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("PhysicalBlank".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let fact_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, fact_options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Null,
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
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_calculated_column("Dim", "Related Fact Count", "COUNTROWS(RELATEDTABLE(Fact))")
        .unwrap();

    let dim = model.table("Dim").unwrap();
    assert_eq!(dim.value(0, "Related Fact Count").unwrap(), 1.into());
    assert_eq!(dim.value(1, "Related Fact Count").unwrap(), 0.into());
}

#[test]
fn relatedtable_from_virtual_blank_dimension_member_includes_unmatched_facts_m2m_for_columnar_fact() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    let schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
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
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(schema, options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(999.0),
        formula_columnar::Value::Number(7.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::Null,
        formula_columnar::Value::Number(5.0),
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

    let engine = DaxEngine::new();
    let blank_row = model.table("Dim").unwrap().row_count();

    let mut ctx = RowContext::default();
    ctx.push("Dim", blank_row);

    assert_eq!(
        engine
            .evaluate(
                &model,
                "SUMX(RELATEDTABLE(Fact), Fact[Amount])",
                &FilterContext::empty(),
                &ctx,
            )
            .unwrap(),
        12.0.into()
    );
}

#[test]
fn relatedtable_from_virtual_blank_dimension_member_includes_unmatched_facts_m2m_for_columnar_dim() {
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
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "Amount"]);
    fact.push_row(vec![1.into(), 1.into(), 10.0.into()]).unwrap();
    fact.push_row(vec![2.into(), 999.into(), 7.0.into()]).unwrap();
    fact.push_row(vec![3.into(), Value::Blank, 5.0.into()])
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    let engine = DaxEngine::new();
    let blank_row = model.table("Dim").unwrap().row_count();

    let mut ctx = RowContext::default();
    ctx.push("Dim", blank_row);

    assert_eq!(
        engine
            .evaluate(
                &model,
                "SUMX(RELATEDTABLE(Fact), Fact[Amount])",
                &FilterContext::empty(),
                &ctx,
            )
            .unwrap(),
        12.0.into()
    );
}

#[test]
fn relatedtable_from_virtual_blank_dimension_member_includes_unmatched_facts_m2m_for_columnar_dim_and_fact(
) {
    // Regression coverage: when *both* dimension and fact tables are columnar-backed, the virtual
    // blank member created by a ManyToMany relationship with RI disabled should expose both
    // unmatched fact keys and BLANK foreign keys.
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
    let dim_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, dim_options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let fact_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, fact_options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(999.0),
        formula_columnar::Value::Number(7.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::Null,
        formula_columnar::Value::Number(5.0),
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

    let engine = DaxEngine::new();
    let blank_row = model.table("Dim").unwrap().row_count();

    let mut ctx = RowContext::default();
    ctx.push("Dim", blank_row);

    assert_eq!(
        engine
            .evaluate(
                &model,
                "SUMX(RELATEDTABLE(Fact), Fact[Amount])",
                &FilterContext::empty(),
                &ctx,
            )
            .unwrap(),
        12.0.into()
    );
}

#[test]
fn relatedtable_from_physical_blank_dimension_key_does_not_include_unmatched_facts_m2m() {
    // Regression: the relationship-generated blank member is distinct from a *physical* BLANK key
    // on the dimension side. RELATEDTABLE from a physical blank-key row should not return unmatched
    // facts.
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![Value::Blank, "PhysicalBlank".into()])
        .unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "Amount"]);
    fact.push_row(vec![1.into(), 1.into(), 10.0.into()]).unwrap();
    fact.push_row(vec![2.into(), 999.into(), 7.0.into()]).unwrap();
    fact.push_row(vec![3.into(), Value::Blank, 5.0.into()])
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
        .add_measure("Total Amount", "SUM(Fact[Amount])")
        .unwrap();
    let physical_blank =
        FilterContext::empty().with_column_equals("Dim", "Attr", "PhysicalBlank".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &physical_blank).unwrap(),
        Value::Blank
    );

    let physical_blank_row = 1;
    let mut ctx = RowContext::default();
    ctx.push("Dim", physical_blank_row);

    assert_eq!(
        DaxEngine::new()
            .evaluate(
                &model,
                "SUMX(RELATEDTABLE(Fact), Fact[Amount])",
                &FilterContext::empty(),
                &ctx,
            )
            .unwrap(),
        Value::Blank
    );
}

#[test]
fn relatedtable_from_physical_blank_dimension_key_does_not_include_unmatched_facts_m2m_for_columnar_dim(
) {
    // Same regression as `relatedtable_from_physical_blank_dimension_key_does_not_include_unmatched_facts_m2m`,
    // but with a columnar-backed dimension table.
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
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    // Physical BLANK key row.
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("PhysicalBlank".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "Amount"]);
    fact.push_row(vec![1.into(), 1.into(), 10.0.into()]).unwrap();
    // Unmatched key (should map to virtual blank member, not the physical BLANK key row).
    fact.push_row(vec![2.into(), 999.into(), 7.0.into()]).unwrap();
    // BLANK FK (also belongs to virtual blank member).
    fact.push_row(vec![3.into(), Value::Blank, 5.0.into()])
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    model
        .add_measure("Total Amount", "SUM(Fact[Amount])")
        .unwrap();
    let physical_blank =
        FilterContext::empty().with_column_equals("Dim", "Attr", "PhysicalBlank".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &physical_blank).unwrap(),
        Value::Blank
    );

    let physical_blank_row = 1;
    let mut ctx = RowContext::default();
    ctx.push("Dim", physical_blank_row);

    assert_eq!(
        DaxEngine::new()
            .evaluate(
                &model,
                "SUMX(RELATEDTABLE(Fact), Fact[Amount])",
                &FilterContext::empty(),
                &ctx,
            )
            .unwrap(),
        Value::Blank
    );
}

#[test]
fn relatedtable_from_physical_blank_dimension_key_does_not_include_unmatched_facts_m2m_for_columnar_fact(
) {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![Value::Blank, "PhysicalBlank".into()])
        .unwrap();
    model.add_table(dim).unwrap();

    let schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
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
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(schema, options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(999.0),
        formula_columnar::Value::Number(7.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::Null,
        formula_columnar::Value::Number(5.0),
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
    let physical_blank =
        FilterContext::empty().with_column_equals("Dim", "Attr", "PhysicalBlank".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &physical_blank).unwrap(),
        Value::Blank
    );

    let physical_blank_row = 1;
    let mut ctx = RowContext::default();
    ctx.push("Dim", physical_blank_row);

    assert_eq!(
        DaxEngine::new()
            .evaluate(
                &model,
                "SUMX(RELATEDTABLE(Fact), Fact[Amount])",
                &FilterContext::empty(),
                &ctx,
            )
            .unwrap(),
        Value::Blank
    );
}

#[test]
fn relatedtable_from_physical_blank_dimension_key_does_not_include_unmatched_facts_m2m_for_columnar_dim_and_fact(
) {
    // Same regression as `relatedtable_from_physical_blank_dimension_key_does_not_include_unmatched_facts_m2m`,
    // but with *both* the dimension and fact tables columnar-backed.
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
    let dim_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, dim_options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    // Physical BLANK key row.
    dim.append_row(&[
        formula_columnar::Value::Null,
        formula_columnar::Value::String("PhysicalBlank".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let fact_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, fact_options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    // Unmatched key (should map to virtual blank member, not the physical BLANK key row).
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(999.0),
        formula_columnar::Value::Number(7.0),
    ]);
    // BLANK FK (also belongs to virtual blank member).
    fact.append_row(&[
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::Null,
        formula_columnar::Value::Number(5.0),
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
    let physical_blank =
        FilterContext::empty().with_column_equals("Dim", "Attr", "PhysicalBlank".into());
    assert_eq!(
        model.evaluate_measure("Total Amount", &physical_blank).unwrap(),
        Value::Blank
    );

    let physical_blank_row = 1;
    let mut ctx = RowContext::default();
    ctx.push("Dim", physical_blank_row);

    assert_eq!(
        DaxEngine::new()
            .evaluate(
                &model,
                "SUMX(RELATEDTABLE(Fact), Fact[Amount])",
                &FilterContext::empty(),
                &ctx,
            )
            .unwrap(),
        Value::Blank
    );
}

#[test]
fn relatedtable_respects_userelationship_overrides_with_m2m() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["KeyA", "KeyB", "Attr"]);
    dim.push_row(vec![1.into(), 10.into(), "A".into()]).unwrap();
    dim.push_row(vec![2.into(), 20.into(), "B".into()]).unwrap();
    dim.push_row(vec![3.into(), 30.into(), "C".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "KeyA", "KeyB"]);
    // KeyA=1 has 3 facts; KeyB=10 has 2 facts. This makes it easy to detect whether
    // USERELATIONSHIP is respected (should switch the navigation key used by RELATEDTABLE).
    fact.push_row(vec![100.into(), 1.into(), 10.into()]).unwrap();
    fact.push_row(vec![101.into(), 1.into(), 20.into()]).unwrap();
    fact.push_row(vec![102.into(), 1.into(), 30.into()]).unwrap();
    fact.push_row(vec![103.into(), 2.into(), 10.into()]).unwrap();
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
        3.into()
    );

    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(RELATEDTABLE(Fact)), USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
                &FilterContext::empty(),
                &ctx
            )
            .unwrap(),
        2.into()
    );

    // Another sanity check: for the third dimension row, KeyA has no matches but KeyB does.
    let mut ctx_c = RowContext::default();
    ctx_c.push("Dim", 2);
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(RELATEDTABLE(Fact))",
                &FilterContext::empty(),
                &ctx_c
            )
            .unwrap(),
        0.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(RELATEDTABLE(Fact)), USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
                &FilterContext::empty(),
                &ctx_c
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn relatedtable_respects_userelationship_overrides_with_m2m_for_columnar_dim() {
    let mut model = DataModel::new();

    let dim_schema = vec![
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Attr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::String("A".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(20.0),
        formula_columnar::Value::String("B".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::Number(30.0),
        formula_columnar::Value::String("C".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "KeyA", "KeyB"]);
    fact.push_row(vec![100.into(), 1.into(), 10.into()]).unwrap();
    fact.push_row(vec![101.into(), 1.into(), 20.into()]).unwrap();
    fact.push_row(vec![102.into(), 1.into(), 30.into()]).unwrap();
    fact.push_row(vec![103.into(), 2.into(), 10.into()]).unwrap();
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
        3.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(RELATEDTABLE(Fact)), USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
                &FilterContext::empty(),
                &ctx
            )
            .unwrap(),
        2.into()
    );

    let mut ctx_c = RowContext::default();
    ctx_c.push("Dim", 2);
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(RELATEDTABLE(Fact))",
                &FilterContext::empty(),
                &ctx_c
            )
            .unwrap(),
        0.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(RELATEDTABLE(Fact)), USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
                &FilterContext::empty(),
                &ctx_c
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn relatedtable_respects_userelationship_overrides_with_m2m_for_columnar_fact() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["KeyA", "KeyB", "Attr"]);
    dim.push_row(vec![1.into(), 10.into(), "A".into()]).unwrap();
    dim.push_row(vec![2.into(), 20.into(), "B".into()]).unwrap();
    dim.push_row(vec![3.into(), 30.into(), "C".into()]).unwrap();
    model.add_table(dim).unwrap();

    let schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(schema, options);
    fact.append_row(&[
        formula_columnar::Value::Number(100.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(20.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(102.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(30.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(103.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(10.0),
    ]);
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

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
        3.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(RELATEDTABLE(Fact)), USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
                &FilterContext::empty(),
                &ctx
            )
            .unwrap(),
        2.into()
    );

    let mut ctx_c = RowContext::default();
    ctx_c.push("Dim", 2);
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(RELATEDTABLE(Fact))",
                &FilterContext::empty(),
                &ctx_c
            )
            .unwrap(),
        0.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(RELATEDTABLE(Fact)), USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
                &FilterContext::empty(),
                &ctx_c
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn relatedtable_respects_userelationship_overrides_with_m2m_for_columnar_dim_and_fact() {
    let mut model = DataModel::new();

    let dim_schema = vec![
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Attr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let dim_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, dim_options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::String("A".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(20.0),
        formula_columnar::Value::String("B".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::Number(30.0),
        formula_columnar::Value::String("C".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let fact_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, fact_options);
    fact.append_row(&[
        formula_columnar::Value::Number(100.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(20.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(102.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(30.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(103.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(10.0),
    ]);
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

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
        3.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(RELATEDTABLE(Fact)), USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
                &FilterContext::empty(),
                &ctx
            )
            .unwrap(),
        2.into()
    );

    let mut ctx_c = RowContext::default();
    ctx_c.push("Dim", 2);
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(RELATEDTABLE(Fact))",
                &FilterContext::empty(),
                &ctx_c
            )
            .unwrap(),
        0.into()
    );
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(RELATEDTABLE(Fact)), USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
                &FilterContext::empty(),
                &ctx_c
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn relatedtable_errors_on_ambiguous_relationship_paths_with_m2m() {
    // Build a model where Dim -> Fact has two active relationship paths:
    //   Dim -> Fact (direct)
    //   Dim -> Bridge -> Fact
    // `RELATEDTABLE(Fact)` should error with an "ambiguous relationship path" message. Disabling one
    // relationship via CROSSFILTER should make navigation deterministic again.
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![2.into(), "B".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut bridge = Table::new("Bridge", vec!["Key"]);
    bridge.push_row(vec![1.into()]).unwrap();
    bridge.push_row(vec![2.into()]).unwrap();
    model.add_table(bridge).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "BridgeKey"]);
    // Direct Dim-Key relationship would include both facts; BridgeKey path only includes the first.
    fact.push_row(vec![1.into(), 1.into(), 1.into()]).unwrap();
    fact.push_row(vec![2.into(), 1.into(), 2.into()]).unwrap();
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
    model
        .add_relationship(Relationship {
            name: "Bridge_Dim".into(),
            from_table: "Bridge".into(),
            from_column: "Key".into(),
            to_table: "Dim".into(),
            to_column: "Key".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();
    model
        .add_relationship(Relationship {
            name: "Fact_Bridge".into(),
            from_table: "Fact".into(),
            from_column: "BridgeKey".into(),
            to_table: "Bridge".into(),
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

    let err = engine
        .evaluate(
            &model,
            "COUNTROWS(RELATEDTABLE(Fact))",
            &FilterContext::empty(),
            &ctx,
        )
        .unwrap_err();
    let msg = err.to_string().to_ascii_lowercase();
    assert!(
        msg.contains("ambiguous") && msg.contains("relationship path"),
        "unexpected error: {err}"
    );

    // Disable the Bridge->Dim relationship: only the direct path remains.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(RELATEDTABLE(Fact)), CROSSFILTER(Bridge[Key], Dim[Key], NONE))",
                &FilterContext::empty(),
                &ctx
            )
            .unwrap(),
        2.into()
    );

    // Disable the direct Fact->Dim relationship: only the Bridge path remains.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(RELATEDTABLE(Fact)), CROSSFILTER(Fact[Key], Dim[Key], NONE))",
                &FilterContext::empty(),
                &ctx
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn relatedtable_errors_on_ambiguous_relationship_paths_with_m2m_for_columnar_dim_and_fact() {
    // Same ambiguity regression as `relatedtable_errors_on_ambiguous_relationship_paths_with_m2m`,
    // but with both Dim and Fact columnar-backed.
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
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("A".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String("B".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let mut bridge = Table::new("Bridge", vec!["Key"]);
    bridge.push_row(vec![1.into()]).unwrap();
    bridge.push_row(vec![2.into()]).unwrap();
    model.add_table(bridge).unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "BridgeKey".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(2.0),
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
            enforce_referential_integrity: true,
        })
        .unwrap();
    model
        .add_relationship(Relationship {
            name: "Bridge_Dim".into(),
            from_table: "Bridge".into(),
            from_column: "Key".into(),
            to_table: "Dim".into(),
            to_column: "Key".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();
    model
        .add_relationship(Relationship {
            name: "Fact_Bridge".into(),
            from_table: "Fact".into(),
            from_column: "BridgeKey".into(),
            to_table: "Bridge".into(),
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

    let err = engine
        .evaluate(
            &model,
            "COUNTROWS(RELATEDTABLE(Fact))",
            &FilterContext::empty(),
            &ctx,
        )
        .unwrap_err();
    let msg = err.to_string().to_ascii_lowercase();
    assert!(
        msg.contains("ambiguous") && msg.contains("relationship path"),
        "unexpected error: {err}"
    );

    // Disable the Bridge->Dim relationship: only the direct path remains.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(RELATEDTABLE(Fact)), CROSSFILTER(Bridge[Key], Dim[Key], NONE))",
                &FilterContext::empty(),
                &ctx
            )
            .unwrap(),
        2.into()
    );

    // Disable the direct Fact->Dim relationship: only the Bridge path remains.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(COUNTROWS(RELATEDTABLE(Fact)), CROSSFILTER(Fact[Key], Dim[Key], NONE))",
                &FilterContext::empty(),
                &ctx
            )
            .unwrap(),
        1.into()
    );
}

#[test]
fn related_errors_on_ambiguous_relationship_paths_with_m2m() {
    // Similar to the RELATEDTABLE ambiguity test above, but for RELATED navigation (fact -> dim).
    // There are two active paths from Fact to Dim:
    //   Fact -> Dim (direct)
    //   Fact -> Bridge -> Dim
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "Direct".into()]).unwrap();
    dim.push_row(vec![2.into(), "ViaBridge".into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut bridge = Table::new("Bridge", vec!["BridgeKey", "DimKey"]);
    bridge.push_row(vec![10.into(), 2.into()]).unwrap();
    model.add_table(bridge).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "BridgeKey"]);
    fact.push_row(vec![1.into(), 1.into(), 10.into()]).unwrap();
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
    model
        .add_relationship(Relationship {
            name: "Fact_Bridge".into(),
            from_table: "Fact".into(),
            from_column: "BridgeKey".into(),
            to_table: "Bridge".into(),
            to_column: "BridgeKey".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();
    model
        .add_relationship(Relationship {
            name: "Bridge_Dim".into(),
            from_table: "Bridge".into(),
            from_column: "DimKey".into(),
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
    ctx.push("Fact", 0);

    let err = engine
        .evaluate(
            &model,
            "RELATED(Dim[Attr])",
            &FilterContext::empty(),
            &ctx,
        )
        .unwrap_err();
    let msg = err.to_string().to_ascii_lowercase();
    assert!(
        msg.contains("ambiguous") && msg.contains("relationship path"),
        "unexpected error: {err}"
    );

    // Disable the Bridge->Dim hop to force the direct Fact->Dim path.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(RELATED(Dim[Attr]), CROSSFILTER(Bridge[DimKey], Dim[Key], NONE))",
                &FilterContext::empty(),
                &ctx,
            )
            .unwrap(),
        "Direct".into()
    );

    // Disable the direct Fact->Dim relationship to force the Bridge path.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(RELATED(Dim[Attr]), CROSSFILTER(Fact[Key], Dim[Key], NONE))",
                &FilterContext::empty(),
                &ctx,
            )
            .unwrap(),
        "ViaBridge".into()
    );
}

#[test]
fn related_errors_on_ambiguous_relationship_paths_with_m2m_for_columnar_dim_and_fact() {
    // Same ambiguity regression as `related_errors_on_ambiguous_relationship_paths_with_m2m`, but
    // with both Dim and Fact columnar-backed.
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
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String("Direct".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String("ViaBridge".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let mut bridge = Table::new("Bridge", vec!["BridgeKey", "DimKey"]);
    bridge.push_row(vec![10.into(), 2.into()]).unwrap();
    model.add_table(bridge).unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "BridgeKey".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
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
            enforce_referential_integrity: true,
        })
        .unwrap();
    model
        .add_relationship(Relationship {
            name: "Fact_Bridge".into(),
            from_table: "Fact".into(),
            from_column: "BridgeKey".into(),
            to_table: "Bridge".into(),
            to_column: "BridgeKey".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();
    model
        .add_relationship(Relationship {
            name: "Bridge_Dim".into(),
            from_table: "Bridge".into(),
            from_column: "DimKey".into(),
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
    ctx.push("Fact", 0);

    let err = engine
        .evaluate(
            &model,
            "RELATED(Dim[Attr])",
            &FilterContext::empty(),
            &ctx,
        )
        .unwrap_err();
    let msg = err.to_string().to_ascii_lowercase();
    assert!(
        msg.contains("ambiguous") && msg.contains("relationship path"),
        "unexpected error: {err}"
    );

    // Disable the Bridge->Dim hop to force the direct Fact->Dim path.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(RELATED(Dim[Attr]), CROSSFILTER(Bridge[DimKey], Dim[Key], NONE))",
                &FilterContext::empty(),
                &ctx,
            )
            .unwrap(),
        "Direct".into()
    );

    // Disable the direct Fact->Dim relationship to force the Bridge path.
    assert_eq!(
        engine
            .evaluate(
                &model,
                "CALCULATE(RELATED(Dim[Attr]), CROSSFILTER(Fact[Key], Dim[Key], NONE))",
                &FilterContext::empty(),
                &ctx,
            )
            .unwrap(),
        "ViaBridge".into()
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
fn insert_row_can_create_unmatched_fact_rows_and_updates_blank_member() {
    // Regression for incremental updates on the fact side: inserting an unmatched FK row after the
    // relationship exists should materialize the virtual blank member, and inserting the
    // corresponding dimension key should "rescue" the fact row out of that member.
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    model.add_measure("Total Amount", "SUM(Fact[Amount])").unwrap();

    let blank_attr = FilterContext::empty().with_column_equals("Dim", "Attr", Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total Amount", &blank_attr).unwrap(),
        Value::Blank
    );

    // Insert an unmatched FK row after the relationship is defined.
    model
        .insert_row("Fact", vec![2.into(), 999.into(), 7.0.into()])
        .unwrap();

    // The unmatched row should now appear under the virtual blank member.
    assert_eq!(
        model.evaluate_measure("Total Amount", &blank_attr).unwrap(),
        7.0.into()
    );

    // Insert the missing dimension key to "rescue" the row.
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
fn related_respects_userelationship_overrides_with_m2m_for_columnar_dim() {
    let mut model = DataModel::new();

    let dim_schema = vec![
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Attr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::String("RowA".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(20.0),
        formula_columnar::Value::String("RowB".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "KeyA", "KeyB"]);
    // Cross keys so the active vs. USERELATIONSHIP-overridden relationship produces different values.
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
fn related_respects_userelationship_overrides_with_m2m_for_columnar_fact() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["KeyA", "KeyB", "Attr"]);
    dim.push_row(vec![1.into(), 10.into(), "RowA".into()]).unwrap();
    dim.push_row(vec![2.into(), 20.into(), "RowB".into()]).unwrap();
    model.add_table(dim).unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, options);
    // Cross the keys so the active vs. USERELATIONSHIP-overridden relationship produces different
    // RELATED values.
    fact.append_row(&[
        formula_columnar::Value::Number(100.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(20.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(10.0),
    ]);
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

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
fn related_respects_userelationship_overrides_with_m2m_for_columnar_dim_and_fact() {
    let mut model = DataModel::new();

    let dim_schema = vec![
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Attr".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let dim_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut dim = ColumnarTableBuilder::new(dim_schema, dim_options);
    dim.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::String("RowA".into()),
    ]);
    dim.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(20.0),
        formula_columnar::Value::String("RowB".into()),
    ]);
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let fact_options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, fact_options);
    // Cross keys so the active vs. USERELATIONSHIP-overridden relationship produces different values.
    fact.append_row(&[
        formula_columnar::Value::Number(100.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(20.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(101.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(10.0),
    ]);
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

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

#[test]
fn insert_row_rolls_back_when_calculated_column_errors_under_m2m() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    model.add_table(Table::new("Fact", vec!["Id", "Key"])).unwrap();

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

    model
        .add_calculated_column("Fact", "DimAttr", "RELATED(Dim[Attr])")
        .unwrap();

    model.insert_row("Fact", vec![1.into(), 1.into()]).unwrap();
    assert_eq!(model.table("Fact").unwrap().row_count(), 1);
    assert_eq!(
        model.table("Fact").unwrap().value(0, "DimAttr").unwrap(),
        "A".into()
    );

    // Make the RELATED lookup ambiguous by inserting a second Dim row for the same key.
    model
        .insert_row("Dim", vec![1.into(), "B".into()])
        .unwrap();

    let err = model.insert_row("Fact", vec![2.into(), 1.into()]).unwrap_err();
    let msg = err.to_string().to_ascii_lowercase();
    assert!(
        msg.contains("ambig") || msg.contains("multiple") || msg.contains("more than one"),
        "unexpected insert_row error: {err}"
    );

    // insert_row should be atomic: the row should not be present after the error.
    assert_eq!(model.table("Fact").unwrap().row_count(), 1);
}

#[test]
fn insert_row_updates_unmatched_fact_rows_for_columnar_m2m_relationships() {
    // Regression for incremental updates: when the fact table is columnar, the relationship stores
    // a cached list of "unmatched" fact rows for blank-member semantics. Inserting a new dimension
    // key should update that cache so previously-unmatched facts no longer appear under BLANK().
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
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
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(999.0),
        formula_columnar::Value::Number(7.0),
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

    model.add_measure("Total Amount", "SUM(Fact[Amount])").unwrap();

    let blank_attr = FilterContext::empty().with_column_equals("Dim", "Attr", Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total Amount", &blank_attr).unwrap(),
        7.0.into()
    );

    // Insert a new dimension row that resolves the previously-unmatched key.
    model
        .insert_row("Dim", vec![999.into(), "New".into()])
        .unwrap();

    // The fact row should move out of the virtual blank member and under the new Dim row.
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
fn insert_row_updates_unmatched_fact_rows_for_columnar_m2m_relationships_when_cache_is_sparse() {
    // Similar to `insert_row_updates_unmatched_fact_rows_for_columnar_m2m_relationships`, but sized
    // so the relationship's `unmatched_fact_rows` cache stays in the *sparse* representation.
    //
    // This exercises the `UnmatchedFactRows::Sparse` update path (retain-based removal).
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    model.add_table(dim).unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
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
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, options);

    // 64 total rows => sparse_to_dense_threshold = 1. With exactly one unmatched row, the cache
    // remains sparse.
    for id in 1..=63 {
        fact.append_row(&[
            formula_columnar::Value::Number(id as f64),
            formula_columnar::Value::Number(1.0),
            formula_columnar::Value::Number(1.0),
        ]);
    }
    // One unmatched FK (999).
    fact.append_row(&[
        formula_columnar::Value::Number(64.0),
        formula_columnar::Value::Number(999.0),
        formula_columnar::Value::Number(7.0),
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

    model.add_measure("Total Amount", "SUM(Fact[Amount])").unwrap();

    let blank_attr = FilterContext::empty().with_column_equals("Dim", "Attr", Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total Amount", &blank_attr).unwrap(),
        7.0.into()
    );

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
fn userelationship_override_works_with_m2m_for_columnar_fact() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["KeyA", "KeyB", "Attr"]);
    dim.push_row(vec![1.into(), 10.into(), "A".into()]).unwrap();
    dim.push_row(vec![2.into(), 20.into(), "B".into()]).unwrap();
    model.add_table(dim).unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, options);
    // Cross the keys so the active vs. inactive relationship yields different totals under the
    // same Dim filter.
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(20.0),
        formula_columnar::Value::Number(100.0),
    ]);
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(200.0),
    ]);
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

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

    let filter_b = FilterContext::empty().with_column_equals("Dim", "Attr", "B".into());
    assert_eq!(model.evaluate_measure("Total", &filter_b).unwrap(), 200.0.into());
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &filter_b).unwrap(),
        100.0.into()
    );
}

#[test]
fn insert_row_resolves_blank_member_for_inactive_userelationship_with_columnar_fact() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["KeyA", "KeyB", "Attr"]);
    dim.push_row(vec![1.into(), 10.into(), "A".into()]).unwrap();
    dim.push_row(vec![2.into(), 20.into(), "B".into()]).unwrap();
    // Physical BLANK key row on the dimension side should not behave like the relationship-generated
    // blank/unknown member under USERELATIONSHIP.
    dim.push_row(vec![Value::Blank, Value::Blank, "PhysicalBlank".into()])
        .unwrap();
    model.add_table(dim).unwrap();

    let fact_schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyA".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "KeyB".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut fact = ColumnarTableBuilder::new(fact_schema, options);
    fact.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(5.0),
    ]);
    // Unmatched for KeyB (relationship B below), but still matched for KeyA (relationship A).
    fact.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(999.0),
        formula_columnar::Value::Number(7.0),
    ]);
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

    // Active relationship A (RI enforced).
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

    // Inactive relationship B (RI disabled) so we can test the virtual blank member via USERELATIONSHIP.
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
            enforce_referential_integrity: false,
        })
        .unwrap();

    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
    model
        .add_measure(
            "Total via KeyB",
            "CALCULATE([Total], USERELATIONSHIP(Fact[KeyB], Dim[KeyB]))",
        )
        .unwrap();

    let blank_attr = FilterContext::empty().with_column_equals("Dim", "Attr", Value::Blank);
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &blank_attr).unwrap(),
        7.0.into()
    );

    let physical_blank =
        FilterContext::empty().with_column_equals("Dim", "Attr", "PhysicalBlank".into());
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &physical_blank).unwrap(),
        Value::Blank
    );

    // Insert a Dim row that resolves the previously-unmatched KeyB value.
    model
        .insert_row("Dim", vec![3.into(), 999.into(), "New".into()])
        .unwrap();

    // The virtual blank member should disappear for relationship B under USERELATIONSHIP.
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &blank_attr).unwrap(),
        Value::Blank
    );
    let new_attr = FilterContext::empty().with_column_equals("Dim", "Attr", "New".into());
    assert_eq!(
        model.evaluate_measure("Total via KeyB", &new_attr).unwrap(),
        7.0.into()
    );
}
