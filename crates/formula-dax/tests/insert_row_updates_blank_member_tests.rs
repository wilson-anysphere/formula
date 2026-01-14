use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, Relationship, RowContext,
    Table, Value,
};
use pretty_assertions::assert_eq;

fn options() -> TableOptions {
    TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 2 },
    }
}

#[test]
fn insert_row_into_dimension_updates_columnar_fact_blank_member() {
    let mut model = DataModel::new();

    // In-memory dimension table.
    let mut dim = Table::new("Dim", vec!["Id"]);
    dim.push_row(vec![1.into()]).unwrap();
    model.add_table(dim).unwrap();

    // Columnar fact table with one matched key and one initially-unmatched key.
    let schema = vec![ColumnSchema {
        name: "Id".to_string(),
        column_type: ColumnType::Number,
    }];
    let mut fact_builder = ColumnarTableBuilder::new(schema, options());
    fact_builder.append_row(&[formula_columnar::Value::Number(1.0)]);
    fact_builder.append_row(&[formula_columnar::Value::Number(999.0)]);
    model
        .add_table(Table::from_columnar("Fact", fact_builder.finalize()))
        .unwrap();

    model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "Id".into(),
            to_table: "Dim".into(),
            to_column: "Id".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: false,
        })
        .unwrap();

    let engine = DaxEngine::new();
    let filter =
        FilterContext::empty().with_column_in("Dim", "Id", vec![1.into(), Value::Blank]);

    // Before insert: selecting the BLANK member also includes unmatched fact keys.
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(Fact)", &filter, &RowContext::default())
            .unwrap(),
        2.into()
    );

    // Insert a new dimension member that rescues the previously-unmatched fact key.
    model.insert_row("Dim", vec![999.into()]).unwrap();

    // After insert: the fact row is now matched to Dim[Id]=999 and should no longer be included
    // when filtering Dim[Id] IN {1, BLANK()}.
    assert_eq!(
        engine
            .evaluate(&model, "COUNTROWS(Fact)", &filter, &RowContext::default())
            .unwrap(),
        1.into()
    );
}

