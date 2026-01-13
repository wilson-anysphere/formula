use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions,
};
use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, DaxError, Relationship, Table,
};
use std::sync::Arc;

#[test]
fn columnar_relationship_rejects_mismatched_join_column_types() {
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };

    let dim_schema = vec![ColumnSchema {
        name: "Id".to_string(),
        column_type: ColumnType::String,
    }];
    let mut dim = ColumnarTableBuilder::new(dim_schema, options);
    dim.append_row(&[formula_columnar::Value::String(Arc::<str>::from("1"))]);

    let fact_schema = vec![ColumnSchema {
        name: "Id".to_string(),
        column_type: ColumnType::Number,
    }];
    let mut fact = ColumnarTableBuilder::new(fact_schema, options);
    fact.append_row(&[formula_columnar::Value::Number(1.0)]);

    let mut model = DataModel::new();
    model
        .add_table(Table::from_columnar("Dim", dim.finalize()))
        .unwrap();
    model
        .add_table(Table::from_columnar("Fact", fact.finalize()))
        .unwrap();

    let err = model
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
        .unwrap_err();

    match err {
        DaxError::RelationshipJoinColumnTypeMismatch {
            from_type, to_type, ..
        } => {
            assert_eq!(from_type, "Number");
            assert_eq!(to_type, "String");
        }
        other => panic!("expected RelationshipJoinColumnTypeMismatch, got {other:?}"),
    }
}

#[test]
fn in_memory_relationship_rejects_mismatched_join_column_types() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Id"]);
    dim.push_row(vec![1.into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id"]);
    fact.push_row(vec!["1".into()]).unwrap();
    model.add_table(fact).unwrap();

    let err = model
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
        .unwrap_err();

    match err {
        DaxError::RelationshipJoinColumnTypeMismatch {
            from_type, to_type, ..
        } => {
            assert_eq!(from_type, "Text");
            assert_eq!(to_type, "Number");
        }
        other => panic!("expected RelationshipJoinColumnTypeMismatch, got {other:?}"),
    }
}

#[test]
fn valid_relationships_still_work() {
    let mut model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Id"]);
    dim.push_row(vec![1.into()]).unwrap();
    dim.push_row(vec![2.into()]).unwrap();
    model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id"]);
    fact.push_row(vec![1.into()]).unwrap();
    fact.push_row(vec![1.into()]).unwrap();
    model.add_table(fact).unwrap();

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
            enforce_referential_integrity: true,
        })
        .unwrap();
}

