use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{Cardinality, CrossFilterDirection, DataModel, FilterContext, Relationship, Table, Value};
use pretty_assertions::assert_eq;
use std::sync::Arc;

fn build_models() -> (DataModel, DataModel) {
    let mut vec_model = DataModel::new();

    let mut dim = Table::new("Dim", vec!["Key", "Attr"]);
    dim.push_row(vec![1.into(), "A".into()]).unwrap();
    dim.push_row(vec![1.into(), "A2".into()]).unwrap();
    dim.push_row(vec![2.into(), "B".into()]).unwrap();
    vec_model.add_table(dim).unwrap();

    let mut fact = Table::new("Fact", vec!["Id", "Key", "Amount"]);
    fact.push_row(vec![10.into(), 1.into(), 5.0.into()]).unwrap();
    fact.push_row(vec![11.into(), 2.into(), 7.0.into()]).unwrap();
    vec_model.add_table(fact).unwrap();

    vec_model
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
    vec_model.add_measure("Total", "SUM(Fact[Amount])").unwrap();

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };

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
    let mut dim_builder = ColumnarTableBuilder::new(dim_schema, options);
    dim_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("A")),
    ]);
    dim_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("A2")),
    ]);
    dim_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("B")),
    ]);

    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };

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
    let mut fact_builder = ColumnarTableBuilder::new(fact_schema, options);
    fact_builder.append_row(&[
        formula_columnar::Value::Number(10.0),
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(5.0),
    ]);
    fact_builder.append_row(&[
        formula_columnar::Value::Number(11.0),
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::Number(7.0),
    ]);

    let mut col_model = DataModel::new();
    col_model
        .add_table(Table::from_columnar("Dim", dim_builder.finalize()))
        .unwrap();
    col_model
        .add_table(Table::from_columnar("Fact", fact_builder.finalize()))
        .unwrap();
    col_model
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
    col_model.add_measure("Total", "SUM(Fact[Amount])").unwrap();

    (vec_model, col_model)
}

#[test]
fn many_to_many_matches_between_vec_and_columnar_backends() {
    let (vec_model, col_model) = build_models();

    let cases = vec![
        ("no filters", FilterContext::empty(), Value::from(12.0)),
        (
            "Dim[Attr] = \"A\"",
            FilterContext::empty().with_column_equals("Dim", "Attr", "A".into()),
            5.0.into(),
        ),
        (
            "Dim[Attr] = \"B\"",
            FilterContext::empty().with_column_equals("Dim", "Attr", "B".into()),
            7.0.into(),
        ),
        (
            "Dim[Attr] = \"A2\"",
            FilterContext::empty().with_column_equals("Dim", "Attr", "A2".into()),
            5.0.into(),
        ),
    ];

    for (name, filter, expected) in cases {
        let vec_value = vec_model.evaluate_measure("Total", &filter).unwrap();
        let col_value = col_model.evaluate_measure("Total", &filter).unwrap();

        assert_eq!(vec_value, expected, "vec backend case: {name}");
        assert_eq!(col_value, expected, "columnar backend case: {name}");
        assert_eq!(vec_value, col_value, "parity case: {name}");
    }
}

