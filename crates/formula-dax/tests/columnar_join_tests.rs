use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{DataModel, Table, TableBackend};
use std::sync::Arc;

#[test]
fn columnar_backend_hash_join_returns_expected_pairs() {
    let schema = vec![
        ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Value".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    };

    let mut left_builder = ColumnarTableBuilder::new(schema.clone(), options);
    left_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("L1")),
    ]);
    left_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("L2")),
    ]);
    left_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("L3")),
    ]);
    left_builder.append_row(&[
        formula_columnar::Value::Number(3.0),
        formula_columnar::Value::String(Arc::<str>::from("L4")),
    ]);

    let mut right_builder = ColumnarTableBuilder::new(schema, options);
    right_builder.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("R1")),
    ]);
    right_builder.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("R2")),
    ]);

    let mut model = DataModel::new();
    model
        .add_table(Table::from_columnar("Left", left_builder.finalize()))
        .unwrap();
    model
        .add_table(Table::from_columnar("Right", right_builder.finalize()))
        .unwrap();

    let left = model.table("Left").unwrap();
    let right = model.table("Right").unwrap();

    let joined = left.hash_join(right, 0, 0).expect("join supported");
    assert_eq!(joined.left_indices, vec![0, 1, 2]);
    assert_eq!(joined.right_indices, vec![0, 0, 1]);
}

