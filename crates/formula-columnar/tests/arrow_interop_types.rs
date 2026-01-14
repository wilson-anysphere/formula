#![cfg(feature = "arrow")]

use arrow_array::builder::StringDictionaryBuilder;
use arrow_array::{
    ArrayRef, Float32Array, Int16Array, Int32Array, RecordBatch, UInt16Array, UInt32Array,
    UInt64Array,
};
use arrow_schema::{DataType, Field, Schema};
use formula_columnar::arrow::{columnar_to_record_batch, record_batch_to_columnar};
use formula_columnar::{ColumnType, Value};
use std::sync::Arc;

fn assert_tables_equal(a: &formula_columnar::ColumnarTable, b: &formula_columnar::ColumnarTable) {
    assert_eq!(a.row_count(), b.row_count());
    assert_eq!(a.column_count(), b.column_count());
    assert_eq!(a.schema(), b.schema());

    for row in 0..a.row_count() {
        for col in 0..a.column_count() {
            assert_eq!(a.get_cell(row, col), b.get_cell(row, col), "r={row} c={col}");
        }
    }
}

#[test]
fn record_batch_to_columnar_accepts_common_numeric_and_dictionary_types(
) -> Result<(), Box<dyn std::error::Error>> {
    // Build a batch using alternative physical Arrow types (Float32/Int32/UInt32/UInt64)
    // and a dictionary-encoded string column with a non-UInt32 key type (Int32).
    let f32_arr = Arc::new(Float32Array::from(vec![
        Some(1.25_f32),
        None,
        Some(-3.5_f32),
        Some(0.0_f32),
    ])) as ArrayRef;
    let i16_arr = Arc::new(Int16Array::from(vec![Some(-7), Some(0), None, Some(123)])) as ArrayRef;
    let u16_arr =
        Arc::new(UInt16Array::from(vec![Some(1), Some(2), None, Some(65_535)])) as ArrayRef;
    let i32_arr = Arc::new(Int32Array::from(vec![Some(-1), Some(0), None, Some(123)])) as ArrayRef;
    let u32_arr = Arc::new(UInt32Array::from(vec![Some(5), None, Some(42), Some(0)])) as ArrayRef;
    let u64_arr = Arc::new(UInt64Array::from(vec![
        Some(9_007_199_254_740_991_u64),
        Some(42_u64),
        None,
        Some(1_234_u64),
    ])) as ArrayRef;

    let mut dict_builder =
        StringDictionaryBuilder::<arrow_array::types::Int32Type>::new();
    dict_builder.append("A")?;
    dict_builder.append("B")?;
    dict_builder.append_null();
    dict_builder.append("A")?;
    let dict_arr = Arc::new(dict_builder.finish()) as ArrayRef;

    let mut dict_u16_builder = StringDictionaryBuilder::<arrow_array::types::UInt16Type>::new();
    dict_u16_builder.append("X")?;
    dict_u16_builder.append_null();
    dict_u16_builder.append("Y")?;
    dict_u16_builder.append("X")?;
    let dict_u16_arr = Arc::new(dict_u16_builder.finish()) as ArrayRef;

    let schema = Arc::new(Schema::new(vec![
        Field::new("f32", DataType::Float32, true),
        Field::new("i16", DataType::Int16, true),
        Field::new("u16", DataType::UInt16, true),
        Field::new("i32", DataType::Int32, true),
        Field::new("u32", DataType::UInt32, true),
        Field::new("u64", DataType::UInt64, true),
        Field::new(
            "dict_i32",
            DataType::Dictionary(Box::new(DataType::Int32), Box::new(DataType::Utf8)),
            true,
        ),
        Field::new(
            "dict_u16",
            DataType::Dictionary(Box::new(DataType::UInt16), Box::new(DataType::Utf8)),
            true,
        ),
    ]));
    let batch = RecordBatch::try_new(
        schema,
        vec![
            f32_arr,
            i16_arr,
            u16_arr,
            i32_arr,
            u32_arr,
            u64_arr,
            dict_arr,
            dict_u16_arr,
        ],
    )?;

    let table = record_batch_to_columnar(&batch)?;

    assert_eq!(table.row_count(), 4);
    assert_eq!(
        table
            .schema()
            .iter()
            .map(|c| c.column_type)
            .collect::<Vec<_>>(),
        vec![
            ColumnType::Number,
            ColumnType::Number,
            ColumnType::Number,
            ColumnType::Number,
            ColumnType::Number,
            ColumnType::Number,
            ColumnType::String,
            ColumnType::String,
        ]
    );

    let expected: Vec<Vec<Value>> = vec![
        vec![
            Value::Number(1.25_f32 as f64),
            Value::Number(-7.0),
            Value::Number(1.0),
            Value::Number(-1.0),
            Value::Number(5.0),
            Value::Number(9_007_199_254_740_991_u64 as f64),
            Value::String(Arc::<str>::from("A")),
            Value::String(Arc::<str>::from("X")),
        ],
        vec![
            Value::Null,
            Value::Number(0.0),
            Value::Number(2.0),
            Value::Number(0.0),
            Value::Null,
            Value::Number(42_u64 as f64),
            Value::String(Arc::<str>::from("B")),
            Value::Null,
        ],
        vec![
            Value::Number(-3.5_f32 as f64),
            Value::Null,
            Value::Null,
            Value::Null,
            Value::Number(42.0),
            Value::Null,
            Value::Null,
            Value::String(Arc::<str>::from("Y")),
        ],
        vec![
            Value::Number(0.0),
            Value::Number(123.0),
            Value::Number(65_535.0),
            Value::Number(123.0),
            Value::Number(0.0),
            Value::Number(1_234_u64 as f64),
            Value::String(Arc::<str>::from("A")),
            Value::String(Arc::<str>::from("X")),
        ],
    ];

    for (row, row_vals) in expected.iter().enumerate() {
        for (col, expected_val) in row_vals.iter().enumerate() {
            assert_eq!(
                table.get_cell(row, col),
                expected_val.clone(),
                "mismatch r={row} c={col}"
            );
        }
    }

    // Round-trip through the export path as well.
    let batch2 = columnar_to_record_batch(&table)?;
    let table2 = record_batch_to_columnar(&batch2)?;
    assert_tables_equal(&table, &table2);

    Ok(())
}
