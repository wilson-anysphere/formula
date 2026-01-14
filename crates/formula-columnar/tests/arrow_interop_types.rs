#![cfg(feature = "arrow")]

use arrow_array::builder::StringDictionaryBuilder;
use arrow_array::{
    ArrayRef, DictionaryArray, Float16Array, Float32Array, Int16Array, Int32Array, RecordBatch,
    StringViewArray, UInt16Array, UInt32Array, UInt64Array,
};
use arrow_schema::{DataType, Field, Schema};
use formula_columnar::arrow::{columnar_to_record_batch, record_batch_to_columnar};
use formula_columnar::{ColumnType, Value};
use half::f16;
use std::collections::HashMap;
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
    let f16_arr = Arc::new(Float16Array::from(vec![
        Some(f16::from_f32(1.5)),
        None,
        Some(f16::from_f32(-2.0)),
        Some(f16::from_f32(0.0)),
    ])) as ArrayRef;
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
    let view_arr = Arc::new(StringViewArray::from(vec![
        Some("V1"),
        None,
        Some("V2"),
        Some("V1"),
    ])) as ArrayRef;
    let dict_view_values = Arc::new(StringViewArray::from(vec!["DV0", "DV1"])) as ArrayRef;
    let dict_view_keys = Int32Array::from(vec![Some(0), Some(1), None, Some(0)]);
    let dict_view_arr = Arc::new(DictionaryArray::<arrow_array::types::Int32Type>::try_new(
        dict_view_keys,
        dict_view_values,
    )?) as ArrayRef;

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
        Field::new("f16", DataType::Float16, true),
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
        Field::new("view", DataType::Utf8View, true),
        Field::new(
            "dict_view",
            DataType::Dictionary(Box::new(DataType::Int32), Box::new(DataType::Utf8View)),
            true,
        ),
    ]));
    let batch = RecordBatch::try_new(
        schema,
        vec![
            f16_arr,
            f32_arr,
            i16_arr,
            u16_arr,
            i32_arr,
            u32_arr,
            u64_arr,
            dict_arr,
            dict_u16_arr,
            view_arr,
            dict_view_arr,
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
            ColumnType::Number,
            ColumnType::String,
            ColumnType::String,
            ColumnType::String,
            ColumnType::String,
        ]
    );

    let expected: Vec<Vec<Value>> = vec![
        vec![
            Value::Number(1.5),
            Value::Number(1.25_f32 as f64),
            Value::Number(-7.0),
            Value::Number(1.0),
            Value::Number(-1.0),
            Value::Number(5.0),
            Value::Number(9_007_199_254_740_991_u64 as f64),
            Value::String(Arc::<str>::from("A")),
            Value::String(Arc::<str>::from("X")),
            Value::String(Arc::<str>::from("V1")),
            Value::String(Arc::<str>::from("DV0")),
        ],
        vec![
            Value::Null,
            Value::Null,
            Value::Number(0.0),
            Value::Number(2.0),
            Value::Number(0.0),
            Value::Null,
            Value::Number(42_u64 as f64),
            Value::String(Arc::<str>::from("B")),
            Value::Null,
            Value::Null,
            Value::String(Arc::<str>::from("DV1")),
        ],
        vec![
            Value::Number(-2.0),
            Value::Number(-3.5_f32 as f64),
            Value::Null,
            Value::Null,
            Value::Null,
            Value::Number(42.0),
            Value::Null,
            Value::Null,
            Value::String(Arc::<str>::from("Y")),
            Value::String(Arc::<str>::from("V2")),
            Value::Null,
        ],
        vec![
            Value::Number(0.0),
            Value::Number(0.0),
            Value::Number(123.0),
            Value::Number(65_535.0),
            Value::Number(123.0),
            Value::Number(0.0),
            Value::Number(1_234_u64 as f64),
            Value::String(Arc::<str>::from("A")),
            Value::String(Arc::<str>::from("X")),
            Value::String(Arc::<str>::from("V1")),
            Value::String(Arc::<str>::from("DV0")),
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

#[test]
fn record_batch_to_columnar_treats_null_dictionary_values_as_nulls(
) -> Result<(), Box<dyn std::error::Error>> {
    use arrow_array::{DictionaryArray, Int32Array, StringArray};

    // Dictionary values are allowed to contain nulls. We should treat rows that
    // reference such values as Value::Null.
    let values: ArrayRef = Arc::new(StringArray::from(vec![Some("A"), None, Some("B")]));
    let keys = Int32Array::from(vec![Some(0), Some(1), Some(2)]);
    let dict = Arc::new(DictionaryArray::<arrow_array::types::Int32Type>::try_new(
        keys, values,
    )?) as ArrayRef;

    let schema = Arc::new(Schema::new(vec![Field::new(
        "dict",
        DataType::Dictionary(Box::new(DataType::Int32), Box::new(DataType::Utf8)),
        true,
    )]));
    let batch = RecordBatch::try_new(schema, vec![dict])?;

    let table = record_batch_to_columnar(&batch)?;
    assert_eq!(table.row_count(), 3);
    assert_eq!(table.get_cell(0, 0), Value::String(Arc::<str>::from("A")));
    assert_eq!(table.get_cell(1, 0), Value::Null);
    assert_eq!(table.get_cell(2, 0), Value::String(Arc::<str>::from("B")));
    Ok(())
}

#[test]
fn record_batch_to_columnar_accepts_decimal128_as_number() -> Result<(), Box<dyn std::error::Error>>
{
    use arrow_array::Decimal128Array;

    let arr = Decimal128Array::from(vec![Some(12_345_i128), None, Some(-100_i128)])
        .with_precision_and_scale(10, 2)?;
    let schema = Arc::new(Schema::new(vec![Field::new(
        "dec",
        DataType::Decimal128(10, 2),
        true,
    )]));
    let batch = RecordBatch::try_new(schema, vec![Arc::new(arr) as ArrayRef])?;

    let table = record_batch_to_columnar(&batch)?;
    assert_eq!(table.row_count(), 3);
    assert_eq!(table.get_cell(0, 0), Value::Number(123.45));
    assert_eq!(table.get_cell(1, 0), Value::Null);
    assert_eq!(table.get_cell(2, 0), Value::Number(-1.0));
    Ok(())
}

#[test]
fn record_batch_to_columnar_accepts_currency_decimal128_with_metadata(
) -> Result<(), Box<dyn std::error::Error>> {
    use arrow_array::Decimal128Array;

    let arr = Decimal128Array::from(vec![Some(12_345_i128), None, Some(-5_i128)])
        .with_precision_and_scale(10, 2)?;

    let mut meta = HashMap::new();
    meta.insert("formula:column_type".to_owned(), "currency".to_owned());
    meta.insert("formula:scale".to_owned(), "2".to_owned());

    let field = Field::new("money", DataType::Decimal128(10, 2), true).with_metadata(meta);
    let schema = Arc::new(Schema::new(vec![field]));
    let batch = RecordBatch::try_new(schema, vec![Arc::new(arr) as ArrayRef])?;

    let table = record_batch_to_columnar(&batch)?;
    assert_eq!(table.schema()[0].column_type, ColumnType::Currency { scale: 2 });
    assert_eq!(table.get_cell(0, 0), Value::Currency(12_345));
    assert_eq!(table.get_cell(1, 0), Value::Null);
    assert_eq!(table.get_cell(2, 0), Value::Currency(-5));
    Ok(())
}

#[test]
fn record_batch_to_columnar_accepts_percentage_decimal128_with_metadata(
) -> Result<(), Box<dyn std::error::Error>> {
    use arrow_array::Decimal128Array;

    let arr = Decimal128Array::from(vec![Some(1_234_i128), Some(0_i128)])
        .with_precision_and_scale(10, 3)?;

    let mut meta = HashMap::new();
    meta.insert("formula:column_type".to_owned(), "percentage".to_owned());
    meta.insert("formula:scale".to_owned(), "3".to_owned());

    let field = Field::new("pct", DataType::Decimal128(10, 3), true).with_metadata(meta);
    let schema = Arc::new(Schema::new(vec![field]));
    let batch = RecordBatch::try_new(schema, vec![Arc::new(arr) as ArrayRef])?;

    let table = record_batch_to_columnar(&batch)?;
    assert_eq!(table.schema()[0].column_type, ColumnType::Percentage { scale: 3 });
    assert_eq!(table.get_cell(0, 0), Value::Percentage(1_234));
    assert_eq!(table.get_cell(1, 0), Value::Percentage(0));
    Ok(())
}

#[test]
fn record_batch_to_columnar_accepts_more_integer_widths_as_number(
) -> Result<(), Box<dyn std::error::Error>> {
    use arrow_array::{Int64Array, Int8Array, UInt8Array};

    let i8_arr: ArrayRef = Arc::new(Int8Array::from(vec![Some(-1), None, Some(2)]));
    let u8_arr: ArrayRef = Arc::new(UInt8Array::from(vec![Some(0), Some(255), None]));
    let i64_arr: ArrayRef = Arc::new(Int64Array::from(vec![
        Some(-9_007_199_254_740_992_i64),
        Some(0_i64),
        None,
    ]));

    let schema = Arc::new(Schema::new(vec![
        Field::new("i8", DataType::Int8, true),
        Field::new("u8", DataType::UInt8, true),
        Field::new("i64", DataType::Int64, true),
    ]));
    let batch = RecordBatch::try_new(schema, vec![i8_arr, u8_arr, i64_arr])?;

    let table = record_batch_to_columnar(&batch)?;
    assert_eq!(table.schema()[0].column_type, ColumnType::Number);
    assert_eq!(table.schema()[1].column_type, ColumnType::Number);
    assert_eq!(table.schema()[2].column_type, ColumnType::Number);

    assert_eq!(table.get_cell(0, 0), Value::Number(-1.0));
    assert_eq!(table.get_cell(1, 0), Value::Null);
    assert_eq!(table.get_cell(2, 0), Value::Number(2.0));

    assert_eq!(table.get_cell(0, 1), Value::Number(0.0));
    assert_eq!(table.get_cell(1, 1), Value::Number(255.0));
    assert_eq!(table.get_cell(2, 1), Value::Null);

    assert_eq!(
        table.get_cell(0, 2),
        Value::Number(-9_007_199_254_740_992_i64 as f64)
    );
    assert_eq!(table.get_cell(1, 2), Value::Number(0.0));
    assert_eq!(table.get_cell(2, 2), Value::Null);

    Ok(())
}
