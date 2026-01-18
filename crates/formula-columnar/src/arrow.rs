#![forbid(unsafe_code)]

//! Arrow interoperability for `formula-columnar`.
//!
//! This module is behind the crate feature flag `arrow`.
//!
//! # Type mapping
//!
//! `formula-columnar` stores values using a small set of logical types
//! (`ColumnType`). When exporting to Arrow we map them as follows:
//!
//! | `ColumnType` | Arrow `DataType` | Notes |
//! |---|---|---|
//! | `Number` | `Float64` | |
//! | `Boolean` | `Boolean` | |
//! | `String` | `Dictionary(UInt32, Utf8)` | Preserves dictionary encoding. |
//! | `DateTime` | `Int64` | Stored as an integer with metadata. |
//! | `Currency { scale }` | `Int64` | Stored as an integer with metadata (`scale`). |
//! | `Percentage { scale }` | `Int64` | Stored as an integer with metadata (`scale`). |
//!
//! The `Int64` representation for `DateTime` / `Currency` / `Percentage` is a
//! deliberate choice: the in-memory engine already stores these values as
//! `i64`, and the exact units/semantics are defined by the higher-level model.
//! To allow round-tripping without loss, we attach column type information as
//! Arrow field metadata:
//!
//! - `formula:column_type` = one of `number`, `string`, `boolean`, `datetime`,
//!   `currency`, `percentage`
//! - `formula:scale` = `<u8>` for `currency` / `percentage`
//!
//! Consumers that do not understand these metadata keys will still see valid
//! Arrow arrays (e.g. `Int64`), but may not interpret them as specialized
//! logical types.
//!
//! ## Import compatibility
//!
//! While exports use a small fixed set of physical Arrow types (to keep the
//! Parquet representation stable and predictable), imports are intentionally
//! more permissive to handle real-world Arrow/Parquet schemas:
//!
//! - `ColumnType::Number` accepts `Float16/32/64`, `Int*/UInt*` (8/16/32/64),
//!   and `Decimal128`, upcasting to `f64`.
//! - `ColumnType::String` accepts `Utf8`, `LargeUtf8`, `Utf8View`, and
//!   dictionary-encoded strings with a variety of integer key types.

use crate::table::{ColumnSchema, ColumnarTable, ColumnarTableBuilder, TableOptions};
use crate::types::{ColumnType, Value};
use arrow_array::builder::{
    BooleanBuilder, Float64Builder, Int64Builder, StringDictionaryBuilder,
};
use arrow_array::{Array, ArrayRef, RecordBatch};
use arrow_schema::{DataType, Field, Schema, TimeUnit};
use std::collections::HashMap;
use std::sync::Arc;

const META_COLUMN_TYPE: &str = "formula:column_type";
const META_SCALE: &str = "formula:scale";

#[derive(Debug)]
pub enum ArrowInteropError {
    Arrow(arrow_schema::ArrowError),
    UnsupportedDataType(DataType),
    UnsupportedDictionaryValueType(DataType),
    InvalidDictionaryKey { key: String, dictionary_len: usize },
    InvalidDecimalScaleConversion {
        value: i128,
        from_scale: i32,
        to_scale: i32,
        reason: &'static str,
    },
    Context {
        context: String,
        source: Box<ArrowInteropError>,
    },
    InvalidMetadata { key: &'static str, value: String },
}

impl std::fmt::Display for ArrowInteropError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Arrow(err) => write!(f, "{err}"),
            Self::UnsupportedDataType(dt) => write!(f, "unsupported Arrow data type: {dt:?}"),
            Self::UnsupportedDictionaryValueType(dt) => {
                write!(f, "unsupported dictionary value type: {dt:?}")
            }
            Self::InvalidDictionaryKey { key, dictionary_len } => write!(
                f,
                "invalid dictionary key {key} (dictionary has {dictionary_len} values)"
            ),
            Self::InvalidDecimalScaleConversion {
                value,
                from_scale,
                to_scale,
                reason,
            } => write!(
                f,
                "cannot convert decimal {value} from scale {from_scale} to scale {to_scale}: {reason}"
            ),
            Self::Context { context, source } => write!(f, "{context}: {source}"),
            Self::InvalidMetadata { key, value } => {
                write!(f, "invalid Arrow field metadata {key}={value:?}")
            }
        }
    }
}

impl std::error::Error for ArrowInteropError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Arrow(err) => Some(err),
            Self::Context { source, .. } => Some(source),
            _ => None,
        }
    }
}

impl From<arrow_schema::ArrowError> for ArrowInteropError {
    fn from(value: arrow_schema::ArrowError) -> Self {
        Self::Arrow(value)
    }
}

fn scale_decimal_i128_to_i64(
    value: i128,
    from_scale: i32,
    to_scale: i32,
) -> Result<i64, ArrowInteropError> {
    let scale_diff = to_scale - from_scale;
    let mut out = value;

    if scale_diff > 0 {
        let pow10 = 10_i128.checked_pow(scale_diff as u32).ok_or_else(|| {
            ArrowInteropError::InvalidDecimalScaleConversion {
                value,
                from_scale,
                to_scale,
                reason: "scale conversion overflows i128",
            }
        })?;
        out = out.checked_mul(pow10).ok_or_else(|| ArrowInteropError::InvalidDecimalScaleConversion {
            value,
            from_scale,
            to_scale,
            reason: "scaled value overflows i128",
        })?;
    } else if scale_diff < 0 {
        let pow10 = 10_i128
            .checked_pow((-scale_diff) as u32)
            .ok_or_else(|| ArrowInteropError::InvalidDecimalScaleConversion {
                value,
                from_scale,
                to_scale,
                reason: "scale conversion overflows i128",
            })?;
        if out % pow10 != 0 {
            return Err(ArrowInteropError::InvalidDecimalScaleConversion {
                value,
                from_scale,
                to_scale,
                reason: "value has fractional digits beyond target scale",
            });
        }
        out /= pow10;
    }

    i64::try_from(out).map_err(|_| ArrowInteropError::InvalidDecimalScaleConversion {
        value,
        from_scale,
        to_scale,
        reason: "scaled value does not fit in i64",
    })
}

fn column_type_tag(column_type: ColumnType) -> &'static str {
    match column_type {
        ColumnType::Number => "number",
        ColumnType::String => "string",
        ColumnType::Boolean => "boolean",
        ColumnType::DateTime => "datetime",
        ColumnType::Currency { .. } => "currency",
        ColumnType::Percentage { .. } => "percentage",
    }
}

fn arrow_data_type_for_column_type(column_type: ColumnType) -> DataType {
    match column_type {
        ColumnType::Number => DataType::Float64,
        ColumnType::Boolean => DataType::Boolean,
        ColumnType::String => DataType::Dictionary(Box::new(DataType::UInt32), Box::new(DataType::Utf8)),
        ColumnType::DateTime | ColumnType::Currency { .. } | ColumnType::Percentage { .. } => {
            DataType::Int64
        }
    }
}

fn field_metadata(column_type: ColumnType) -> HashMap<String, String> {
    let mut meta = HashMap::new();
    meta.insert(META_COLUMN_TYPE.to_owned(), column_type_tag(column_type).to_owned());
    match column_type {
        ColumnType::Currency { scale } | ColumnType::Percentage { scale } => {
            meta.insert(META_SCALE.to_owned(), scale.to_string());
        }
        _ => {}
    }
    meta
}

fn arrow_field(schema: &ColumnSchema, nullable: bool) -> Field {
    Field::new(
        schema.name.clone(),
        arrow_data_type_for_column_type(schema.column_type),
        nullable,
    )
    .with_metadata(field_metadata(schema.column_type))
}

pub(crate) fn column_type_from_field(field: &Field) -> Result<ColumnType, ArrowInteropError> {
    // Prefer explicit metadata as it is required to disambiguate Int64 columns.
    let meta = field.metadata();
    if let Some(tag) = meta.get(META_COLUMN_TYPE) {
        let parsed = if tag.eq_ignore_ascii_case("number") {
            ColumnType::Number
        } else if tag.eq_ignore_ascii_case("string") {
            ColumnType::String
        } else if tag.eq_ignore_ascii_case("boolean") {
            ColumnType::Boolean
        } else if tag.eq_ignore_ascii_case("datetime") {
            ColumnType::DateTime
        } else if tag.eq_ignore_ascii_case("currency") {
                let scale = meta
                    .get(META_SCALE)
                    .and_then(|s| s.parse::<u8>().ok())
                    .unwrap_or(0);
            ColumnType::Currency { scale }
        } else if tag.eq_ignore_ascii_case("percentage") {
                let scale = meta
                    .get(META_SCALE)
                    .and_then(|s| s.parse::<u8>().ok())
                    .unwrap_or(0);
            ColumnType::Percentage { scale }
        } else {
            // Keep the error payload stable and user-friendly: report the metadata value in a
            // canonical ASCII-lowercased form.
            let mut value = tag.to_string();
            value.make_ascii_lowercase();
            return Err(ArrowInteropError::InvalidMetadata {
                key: META_COLUMN_TYPE,
                value,
            });
        };
        return Ok(parsed);
    }

    // Fall back to the physical Arrow data type.
    Ok(match field.data_type() {
        DataType::Float16
        | DataType::Float32
        | DataType::Float64
        | DataType::Decimal128(_, _) => ColumnType::Number,
        DataType::Boolean => ColumnType::Boolean,
        DataType::Utf8 | DataType::LargeUtf8 | DataType::Utf8View => ColumnType::String,
        // When a column is entirely null, Arrow may use the dedicated `Null` type.
        // The logical column type is ambiguous; default to String for compatibility
        // (values will still be imported as `Value::Null`).
        DataType::Null => ColumnType::String,
        DataType::Date32 | DataType::Date64 | DataType::Timestamp(_, _) => ColumnType::DateTime,
        DataType::Dictionary(_, value) => match value.as_ref() {
            DataType::Utf8 | DataType::LargeUtf8 | DataType::Utf8View => ColumnType::String,
            other => return Err(ArrowInteropError::UnsupportedDictionaryValueType(other.clone())),
        },
        DataType::Int8
        | DataType::UInt8
        | DataType::Int16
        | DataType::UInt16
        | DataType::Int32
        | DataType::UInt32
        | DataType::Int64
        | DataType::UInt64 => ColumnType::Number,
        other => return Err(ArrowInteropError::UnsupportedDataType(other.clone())),
    })
}

pub(crate) fn value_from_array(
    array: &dyn Array,
    row: usize,
    column_type: ColumnType,
) -> Result<Value, ArrowInteropError> {
    if array.is_null(row) {
        return Ok(Value::Null);
    }
    if array.data_type() == &DataType::Null {
        // `NullArray` is a special Arrow array where `null_count()` is 0 (no null
        // bitmap) but all values are logically null. Handle it explicitly so
        // we don't attempt to interpret it as any concrete physical type.
        return Ok(Value::Null);
    }

    match column_type {
        ColumnType::Number => match array.data_type() {
            DataType::Float16 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::Float16Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::Number(f32::from(arr.value(row)) as f64))
            }
            DataType::Decimal128(_, scale) => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::Decimal128Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                let v = arr.value(row);
                let scale = *scale as i32;
                let out = if scale >= 0 {
                    (v as f64) / 10_f64.powi(scale)
                } else {
                    (v as f64) * 10_f64.powi(-scale)
                };
                Ok(Value::Number(out))
            }
            DataType::Float32 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::Float32Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::Number(arr.value(row) as f64))
            }
            DataType::Float64 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::Float64Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::Number(arr.value(row)))
            }
            DataType::Int8 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::Int8Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::Number(arr.value(row) as f64))
            }
            DataType::UInt8 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::UInt8Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::Number(arr.value(row) as f64))
            }
            DataType::Int16 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::Int16Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::Number(arr.value(row) as f64))
            }
            DataType::UInt16 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::UInt16Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::Number(arr.value(row) as f64))
            }
            DataType::Int32 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::Int32Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::Number(arr.value(row) as f64))
            }
            DataType::UInt32 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::UInt32Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::Number(arr.value(row) as f64))
            }
            DataType::Int64 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::Int64Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::Number(arr.value(row) as f64))
            }
            DataType::UInt64 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::UInt64Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::Number(arr.value(row) as f64))
            }
            other => Err(ArrowInteropError::UnsupportedDataType(other.clone())),
        },
        ColumnType::Boolean => {
            let arr = array
                .as_any()
                .downcast_ref::<arrow_array::BooleanArray>()
                .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
            Ok(Value::Boolean(arr.value(row)))
        }
        ColumnType::String => match array.data_type() {
            DataType::Utf8 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::StringArray>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::String(Arc::<str>::from(arr.value(row))))
            }
            DataType::Utf8View => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::StringViewArray>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::String(Arc::<str>::from(arr.value(row))))
            }
            DataType::LargeUtf8 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::LargeStringArray>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::String(Arc::<str>::from(arr.value(row))))
            }
            DataType::Dictionary(key, value) => {
                macro_rules! dict_string_value {
                    ($key_ty:ty, $value_ty:ty) => {{
                        let dict = array
                            .as_any()
                            .downcast_ref::<arrow_array::DictionaryArray<$key_ty>>()
                            .ok_or_else(|| {
                                ArrowInteropError::UnsupportedDataType(array.data_type().clone())
                            })?;
                        let dictionary_len = dict.values().len();
                        let raw_key = dict.keys().value(row);
                        let key: usize = raw_key.try_into().map_err(|_| {
                            ArrowInteropError::InvalidDictionaryKey {
                                key: raw_key.to_string(),
                                dictionary_len,
                            }
                        })?;
                        if key >= dictionary_len {
                            Err(ArrowInteropError::InvalidDictionaryKey {
                                key: raw_key.to_string(),
                                dictionary_len,
                            })
                        } else {
                            let dict_values = dict
                                .values()
                                .as_any()
                                .downcast_ref::<$value_ty>()
                                .ok_or_else(|| {
                                    ArrowInteropError::UnsupportedDictionaryValueType(
                                        dict.values().data_type().clone(),
                                    )
                                })?;
                            if dict_values.is_null(key) {
                                Ok(Value::Null)
                            } else {
                                Ok(Value::String(Arc::<str>::from(dict_values.value(key))))
                            }
                        }
                    }};
                }

                match value.as_ref() {
                    DataType::Utf8 => match key.as_ref() {
                        DataType::UInt8 => dict_string_value!(
                            arrow_array::types::UInt8Type,
                            arrow_array::StringArray
                        ),
                        DataType::UInt16 => dict_string_value!(
                            arrow_array::types::UInt16Type,
                            arrow_array::StringArray
                        ),
                        DataType::UInt32 => dict_string_value!(
                            arrow_array::types::UInt32Type,
                            arrow_array::StringArray
                        ),
                        DataType::UInt64 => dict_string_value!(
                            arrow_array::types::UInt64Type,
                            arrow_array::StringArray
                        ),
                        DataType::Int8 => dict_string_value!(
                            arrow_array::types::Int8Type,
                            arrow_array::StringArray
                        ),
                        DataType::Int16 => dict_string_value!(
                            arrow_array::types::Int16Type,
                            arrow_array::StringArray
                        ),
                        DataType::Int32 => dict_string_value!(
                            arrow_array::types::Int32Type,
                            arrow_array::StringArray
                        ),
                        DataType::Int64 => dict_string_value!(
                            arrow_array::types::Int64Type,
                            arrow_array::StringArray
                        ),
                        other => Err(ArrowInteropError::UnsupportedDataType(other.clone())),
                    },
                    DataType::Utf8View => match key.as_ref() {
                        DataType::UInt8 => dict_string_value!(
                            arrow_array::types::UInt8Type,
                            arrow_array::StringViewArray
                        ),
                        DataType::UInt16 => dict_string_value!(
                            arrow_array::types::UInt16Type,
                            arrow_array::StringViewArray
                        ),
                        DataType::UInt32 => dict_string_value!(
                            arrow_array::types::UInt32Type,
                            arrow_array::StringViewArray
                        ),
                        DataType::UInt64 => dict_string_value!(
                            arrow_array::types::UInt64Type,
                            arrow_array::StringViewArray
                        ),
                        DataType::Int8 => dict_string_value!(
                            arrow_array::types::Int8Type,
                            arrow_array::StringViewArray
                        ),
                        DataType::Int16 => dict_string_value!(
                            arrow_array::types::Int16Type,
                            arrow_array::StringViewArray
                        ),
                        DataType::Int32 => dict_string_value!(
                            arrow_array::types::Int32Type,
                            arrow_array::StringViewArray
                        ),
                        DataType::Int64 => dict_string_value!(
                            arrow_array::types::Int64Type,
                            arrow_array::StringViewArray
                        ),
                        other => Err(ArrowInteropError::UnsupportedDataType(other.clone())),
                    },
                    DataType::LargeUtf8 => match key.as_ref() {
                        DataType::UInt8 => dict_string_value!(
                            arrow_array::types::UInt8Type,
                            arrow_array::LargeStringArray
                        ),
                        DataType::UInt16 => dict_string_value!(
                            arrow_array::types::UInt16Type,
                            arrow_array::LargeStringArray
                        ),
                        DataType::UInt32 => dict_string_value!(
                            arrow_array::types::UInt32Type,
                            arrow_array::LargeStringArray
                        ),
                        DataType::UInt64 => dict_string_value!(
                            arrow_array::types::UInt64Type,
                            arrow_array::LargeStringArray
                        ),
                        DataType::Int8 => dict_string_value!(
                            arrow_array::types::Int8Type,
                            arrow_array::LargeStringArray
                        ),
                        DataType::Int16 => dict_string_value!(
                            arrow_array::types::Int16Type,
                            arrow_array::LargeStringArray
                        ),
                        DataType::Int32 => dict_string_value!(
                            arrow_array::types::Int32Type,
                            arrow_array::LargeStringArray
                        ),
                        DataType::Int64 => dict_string_value!(
                            arrow_array::types::Int64Type,
                            arrow_array::LargeStringArray
                        ),
                        other => Err(ArrowInteropError::UnsupportedDataType(other.clone())),
                    },
                    other => Err(ArrowInteropError::UnsupportedDictionaryValueType(other.clone())),
                }
            }
            other => Err(ArrowInteropError::UnsupportedDataType(other.clone())),
        },
        ColumnType::DateTime => {
            match array.data_type() {
                DataType::Int64 => {
                    let arr = array
                        .as_any()
                        .downcast_ref::<arrow_array::Int64Array>()
                        .ok_or_else(|| {
                            ArrowInteropError::UnsupportedDataType(array.data_type().clone())
                        })?;
                    Ok(Value::DateTime(arr.value(row)))
                }
                DataType::Int32 => {
                    let arr = array
                        .as_any()
                        .downcast_ref::<arrow_array::Int32Array>()
                        .ok_or_else(|| {
                            ArrowInteropError::UnsupportedDataType(array.data_type().clone())
                        })?;
                    Ok(Value::DateTime(arr.value(row) as i64))
                }
                DataType::Date32 => {
                    let arr = array
                        .as_any()
                        .downcast_ref::<arrow_array::Date32Array>()
                        .ok_or_else(|| {
                            ArrowInteropError::UnsupportedDataType(array.data_type().clone())
                        })?;
                    Ok(Value::DateTime(arr.value(row) as i64))
                }
                DataType::Date64 => {
                    let arr = array
                        .as_any()
                        .downcast_ref::<arrow_array::Date64Array>()
                        .ok_or_else(|| {
                            ArrowInteropError::UnsupportedDataType(array.data_type().clone())
                        })?;
                    Ok(Value::DateTime(arr.value(row)))
                }
                DataType::Timestamp(unit, _) => match unit {
                    TimeUnit::Second => {
                        let arr = array
                            .as_any()
                            .downcast_ref::<arrow_array::TimestampSecondArray>()
                            .ok_or_else(|| {
                                ArrowInteropError::UnsupportedDataType(array.data_type().clone())
                            })?;
                        Ok(Value::DateTime(arr.value(row)))
                    }
                    TimeUnit::Millisecond => {
                        let arr = array
                            .as_any()
                            .downcast_ref::<arrow_array::TimestampMillisecondArray>()
                            .ok_or_else(|| {
                                ArrowInteropError::UnsupportedDataType(array.data_type().clone())
                            })?;
                        Ok(Value::DateTime(arr.value(row)))
                    }
                    TimeUnit::Microsecond => {
                        let arr = array
                            .as_any()
                            .downcast_ref::<arrow_array::TimestampMicrosecondArray>()
                            .ok_or_else(|| {
                                ArrowInteropError::UnsupportedDataType(array.data_type().clone())
                            })?;
                        Ok(Value::DateTime(arr.value(row)))
                    }
                    TimeUnit::Nanosecond => {
                        let arr = array
                            .as_any()
                            .downcast_ref::<arrow_array::TimestampNanosecondArray>()
                            .ok_or_else(|| {
                                ArrowInteropError::UnsupportedDataType(array.data_type().clone())
                            })?;
                        Ok(Value::DateTime(arr.value(row)))
                    }
                },
                other => Err(ArrowInteropError::UnsupportedDataType(other.clone())),
            }
        }
        ColumnType::Currency { scale } => match array.data_type() {
            DataType::Int64 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::Int64Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::Currency(arr.value(row)))
            }
            DataType::Decimal128(_, dec_scale) => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::Decimal128Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                let scaled = scale_decimal_i128_to_i64(
                    arr.value(row),
                    *dec_scale as i32,
                    scale as i32,
                )?;
                Ok(Value::Currency(scaled))
            }
            other => Err(ArrowInteropError::UnsupportedDataType(other.clone())),
        }
        ColumnType::Percentage { scale } => match array.data_type() {
            DataType::Int64 => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::Int64Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                Ok(Value::Percentage(arr.value(row)))
            }
            DataType::Decimal128(_, dec_scale) => {
                let arr = array
                    .as_any()
                    .downcast_ref::<arrow_array::Decimal128Array>()
                    .ok_or_else(|| ArrowInteropError::UnsupportedDataType(array.data_type().clone()))?;
                let scaled = scale_decimal_i128_to_i64(
                    arr.value(row),
                    *dec_scale as i32,
                    scale as i32,
                )?;
                Ok(Value::Percentage(scaled))
            }
            other => Err(ArrowInteropError::UnsupportedDataType(other.clone())),
        }
    }
}

fn array_from_column(
    table: &ColumnarTable,
    col: usize,
    column_schema: &ColumnSchema,
) -> Result<ArrayRef, ArrowInteropError> {
    let rows = table.row_count();

    let array: ArrayRef = match column_schema.column_type {
        ColumnType::Number => {
            let mut builder = Float64Builder::new();
            for row in 0..rows {
                match table.get_cell(row, col) {
                    Value::Number(v) => builder.append_value(v),
                    Value::Null => builder.append_null(),
                    _ => builder.append_null(),
                }
            }
            Arc::new(builder.finish())
        }
        ColumnType::Boolean => {
            let mut builder = BooleanBuilder::new();
            for row in 0..rows {
                match table.get_cell(row, col) {
                    Value::Boolean(v) => builder.append_value(v),
                    Value::Null => builder.append_null(),
                    _ => builder.append_null(),
                }
            }
            Arc::new(builder.finish())
        }
        ColumnType::String => {
            let mut builder = StringDictionaryBuilder::<arrow_array::types::UInt32Type>::new();
            for row in 0..rows {
                match table.get_cell(row, col) {
                    Value::String(v) => {
                        builder.append(v.as_ref())?;
                    }
                    Value::Null => builder.append_null(),
                    _ => builder.append_null(),
                }
            }
            Arc::new(builder.finish())
        }
        ColumnType::DateTime => {
            let mut builder = Int64Builder::new();
            for row in 0..rows {
                match table.get_cell(row, col) {
                    Value::DateTime(v) => builder.append_value(v),
                    Value::Null => builder.append_null(),
                    _ => builder.append_null(),
                }
            }
            Arc::new(builder.finish())
        }
        ColumnType::Currency { .. } => {
            let mut builder = Int64Builder::new();
            for row in 0..rows {
                match table.get_cell(row, col) {
                    Value::Currency(v) => builder.append_value(v),
                    Value::Null => builder.append_null(),
                    _ => builder.append_null(),
                }
            }
            Arc::new(builder.finish())
        }
        ColumnType::Percentage { .. } => {
            let mut builder = Int64Builder::new();
            for row in 0..rows {
                match table.get_cell(row, col) {
                    Value::Percentage(v) => builder.append_value(v),
                    Value::Null => builder.append_null(),
                    _ => builder.append_null(),
                }
            }
            Arc::new(builder.finish())
        }
    };

    Ok(array)
}

/// Convert a [`ColumnarTable`] into an Arrow [`RecordBatch`].
pub fn columnar_to_record_batch(table: &ColumnarTable) -> Result<RecordBatch, ArrowInteropError> {
    let col_count = table.column_count();
    let mut fields = Vec::new();
    let _ = fields.try_reserve_exact(col_count);
    let mut arrays = Vec::new();
    let _ = arrays.try_reserve_exact(col_count);

    for (col_idx, col_schema) in table.schema().iter().enumerate() {
        let nullable = table
            .scan()
            .stats(col_idx)
            .is_some_and(|stats| stats.null_count > 0);
        fields.push(arrow_field(col_schema, nullable));
        arrays.push(array_from_column(table, col_idx, col_schema)?);
    }

    let schema = Arc::new(Schema::new(fields));
    Ok(RecordBatch::try_new(schema, arrays)?)
}

/// Convert an Arrow [`RecordBatch`] into a [`ColumnarTable`] using [`TableOptions::default`].
pub fn record_batch_to_columnar(batch: &RecordBatch) -> Result<ColumnarTable, ArrowInteropError> {
    record_batch_to_columnar_with_options(batch, TableOptions::default())
}

/// Convert an Arrow [`RecordBatch`] into a [`ColumnarTable`] using the provided [`TableOptions`].
pub fn record_batch_to_columnar_with_options(
    batch: &RecordBatch,
    options: TableOptions,
) -> Result<ColumnarTable, ArrowInteropError> {
    let schema = batch.schema();
    let field_count = schema.fields().len();
    let mut column_schema = Vec::new();
    let _ = column_schema.try_reserve_exact(field_count);
    for field in schema.fields() {
        let column_type = column_type_from_field(field).map_err(|err| ArrowInteropError::Context {
            context: format!("while parsing Arrow field {:?}", field.name()),
            source: Box::new(err),
        })?;
        column_schema.push(ColumnSchema {
            name: field.name().clone(),
            column_type,
        });
    }

    let mut builder = ColumnarTableBuilder::new(column_schema.clone(), options);
    let rows = batch.num_rows();
    let cols = batch.num_columns();
    for row in 0..rows {
        let mut values = Vec::new();
        let _ = values.try_reserve_exact(cols);
        for col in 0..cols {
            let ty = column_schema[col].column_type;
            let array = batch.column(col).as_ref();
            values.push(value_from_array(array, row, ty).map_err(|err| {
                ArrowInteropError::Context {
                    context: format!(
                        "while reading Arrow column {:?} (row {row})",
                        column_schema[col].name
                    ),
                    source: Box::new(err),
                }
            })?);
        }
        builder.append_row(&values);
    }

    Ok(builder.finalize())
}
