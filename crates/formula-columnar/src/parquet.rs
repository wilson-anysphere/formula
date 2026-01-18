#![forbid(unsafe_code)]

//! Parquet I/O for [`crate::ColumnarTable`].
//!
//! This module is behind the crate feature flag `arrow`.
//!
//! The implementation uses the Arrow ↔︎ `ColumnarTable` conversion in
//! [`crate::arrow`] and the `parquet` crate's Arrow integration.

use crate::arrow::{
    columnar_to_record_batch, column_type_from_field, value_from_array, ArrowInteropError,
};
use crate::table::{ColumnSchema, ColumnarTable, ColumnarTableBuilder, TableOptions};
use crate::types::Value;
use bytes::Bytes;
use parquet::arrow::arrow_reader::ParquetRecordBatchReaderBuilder;
use parquet::arrow::ArrowWriter;
use parquet::errors::ParquetError;
use parquet::file::reader::ChunkReader;
use parquet::file::properties::WriterProperties;
use std::fs::File;
use std::path::Path;

#[derive(Debug)]
pub enum ParquetInteropError {
    Io(std::io::Error),
    Arrow(ArrowInteropError),
    Parquet(ParquetError),
}

impl std::fmt::Display for ParquetInteropError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(err) => write!(f, "{err}"),
            Self::Arrow(err) => write!(f, "{err}"),
            Self::Parquet(err) => write!(f, "{err}"),
        }
    }
}

impl std::error::Error for ParquetInteropError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(err) => Some(err),
            Self::Arrow(err) => Some(err),
            Self::Parquet(err) => Some(err),
        }
    }
}

impl From<std::io::Error> for ParquetInteropError {
    fn from(value: std::io::Error) -> Self {
        Self::Io(value)
    }
}

impl From<ArrowInteropError> for ParquetInteropError {
    fn from(value: ArrowInteropError) -> Self {
        Self::Arrow(value)
    }
}

impl From<ParquetError> for ParquetInteropError {
    fn from(value: ParquetError) -> Self {
        Self::Parquet(value)
    }
}

fn table_schema_from_arrow(schema: &arrow_schema::Schema) -> Result<Vec<ColumnSchema>, ArrowInteropError> {
    schema
        .fields()
        .iter()
        .map(|field| {
            let column_type =
                column_type_from_field(field).map_err(|err| ArrowInteropError::Context {
                    context: format!("while parsing Arrow field {:?}", field.name()),
                    source: Box::new(err),
                })?;
            Ok(ColumnSchema {
                name: field.name().clone(),
                column_type,
            })
        })
        .collect()
}

fn append_record_batch(
    builder: &mut ColumnarTableBuilder,
    column_schema: &[ColumnSchema],
    batch: &arrow_array::RecordBatch,
) -> Result<(), ArrowInteropError> {
    let rows = batch.num_rows();
    let cols = batch.num_columns();

    for row in 0..rows {
        let mut values: Vec<Value> = Vec::new();
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
    Ok(())
}

/// Write a [`ColumnarTable`] to a Parquet file on disk.
pub fn write_columnar_to_parquet<P: AsRef<Path>>(
    table: &ColumnarTable,
    path: P,
) -> Result<(), ParquetInteropError> {
    let batch = columnar_to_record_batch(table)?;
    let file = File::create(path)?;

    let props = WriterProperties::builder()
        .set_dictionary_enabled(true)
        .build();
    let mut writer = ArrowWriter::try_new(file, batch.schema(), Some(props))?;
    writer.write(&batch)?;
    writer.close()?;
    Ok(())
}

/// Serialize a [`ColumnarTable`] into Parquet bytes.
pub fn write_columnar_to_parquet_bytes(table: &ColumnarTable) -> Result<Vec<u8>, ParquetInteropError> {
    let batch = columnar_to_record_batch(table)?;
    let props = WriterProperties::builder()
        .set_dictionary_enabled(true)
        .build();
    let mut out = Vec::new();
    {
        let mut writer = ArrowWriter::try_new(&mut out, batch.schema(), Some(props))?;
        writer.write(&batch)?;
        writer.close()?;
    }
    Ok(out)
}

fn read_parquet_reader_to_columnar<R: ChunkReader + 'static>(
    reader: R,
    options: TableOptions,
) -> Result<ColumnarTable, ParquetInteropError> {
    let builder = ParquetRecordBatchReaderBuilder::try_new(reader)?;
    let arrow_schema = builder.schema().clone();
    let column_schema = table_schema_from_arrow(&arrow_schema)?;

    let mut table_builder = ColumnarTableBuilder::new(column_schema.clone(), options);
    let mut reader = builder.build()?;
    while let Some(batch) = reader.next() {
        let batch = batch.map_err(ArrowInteropError::from)?;
        append_record_batch(&mut table_builder, &column_schema, &batch)?;
    }

    Ok(table_builder.finalize())
}

/// Read a Parquet file on disk into a [`ColumnarTable`] using [`TableOptions::default`].
pub fn read_parquet_to_columnar<P: AsRef<Path>>(path: P) -> Result<ColumnarTable, ParquetInteropError> {
    read_parquet_to_columnar_with_options(path, TableOptions::default())
}

/// Read a Parquet file on disk into a [`ColumnarTable`] using the provided [`TableOptions`].
pub fn read_parquet_to_columnar_with_options<P: AsRef<Path>>(
    path: P,
    options: TableOptions,
) -> Result<ColumnarTable, ParquetInteropError> {
    let file = File::open(path)?;
    read_parquet_reader_to_columnar(file, options)
}

/// Read Parquet bytes into a [`ColumnarTable`] using [`TableOptions::default`].
pub fn read_parquet_bytes_to_columnar(bytes: &[u8]) -> Result<ColumnarTable, ParquetInteropError> {
    read_parquet_bytes_to_columnar_with_options(bytes, TableOptions::default())
}

/// Read Parquet bytes into a [`ColumnarTable`] using the provided [`TableOptions`].
pub fn read_parquet_bytes_to_columnar_with_options(
    bytes: &[u8],
    options: TableOptions,
) -> Result<ColumnarTable, ParquetInteropError> {
    read_parquet_reader_to_columnar(Bytes::copy_from_slice(bytes), options)
}
