//! VertiPaq-style columnar storage for Formula.
//!
//! This crate focuses on:
//! - Columnar data representation with compression (dictionary + value encoding + optional RLE).
//! - Streaming ingestion (build chunks incrementally; never materialize a full cell grid).
//! - Virtualization-friendly access (`get_cell` / `get_range`) backed by an LRU of decoded chunks.
//! - Scan APIs for analytics-style operations.
//! - Incremental refresh workflows via [`MutableColumnarTable`] + `compact_in_place()` / `compact()` / `freeze()`.
//!
//! ## Incremental refresh
//!
//! ```no_run
//! use formula_columnar::{ColumnSchema, ColumnType, MutableColumnarTable, TableOptions, Value};
//! use std::sync::Arc;
//!
//! let schema = vec![
//!     ColumnSchema {
//!         name: "id".to_owned(),
//!         column_type: ColumnType::DateTime,
//!     },
//!     ColumnSchema {
//!         name: "category".to_owned(),
//!         column_type: ColumnType::String,
//!     },
//! ];
//!
//! let mut table = MutableColumnarTable::new(schema, TableOptions::default());
//! table.append_row(&[
//!     Value::DateTime(1),
//!     Value::String(Arc::<str>::from("A")),
//! ]);
//! table.append_row(&[
//!     Value::DateTime(2),
//!     Value::String(Arc::<str>::from("B")),
//! ]);
//!
//! // Apply a point update (stored as an overlay until compacted/frozen).
//! table.update_cell(1, 1, Value::String(Arc::<str>::from("C")));
//!
//! // Produce a compact immutable snapshot (suitable for scans / Data Model usage).
//! let snapshot = table.freeze();
//! assert_eq!(
//!     snapshot.get_cell(1, 1),
//!     Value::String(Arc::<str>::from("C"))
//! );
//! ```

#![forbid(unsafe_code)]

mod bitmap;
mod bitpacking;
mod cache;
mod encoding;
mod query;
mod stats;
mod table;
mod types;

#[cfg(feature = "arrow")]
pub mod arrow;
#[cfg(feature = "arrow")]
pub mod parquet;

// Persistence / storage layers sometimes need access to the encoded representation.
// We keep the implementation details in private modules, but re-export the relevant
// types so other workspace crates (e.g. `formula-storage`) can persist them without
// re-encoding to row-wise formats.
pub use crate::bitmap::BitVec;
pub use crate::cache::{CacheStats, PageCacheConfig};
pub use crate::encoding::{
    BoolChunk, DictionaryEncodedChunk, EncodedChunk, FloatChunk, RleEncodedU32, RleEncodedU64,
    U32SequenceEncoding, U64SequenceEncoding, ValueEncodedChunk,
};
pub use crate::query::{
    filter_mask, filter_table, group_by, group_by_mask, group_by_rows, hash_join, AggOp, AggSpec,
    CmpOp, FilterExpr, FilterValue, GroupByEngine, GroupByResult, JoinResult, QueryError,
};
pub use crate::stats::ColumnStats;
pub use crate::table::{
    ColumnAppendError, ColumnSchema, ColumnarRange, ColumnarTable, ColumnarTableBuilder,
    EncodedColumn, MutableColumnarTable, TableOptions, TableScan,
};
pub use crate::types::{ColumnType, Value};
