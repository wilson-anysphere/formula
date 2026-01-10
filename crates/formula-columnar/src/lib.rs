//! VertiPaq-style columnar storage for Formula.
//!
//! This crate focuses on:
//! - Columnar data representation with compression (dictionary + value encoding + optional RLE).
//! - Streaming ingestion (build chunks incrementally; never materialize a full cell grid).
//! - Virtualization-friendly access (`get_cell` / `get_range`) backed by an LRU of decoded chunks.
//! - Scan APIs for analytics-style operations.

#![forbid(unsafe_code)]

mod bitmap;
mod bitpacking;
mod cache;
mod encoding;
mod stats;
mod table;
mod types;

pub use crate::cache::{CacheStats, PageCacheConfig};
pub use crate::stats::ColumnStats;
pub use crate::table::{
    ColumnSchema, ColumnarRange, ColumnarTable, ColumnarTableBuilder, TableOptions, TableScan,
};
pub use crate::types::{ColumnType, Value};
