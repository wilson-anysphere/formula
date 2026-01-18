#![forbid(unsafe_code)]

use crate::bitmap::BitVec;
use crate::cache::{CacheStats, LruCache, PageCacheConfig};
use crate::encoding::{
    BoolChunk, DecodedChunk, DictionaryEncodedChunk, EncodedChunk, FloatChunk, U32SequenceEncoding,
    U64SequenceEncoding, ValueEncodedChunk,
};
use crate::stats::{ColumnStats, DistinctCounter};
use crate::types::{ColumnType, Value};
use std::collections::HashMap;
use std::fmt;
use std::sync::{Arc, Mutex};

#[derive(Debug, Clone, Copy)]
pub struct TableOptions {
    pub page_size_rows: usize,
    pub cache: PageCacheConfig,
}

impl Default for TableOptions {
    fn default() -> Self {
        Self {
            page_size_rows: 65_536,
            cache: PageCacheConfig::default(),
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct ColumnSchema {
    pub name: String,
    pub column_type: ColumnType,
}

/// A fully-encoded column suitable for persistence.
///
/// This mirrors the internal column representation used by [`ColumnarTable`], but is exposed
/// as a public struct so other crates can construct a [`ColumnarTable`] without re-encoding
/// row-wise data.
#[derive(Clone, Debug)]
pub struct EncodedColumn {
    pub schema: ColumnSchema,
    pub chunks: Vec<EncodedChunk>,
    pub stats: ColumnStats,
    pub dictionary: Option<Arc<Vec<Arc<str>>>>,
}

#[derive(Clone, Debug)]
pub struct Column {
    schema: ColumnSchema,
    // `ColumnarTable` is frequently stored behind `Arc` and must be cloned when an `Arc` is not
    // uniquely owned (e.g. during calculated-column fallbacks). The encoded chunks can be large,
    // so keep them behind an `Arc` to make cloning cheap without changing the public
    // `EncodedChunk` / `EncodedColumn` API.
    chunks: Arc<Vec<EncodedChunk>>,
    stats: ColumnStats,
    dictionary: Option<Arc<Vec<Arc<str>>>>,
    distinct: Option<Arc<DistinctCounter>>,
}

impl Column {
    fn chunk_index(&self, row: usize, page_size: usize) -> (usize, usize) {
        (row / page_size, row % page_size)
    }

    fn get_cell(&self, row: usize, page_size: usize) -> Value {
        let (chunk_idx, in_chunk) = self.chunk_index(row, page_size);
        let Some(chunk) = self.chunks.get(chunk_idx) else {
            return Value::Null;
        };

        match (chunk, &self.dictionary, self.schema.column_type) {
            (EncodedChunk::Int(c), _, column_type) => c
                .get_i64(in_chunk)
                .map(|v| value_from_i64(column_type, v))
                .unwrap_or(Value::Null),
            (EncodedChunk::Float(c), _, _) => c
                .get_f64(in_chunk)
                .map(Value::Number)
                .unwrap_or(Value::Null),
            (EncodedChunk::Bool(c), _, _) => c
                .get_bool(in_chunk)
                .map(Value::Boolean)
                .unwrap_or(Value::Null),
            (EncodedChunk::Dict(c), Some(dict), _) => c
                .get_index(in_chunk)
                .and_then(|idx| dict.get(idx as usize).cloned())
                .map(Value::String)
                .unwrap_or(Value::Null),
            _ => Value::Null,
        }
    }

    fn decode_chunk(&self, chunk_idx: usize) -> Option<DecodedChunk> {
        let chunk = self.chunks.get(chunk_idx)?;
        match (chunk, &self.dictionary) {
            (EncodedChunk::Int(c), _) => Some(DecodedChunk::Int {
                values: c.decode_i64(),
                validity: c.validity.clone(),
            }),
            (EncodedChunk::Float(c), _) => Some(DecodedChunk::Float {
                values: c.values.clone(),
                validity: c.validity.clone(),
            }),
            (EncodedChunk::Bool(c), _) => Some(DecodedChunk::Bool {
                values: c.decode_bools(),
                validity: c.validity.clone(),
            }),
            (EncodedChunk::Dict(c), Some(dict)) => Some(DecodedChunk::Dict {
                indices: c.decode_indices(),
                validity: c.validity.clone(),
                dictionary: dict.clone(),
            }),
            _ => None,
        }
    }

    pub fn compressed_size_bytes(&self) -> usize {
        let dict_bytes = self
            .dictionary
            .as_ref()
            .map(|d| d.iter().map(|s| s.len()).sum::<usize>())
            .unwrap_or(0);
        let chunks_bytes: usize = self.chunks.iter().map(|c| c.compressed_size_bytes()).sum();
        dict_bytes + chunks_bytes
    }
}

#[derive(Clone, Debug)]
pub struct ColumnarTable {
    schema: Vec<ColumnSchema>,
    columns: Vec<Column>,
    rows: usize,
    options: TableOptions,
    cache: Arc<Mutex<LruCache<CacheKey, Arc<DecodedChunk>>>>,
}

/// Errors returned by [`ColumnarTable`] column append APIs.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ColumnAppendError {
    /// The provided values did not have exactly one entry per table row.
    LengthMismatch { expected: usize, actual: usize },
    /// A column with the provided name already exists in the table schema.
    DuplicateColumn { name: String },
    /// The encoded column chunks are not aligned to the table page size.
    PageAlignmentMismatch {
        chunk_index: usize,
        expected: usize,
        actual: usize,
    },
    /// Internal invariant violation.
    Internal { message: String },
}

impl fmt::Display for ColumnAppendError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::LengthMismatch { expected, actual } => write!(
                f,
                "column length mismatch: expected {} values (one per row), got {}",
                expected, actual
            ),
            Self::DuplicateColumn { name } => write!(f, "duplicate column name: {}", name),
            Self::PageAlignmentMismatch {
                chunk_index,
                expected,
                actual,
            } => write!(
                f,
                "encoded column chunk {} has length {} (expected {})",
                chunk_index, actual, expected
            ),
            Self::Internal { message } => write!(f, "internal error: {message}"),
        }
    }
}

impl std::error::Error for ColumnAppendError {}

#[derive(Clone, Debug, Hash, PartialEq, Eq)]
struct CacheKey {
    col: usize,
    chunk: usize,
}

fn value_from_i64(column_type: ColumnType, value: i64) -> Value {
    match column_type {
        ColumnType::DateTime => Value::DateTime(value),
        ColumnType::Currency { .. } => Value::Currency(value),
        ColumnType::Percentage { .. } => Value::Percentage(value),
        // Integer chunks can exist for other logical types in the future; default to a number.
        _ => Value::Number(value as f64),
    }
}

fn count_true_bits(data: &[u8], len: usize) -> u64 {
    let full_bytes = len / 8;
    let rem_bits = len % 8;
    let mut count: u64 = 0;
    for b in &data[..full_bytes] {
        count += b.count_ones() as u64;
    }
    if rem_bits > 0 {
        let mask = (1u8 << rem_bits) - 1;
        if let Some(last) = data.get(full_bytes) {
            count += (last & mask).count_ones() as u64;
        }
    }
    count
}

fn bitvec_all_true(len: usize) -> BitVec {
    BitVec::with_len_all_true(len)
}

impl ColumnarTable {
    pub fn schema(&self) -> &[ColumnSchema] {
        &self.schema
    }

    pub fn options(&self) -> TableOptions {
        self.options
    }

    pub fn row_count(&self) -> usize {
        self.rows
    }

    pub fn column_count(&self) -> usize {
        self.columns.len()
    }

    /// Return the dictionary backing a string column (if the column is dictionary encoded).
    pub fn dictionary(&self, col: usize) -> Option<Arc<Vec<Arc<str>>>> {
        self.columns.get(col)?.dictionary.clone()
    }

    pub fn stats(&self, col: usize) -> Option<&ColumnStats> {
        self.columns.get(col).map(|c| &c.stats)
    }

    /// Return the encoded chunks for a column, without decoding them.
    pub fn encoded_chunks(&self, col: usize) -> Option<&[EncodedChunk]> {
        Some(self.columns.get(col)?.chunks.as_slice())
    }

    /// Return a new [`ColumnarTable`] with `schema` and `values` appended as a new column.
    ///
    /// This is primarily intended for:
    /// - **Calculated columns**: query engines (e.g. `formula-dax`) can compute the column values
    ///   eagerly and then materialize them into a columnar snapshot.
    /// - **Persisted / incremental models**: a storage layer can load an existing encoded table and
    ///   append additional derived columns without rewriting the existing encoded pages.
    ///
    /// Notes:
    /// - Existing columns are *reused as-is*: they are not decoded, re-encoded, or rewritten.
    /// - Only the new column is encoded, using the normal [`ColumnarTableBuilder`] path.
    /// - The decoded-page cache is preserved so cached pages for existing columns remain valid.
    ///
    /// Returns [`ColumnAppendError`] if the column length does not match `row_count()` or if a
    /// column with the same name already exists.
    pub fn with_appended_column(
        mut self,
        schema: ColumnSchema,
        values: Vec<Value>,
    ) -> Result<Self, ColumnAppendError> {
        let expected = self.row_count();
        let actual = values.len();
        if actual != expected {
            return Err(ColumnAppendError::LengthMismatch { expected, actual });
        }

        let new_name = schema.name.as_str();
        if self
            .schema()
            .iter()
            .any(|existing| existing.name.as_str() == new_name)
        {
            return Err(ColumnAppendError::DuplicateColumn {
                name: schema.name,
            });
        }

        // Encode the new column using the existing builder logic, without touching the existing
        // columns or reinitializing the decoded-page cache.
        let mut builder = ColumnarTableBuilder::new(vec![schema.clone()], self.options());
        for value in &values {
            builder.append_row(std::slice::from_ref(value));
        }
        let mut encoded = builder.finalize();
        let Some(column) = encoded.columns.pop() else {
            return Err(ColumnAppendError::Internal {
                message: "append-column builder produced no columns".to_string(),
            });
        };

        self.schema.push(schema);
        self.columns.push(column);
        Ok(self)
    }

    /// Construct a [`ColumnarTable`] directly from encoded columns/chunks.
    ///
    /// This is intended for persistence layers that store the encoded chunks on disk.
    pub fn from_encoded(
        schema: Vec<ColumnSchema>,
        columns: Vec<EncodedColumn>,
        rows: usize,
        options: TableOptions,
    ) -> Self {
        let mut out_cols: Vec<Column> = Vec::new();
        let _ = out_cols.try_reserve_exact(columns.len());
        for col in columns {
            out_cols.push(Column {
                schema: col.schema,
                chunks: Arc::new(col.chunks),
                stats: col.stats,
                dictionary: col.dictionary,
                distinct: None,
            });
        }

        Self {
            schema,
            columns: out_cols,
            rows,
            options,
            cache: Arc::new(Mutex::new(LruCache::new(options.cache.max_entries))),
        }
    }

    /// Consume the table and return its encoded columns (chunks/stats/dictionary) without cloning.
    pub fn into_encoded_columns(self) -> Vec<EncodedColumn> {
        self.columns
            .into_iter()
            .map(|col| {
                let Column {
                    schema,
                    chunks,
                    mut stats,
                    dictionary,
                    distinct: _,
                } = col;
                let chunks = Arc::try_unwrap(chunks).unwrap_or_else(|shared| (*shared).clone());
                // Ensure the embedded stats remain consistent with the schema.
                stats.column_type = schema.column_type;
                EncodedColumn {
                    schema,
                    chunks,
                    stats,
                    dictionary,
                }
            })
            .collect()
    }

    /// Append a pre-encoded column (already chunked/encoded) to the table.
    ///
    /// This is useful when the caller already has an [`EncodedColumn`] payload and wants to build a
    /// new columnar snapshot without rewriting the existing encoded data. Common use cases include:
    /// - **Persisted / incremental models**: a storage layer can load encoded columns from disk and
    ///   append additional derived columns.
    /// - **Calculated column materialization**: a query engine can compute values, encode them into
    ///   chunks, and then append the encoded result to the underlying [`ColumnarTable`].
    ///
    /// Notes:
    /// - Existing columns are *reused as-is*: they are not decoded, re-encoded, or rewritten.
    /// - The decoded-page cache is preserved so cached pages for existing columns remain valid.
    /// - The appended column must match `row_count()` and respect the table's page size alignment
    ///   (see [`TableOptions`]).
    ///
    /// Returns [`ColumnAppendError`] if a column with the same name already exists, if the encoded
    /// length does not match `row_count()`, or if the chunks are not page-aligned.
    pub fn with_appended_encoded_column(
        mut self,
        mut column: EncodedColumn,
    ) -> Result<Self, ColumnAppendError> {
        let new_name = column.schema.name.as_str();
        if self
            .schema()
            .iter()
            .any(|existing| existing.name.as_str() == new_name)
        {
            return Err(ColumnAppendError::DuplicateColumn {
                name: column.schema.name,
            });
        }

        let expected_len = self.row_count();
        let actual_len: usize = column.chunks.iter().map(|c| c.len()).sum();
        if actual_len != expected_len {
            return Err(ColumnAppendError::LengthMismatch {
                expected: expected_len,
                actual: actual_len,
            });
        }

        let page_size = self.options.page_size_rows;
        if page_size > 0 && !column.chunks.is_empty() {
            let last_idx = column.chunks.len().saturating_sub(1);
            for (idx, chunk) in column.chunks.iter().enumerate() {
                let len = chunk.len();
                let is_last = idx == last_idx;
                if !is_last {
                    if len != page_size {
                        return Err(ColumnAppendError::PageAlignmentMismatch {
                            chunk_index: idx,
                            expected: page_size,
                            actual: len,
                        });
                    }
                } else {
                    // The last chunk may be smaller than the page size, but must be non-empty when
                    // the table itself has at least one row.
                    if len > page_size || (len == 0 && expected_len > 0) {
                        return Err(ColumnAppendError::PageAlignmentMismatch {
                            chunk_index: idx,
                            expected: page_size,
                            actual: len,
                        });
                    }
                }
            }
        }

        // Keep the stats column type consistent with the schema.
        column.stats.column_type = column.schema.column_type;

        let EncodedColumn {
            schema,
            chunks,
            stats,
            dictionary,
        } = column;

        // `ColumnarTable` stores a schema copy and an internal per-column schema.
        // Avoid re-encoding/cloning the encoded payload; the schema clone is small.
        self.schema.push(schema.clone());
        self.columns.push(Column {
            schema,
            chunks: Arc::new(chunks),
            stats,
            dictionary,
            distinct: None,
        });

        Ok(self)
    }

    pub fn get_cell(&self, row: usize, col: usize) -> Value {
        let Some(column) = self.columns.get(col) else {
            return Value::Null;
        };
        if row >= self.rows {
            return Value::Null;
        }
        column.get_cell(row, self.options.page_size_rows)
    }

    fn decoded_chunk_cached(&self, col: usize, chunk_idx: usize) -> Option<Arc<DecodedChunk>> {
        let key = CacheKey {
            col,
            chunk: chunk_idx,
        };

        let mut cache = self
            .cache
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if let Some(hit) = cache.get(&key) {
            return Some(hit);
        }

        drop(cache);
        let decoded = self.columns.get(col)?.decode_chunk(chunk_idx)?;
        let decoded = Arc::new(decoded);

        let mut cache = self
            .cache
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        cache.insert(key, decoded.clone());
        Some(decoded)
    }

    pub fn cache_stats(&self) -> CacheStats {
        self.cache
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .stats()
    }

    pub fn get_range(
        &self,
        row_start: usize,
        row_end: usize,
        col_start: usize,
        col_end: usize,
    ) -> ColumnarRange {
        let row_end = row_end.min(self.rows);
        let col_end = col_end.min(self.columns.len());
        let row_start = row_start.min(row_end);
        let col_start = col_start.min(col_end);

        let rows = row_end - row_start;
        let cols = col_end - col_start;

        let mut out_columns: Vec<Vec<Value>> = Vec::new();
        let _ = out_columns.try_reserve_exact(cols);
        for col in col_start..col_end {
            let mut values = Vec::new();
            let _ = values.try_reserve_exact(rows);
            let column_type = self.columns.get(col).map(|c| c.schema.column_type);

            let mut r = row_start;
            while r < row_end {
                let chunk_idx = r / self.options.page_size_rows;
                let chunk_row_start = chunk_idx * self.options.page_size_rows;
                let in_chunk_start = r - chunk_row_start;
                let remaining_in_chunk = self.options.page_size_rows - in_chunk_start;
                let take = (row_end - r).min(remaining_in_chunk);

                if let Some(decoded) = self.decoded_chunk_cached(col, chunk_idx) {
                    for i in 0..take {
                        let idx = in_chunk_start + i;
                        values.push(match column_type {
                            Some(ColumnType::Number) => decoded
                                .get_f64(idx)
                                .map(Value::Number)
                                .unwrap_or(Value::Null),
                            Some(ColumnType::String) => decoded
                                .get_string(idx)
                                .map(Value::String)
                                .unwrap_or(Value::Null),
                            Some(ColumnType::Boolean) => decoded
                                .get_bool(idx)
                                .map(Value::Boolean)
                                .unwrap_or(Value::Null),
                            Some(ColumnType::DateTime) => decoded
                                .get_i64(idx)
                                .map(Value::DateTime)
                                .unwrap_or(Value::Null),
                            Some(ColumnType::Currency { .. }) => decoded
                                .get_i64(idx)
                                .map(Value::Currency)
                                .unwrap_or(Value::Null),
                            Some(ColumnType::Percentage { .. }) => decoded
                                .get_i64(idx)
                                .map(Value::Percentage)
                                .unwrap_or(Value::Null),
                            None => Value::Null,
                        });
                    }
                } else {
                    // Fallback: should be unreachable for valid tables, but keep the API
                    // total and safe.
                    for _ in 0..take {
                        values.push(Value::Null);
                    }
                }

                r += take;
            }

            out_columns.push(values);
        }

        ColumnarRange {
            row_start,
            row_end,
            col_start,
            col_end,
            columns: out_columns,
        }
    }

    pub fn scan(&self) -> TableScan<'_> {
        TableScan { table: self }
    }

    pub fn compressed_size_bytes(&self) -> usize {
        self.columns.iter().map(|c| c.compressed_size_bytes()).sum()
    }

    pub(crate) fn page_size_rows(&self) -> usize {
        self.options.page_size_rows
    }

    /// Group rows by one or more key columns and compute aggregations.
    pub fn group_by(
        &self,
        keys: &[usize],
        aggs: &[crate::query::AggSpec],
    ) -> Result<crate::query::GroupByResult, crate::query::QueryError> {
        crate::query::group_by(self, keys, aggs)
    }

    pub fn group_by_rows(
        &self,
        keys: &[usize],
        aggs: &[crate::query::AggSpec],
        rows: &[usize],
    ) -> Result<crate::query::GroupByResult, crate::query::QueryError> {
        crate::query::group_by_rows(self, keys, aggs, rows)
    }

    pub fn group_by_mask(
        &self,
        keys: &[usize],
        aggs: &[crate::query::AggSpec],
        mask: &BitVec,
    ) -> Result<crate::query::GroupByResult, crate::query::QueryError> {
        crate::query::group_by_mask(self, keys, aggs, mask)
    }

    /// Evaluate a filter predicate and return a [`BitVec`] mask of matching rows.
    pub fn filter_mask(
        &self,
        expr: &crate::query::FilterExpr,
    ) -> Result<BitVec, crate::query::QueryError> {
        crate::query::filter_mask(self, expr)
    }

    pub fn filter_indices(
        &self,
        expr: &crate::query::FilterExpr,
    ) -> Result<Vec<usize>, crate::query::QueryError> {
        crate::query::filter_indices(self, expr)
    }

    /// Materialize a filtered table using a previously computed mask.
    pub fn filter_table(&self, mask: &BitVec) -> Result<ColumnarTable, crate::query::QueryError> {
        crate::query::filter_table(self, mask)
    }

    /// Convenience helper that evaluates a predicate and materializes the filtered table.
    pub fn filter(
        &self,
        expr: &crate::query::FilterExpr,
    ) -> Result<ColumnarTable, crate::query::QueryError> {
        let mask = crate::query::filter_mask(self, expr)?;
        crate::query::filter_table(self, &mask)
    }

    /// Hash join on a single key column.
    ///
    /// Returns row index mappings instead of materializing joined rows.
    pub fn hash_join(
        &self,
        right: &ColumnarTable,
        left_on: usize,
        right_on: usize,
    ) -> Result<crate::query::JoinResult, crate::query::QueryError> {
        crate::query::hash_join(self, right, left_on, right_on)
    }

    /// Hash join on a single key column (left join).
    pub fn hash_left_join(
        &self,
        right: &ColumnarTable,
        left_on: usize,
        right_on: usize,
    ) -> Result<crate::query::JoinResult<usize, Option<usize>>, crate::query::QueryError> {
        crate::query::hash_left_join(self, right, left_on, right_on)
    }

    /// Hash join on a single key column (right join).
    pub fn hash_right_join(
        &self,
        right: &ColumnarTable,
        left_on: usize,
        right_on: usize,
    ) -> Result<crate::query::JoinResult<Option<usize>, usize>, crate::query::QueryError> {
        crate::query::hash_right_join(self, right, left_on, right_on)
    }

    /// Hash join on a single key column (full outer join).
    pub fn hash_full_outer_join(
        &self,
        right: &ColumnarTable,
        left_on: usize,
        right_on: usize,
    ) -> Result<crate::query::JoinResult<Option<usize>, Option<usize>>, crate::query::QueryError> {
        crate::query::hash_full_outer_join(self, right, left_on, right_on)
    }

    /// Hash join on a single key column with a runtime join type.
    ///
    /// This is a convenience API that always returns optional indices, regardless of join type.
    pub fn hash_join_with_type(
        &self,
        right: &ColumnarTable,
        left_on: usize,
        right_on: usize,
        join_type: crate::query::JoinType,
    ) -> Result<crate::query::JoinResult<Option<usize>, Option<usize>>, crate::query::QueryError> {
        crate::query::hash_join_with_type(self, right, left_on, right_on, join_type)
    }

    /// Hash join on multiple key columns (inner join).
    pub fn hash_join_multi(
        &self,
        right: &ColumnarTable,
        left_keys: &[usize],
        right_keys: &[usize],
    ) -> Result<crate::query::JoinResult, crate::query::QueryError> {
        crate::query::hash_join_multi(self, right, left_keys, right_keys)
    }

    /// Hash join on multiple key columns (left join).
    ///
    /// Rows from the left table with no match (or NULL in any join key) are included with `None`
    /// for the right index.
    pub fn hash_left_join_multi(
        &self,
        right: &ColumnarTable,
        left_keys: &[usize],
        right_keys: &[usize],
    ) -> Result<crate::query::JoinResult<usize, Option<usize>>, crate::query::QueryError> {
        crate::query::hash_left_join_multi(self, right, left_keys, right_keys)
    }

    /// Hash join on multiple key columns (right join).
    pub fn hash_right_join_multi(
        &self,
        right: &ColumnarTable,
        left_keys: &[usize],
        right_keys: &[usize],
    ) -> Result<crate::query::JoinResult<Option<usize>, usize>, crate::query::QueryError> {
        crate::query::hash_right_join_multi(self, right, left_keys, right_keys)
    }

    /// Hash join on multiple key columns (full outer join).
    pub fn hash_full_outer_join_multi(
        &self,
        right: &ColumnarTable,
        left_keys: &[usize],
        right_keys: &[usize],
    ) -> Result<crate::query::JoinResult<Option<usize>, Option<usize>>, crate::query::QueryError> {
        crate::query::hash_full_outer_join_multi(self, right, left_keys, right_keys)
    }

    /// Hash join on multiple key columns with a runtime join type.
    ///
    /// This is a convenience API that always returns optional indices, regardless of join type.
    pub fn hash_join_multi_with_type(
        &self,
        right: &ColumnarTable,
        left_keys: &[usize],
        right_keys: &[usize],
        join_type: crate::query::JoinType,
    ) -> Result<crate::query::JoinResult<Option<usize>, Option<usize>>, crate::query::QueryError> {
        crate::query::hash_join_multi_with_type(self, right, left_keys, right_keys, join_type)
    }
}

/// A mutable, incrementally updatable columnar table.
///
/// This type supports:
/// - fast appends via page-sized chunk flushing (no rewriting of existing pages)
/// - sparse point/range updates via an in-memory overlay map
/// - `compact_in_place()` / `freeze()` to materialize overlays into the base pages
/// - `compact()` to produce a compact immutable [`ColumnarTable`] snapshot
///
/// The common Power Query refresh pattern is:
/// 1. Start with a mutable table (empty or derived from a previous snapshot)
/// 2. `append_rows` new data and apply any `update_*` fixes
/// 3. `compact()` or `freeze()` to hand a compact snapshot to the Data Model / query engine
#[derive(Debug)]
pub struct MutableColumnarTable {
    schema: Vec<ColumnSchema>,
    columns: Vec<MutableColumn>,
    rows: usize,
    options: TableOptions,
    cache: Arc<Mutex<LruCache<CacheKey, Arc<DecodedChunk>>>>,
    /// Sparse per-column overlays keyed by row index.
    ///
    /// When present, an overlay value takes precedence over the base encoded pages + the current
    /// append buffer. Overlays are cleared by `compact()`.
    overlays: Vec<HashMap<usize, Value>>,
}

#[derive(Debug)]
enum MutableColumn {
    Int(MutableIntColumn),
    Float(MutableFloatColumn),
    Bool(MutableBoolColumn),
    Dict(MutableDictColumn),
}

#[derive(Debug)]
struct MutableIntColumn {
    schema: ColumnSchema,
    page_size: usize,
    current: Vec<i64>,
    validity: BitVec,
    chunks: Vec<EncodedChunk>,
    distinct_base: u64,
    distinct: DistinctCounter,
    null_count: u64,
    min: Option<i64>,
    max: Option<i64>,
    sum: i128,
}

#[derive(Debug)]
struct MutableFloatColumn {
    schema: ColumnSchema,
    page_size: usize,
    current: Vec<f64>,
    validity: BitVec,
    chunks: Vec<EncodedChunk>,
    distinct_base: u64,
    distinct: DistinctCounter,
    null_count: u64,
    min: Option<f64>,
    max: Option<f64>,
    sum: f64,
}

#[derive(Debug)]
struct MutableBoolColumn {
    schema: ColumnSchema,
    page_size: usize,
    current: BitVec,
    validity: BitVec,
    chunks: Vec<EncodedChunk>,
    null_count: u64,
    true_count: u64,
}

#[derive(Debug)]
struct MutableDictColumn {
    schema: ColumnSchema,
    page_size: usize,
    dictionary: Arc<Vec<Arc<str>>>,
    dict_map: HashMap<Arc<str>, u32>,
    current: Vec<u32>,
    validity: BitVec,
    chunks: Vec<EncodedChunk>,
    null_count: u64,
    min: Option<Arc<str>>,
    max: Option<Arc<str>>,
    total_len: u64,
}

fn coerce_value_for_type(column_type: ColumnType, value: Value) -> Value {
    match (column_type, value) {
        (_, Value::Null) => Value::Null,
        (ColumnType::Number, Value::Number(v)) => Value::Number(v),
        (ColumnType::String, Value::String(s)) => Value::String(s),
        (ColumnType::Boolean, Value::Boolean(b)) => Value::Boolean(b),
        (ColumnType::DateTime, v)
        | (ColumnType::Currency { .. }, v)
        | (ColumnType::Percentage { .. }, v) => {
            // Int-backed logical types: accept any i64-like payload and coerce into the logical
            // column type for stable reads.
            let raw = match v {
                Value::DateTime(v) | Value::Currency(v) | Value::Percentage(v) => Some(v),
                Value::Number(v) => Some(v as i64),
                _ => None,
            };
            raw.map(|v| value_from_i64(column_type, v))
                .unwrap_or(Value::Null)
        }
        _ => Value::Null,
    }
}

fn i64_from_value(value: &Value) -> Option<i64> {
    match value {
        Value::DateTime(v) | Value::Currency(v) | Value::Percentage(v) => Some(*v),
        Value::Number(v) => Some(*v as i64),
        _ => None,
    }
}

impl MutableColumnarTable {
    pub fn new(schema: Vec<ColumnSchema>, options: TableOptions) -> Self {
        let columns = schema
            .iter()
            .cloned()
            .map(|col| match col.column_type {
                ColumnType::Number => MutableColumn::Float(MutableFloatColumn::new(
                    col,
                    options.page_size_rows,
                )),
                ColumnType::String => {
                    MutableColumn::Dict(MutableDictColumn::new(col, options.page_size_rows))
                }
                ColumnType::Boolean => {
                    MutableColumn::Bool(MutableBoolColumn::new(col, options.page_size_rows))
                }
                ColumnType::DateTime
                | ColumnType::Currency { .. }
                | ColumnType::Percentage { .. } => {
                    MutableColumn::Int(MutableIntColumn::new(col, options.page_size_rows))
                }
            })
            .collect::<Vec<_>>();

        let overlays = (0..schema.len()).map(|_| HashMap::new()).collect();

        Self {
            schema,
            columns,
            rows: 0,
            options,
            cache: Arc::new(Mutex::new(LruCache::new(options.cache.max_entries))),
            overlays,
        }
    }

    pub fn schema(&self) -> &[ColumnSchema] {
        &self.schema
    }

    pub fn row_count(&self) -> usize {
        self.rows
    }

    pub fn column_count(&self) -> usize {
        self.columns.len()
    }

    /// Return the dictionary backing a string column (if the column is dictionary encoded).
    pub fn dictionary(&self, col: usize) -> Option<Arc<Vec<Arc<str>>>> {
        match self.columns.get(col)? {
            MutableColumn::Dict(c) => Some(c.dictionary.clone()),
            _ => None,
        }
    }

    pub fn overlay_cell_count(&self) -> usize {
        self.overlays.iter().map(|m| m.len()).sum()
    }

    pub fn get_cell(&self, row: usize, col: usize) -> Value {
        if row >= self.rows {
            return Value::Null;
        }
        let Some(column) = self.columns.get(col) else {
            return Value::Null;
        };
        if let Some(overlay) = self.overlays.get(col).and_then(|m| m.get(&row)) {
            return overlay.clone();
        }
        column.get_cell(row, self.options.page_size_rows)
    }

    fn get_cell_base(&self, row: usize, col: usize) -> Value {
        if row >= self.rows {
            return Value::Null;
        }
        let Some(column) = self.columns.get(col) else {
            return Value::Null;
        };
        column.get_cell(row, self.options.page_size_rows)
    }

    fn decoded_chunk_cached(&self, col: usize, chunk_idx: usize) -> Option<Arc<DecodedChunk>> {
        let key = CacheKey {
            col,
            chunk: chunk_idx,
        };

        let mut cache = self
            .cache
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if let Some(hit) = cache.get(&key) {
            return Some(hit);
        }
        drop(cache);

        let decoded = self.columns.get(col)?.decode_chunk(chunk_idx)?;
        let decoded = Arc::new(decoded);

        let mut cache = self
            .cache
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        cache.insert(key, decoded.clone());
        Some(decoded)
    }

    pub fn cache_stats(&self) -> CacheStats {
        self.cache
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .stats()
    }

    pub fn column_stats(&self, col: usize) -> Option<ColumnStats> {
        self.columns.get(col).map(|c| c.stats())
    }

    pub fn append_row(&mut self, row: &[Value]) {
        assert_eq!(
            row.len(),
            self.columns.len(),
            "row length must match schema"
        );

        for col_idx in 0..self.columns.len() {
            self.maybe_clear_cache_for_dict_growth(col_idx, &row[col_idx]);
            self.columns[col_idx].push(&row[col_idx]);
        }

        self.rows += 1;
        if self.rows % self.options.page_size_rows == 0 {
            for column in &mut self.columns {
                column.flush();
            }
        }
    }

    pub fn append_rows<I, R>(&mut self, rows: I)
    where
        I: IntoIterator<Item = R>,
        R: AsRef<[Value]>,
    {
        for row in rows {
            self.append_row(row.as_ref());
        }
    }

    pub fn update_cell(&mut self, row: usize, col: usize, value: Value) -> bool {
        let Some(recompute) = self.update_cell_core(row, col, value) else {
            return false;
        };

        if recompute {
            self.recompute_min_max(col);
        }

        true
    }

    pub fn update_range(
        &mut self,
        row_start: usize,
        row_end: usize,
        col_start: usize,
        col_end: usize,
        values: &[Value],
    ) -> usize {
        let row_end = row_end.min(self.rows);
        let col_end = col_end.min(self.columns.len());
        let row_start = row_start.min(row_end);
        let col_start = col_start.min(col_end);

        let rows = row_end - row_start;
        let cols = col_end - col_start;
        let expected = rows.saturating_mul(cols);
        assert_eq!(
            values.len(),
            expected,
            "values must be row-major with len == rows*cols"
        );

        if expected == 0 {
            return 0;
        }

        let mut needs_recompute: Vec<bool> = vec![false; self.columns.len()];
        for r in 0..rows {
            for c in 0..cols {
                let idx = r * cols + c;
                if let Some(recompute) = self.update_cell_core(
                    row_start + r,
                    col_start + c,
                    values[idx].clone(),
                ) {
                    needs_recompute[col_start + c] |= recompute;
                }
            }
        }

        for col in col_start..col_end {
            if needs_recompute[col] {
                self.recompute_min_max(col);
            }
        }

        expected
    }

    fn update_cell_core(&mut self, row: usize, col: usize, value: Value) -> Option<bool> {
        if row >= self.rows {
            return None;
        }
        let column_type = self.columns.get(col)?.column_type();

        let coerced = coerce_value_for_type(column_type, value);
        let base = self.get_cell_base(row, col);
        let old = self.get_cell(row, col);

        if old == coerced {
            return Some(false);
        }

        if coerced == base {
            self.overlays.get_mut(col)?.remove(&row);
        } else {
            self.overlays.get_mut(col)?.insert(row, coerced.clone());
        }

        self.maybe_clear_cache_for_dict_growth(col, &coerced);
        let recompute = self.columns[col].apply_update(&old, &coerced);
        Some(recompute)
    }

    /// Delete rows in `[row_start, row_end)` (0-based, half-open).
    ///
    /// This operation currently rebuilds the table (similar to `compact()`), because deletions are
    /// expected to be rare compared to appends/updates in Power Query refresh flows.
    ///
    /// Returns the number of rows deleted.
    pub fn delete_rows(&mut self, row_start: usize, row_end: usize) -> usize {
        let row_end = row_end.min(self.rows);
        let row_start = row_start.min(row_end);
        let delete_count = row_end - row_start;
        if delete_count == 0 {
            return 0;
        }

        let mut rebuilt = MutableColumnarTable::new(self.schema.clone(), self.options);
        let mut row_buf: Vec<Value> = Vec::new();
        let _ = row_buf.try_reserve_exact(self.columns.len());
        for row in 0..self.rows {
            if row >= row_start && row < row_end {
                continue;
            }
            row_buf.clear();
            for col in 0..self.columns.len() {
                row_buf.push(self.get_cell(row, col));
            }
            rebuilt.append_row(&row_buf);
        }
        *self = rebuilt;

        delete_count
    }

    fn flush_all(&mut self) {
        for column in &mut self.columns {
            column.flush();
        }
    }

    fn maybe_clear_cache_for_dict_growth(&mut self, col: usize, value: &Value) {
        let Value::String(s) = value else {
            return;
        };

        let (dict_ref_count, contains) = match self.columns.get(col) {
            Some(MutableColumn::Dict(c)) => {
                (Arc::strong_count(&c.dictionary), c.dict_map.contains_key(s.as_ref()))
            }
            _ => return,
        };

        if contains {
            return;
        }

        // If the dictionary is currently shared (e.g. referenced by cached decoded pages), growing
        // it would trigger an `Arc::make_mut` clone of the entire dictionary Vec. To keep string
        // appends cheap, clear the page cache first so the dictionary is likely uniquely owned.
        if dict_ref_count > 1 {
            self.cache
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner())
                .remove_if(|key| key.col == col);
        }
    }

    fn recompute_min_max(&mut self, col: usize) {
        if col >= self.columns.len() {
            return;
        }

        let column_type = match self.columns.get(col) {
            Some(c) => c.column_type(),
            None => return,
        };

        match column_type {
            ColumnType::Number => {
                let mut min: Option<f64> = None;
                let mut max: Option<f64> = None;
                for row in 0..self.rows {
                    if let Value::Number(v) = self.get_cell(row, col) {
                        min = Some(min.map(|m| m.min(v)).unwrap_or(v));
                        max = Some(max.map(|m| m.max(v)).unwrap_or(v));
                    }
                }
                if let MutableColumn::Float(c) = &mut self.columns[col] {
                    c.min = min;
                    c.max = max;
                }
            }
            ColumnType::String => {
                let mut min: Option<Arc<str>> = None;
                let mut max: Option<Arc<str>> = None;
                for row in 0..self.rows {
                    if let Value::String(s) = self.get_cell(row, col) {
                        min = match &min {
                            Some(m) if m.as_ref() <= s.as_ref() => Some(m.clone()),
                            _ => Some(s.clone()),
                        };
                        max = match &max {
                            Some(m) if m.as_ref() >= s.as_ref() => Some(m.clone()),
                            _ => Some(s.clone()),
                        };
                    }
                }
                if let MutableColumn::Dict(c) = &mut self.columns[col] {
                    c.min = min;
                    c.max = max;
                }
            }
            ColumnType::Boolean => {}
            ColumnType::DateTime | ColumnType::Currency { .. } | ColumnType::Percentage { .. } => {
                let mut min: Option<i64> = None;
                let mut max: Option<i64> = None;
                for row in 0..self.rows {
                    if let Some(v) = i64_from_value(&self.get_cell(row, col)) {
                        min = Some(min.map(|m| m.min(v)).unwrap_or(v));
                        max = Some(max.map(|m| m.max(v)).unwrap_or(v));
                    }
                }
                if let MutableColumn::Int(c) = &mut self.columns[col] {
                    c.min = min;
                    c.max = max;
                }
            }
        }
    }

    /// Compact overlay state into the underlying encoded pages.
    ///
    /// This rewrites only the pages that contain overlay edits and clears the overlay maps.
    /// Appended (unflushed) rows are updated in-place in their per-column append buffers.
    pub fn compact_in_place(&mut self) {
        if self.overlay_cell_count() == 0 {
            return;
        }

        let page = self.options.page_size_rows;
        for col in 0..self.columns.len() {
            if self.overlays[col].is_empty() {
                continue;
            }

            let overlay_map = std::mem::take(&mut self.overlays[col]);
            let mut by_chunk: HashMap<usize, Vec<(usize, Value)>> = HashMap::new();
            for (row, value) in overlay_map {
                let chunk_idx = row / page;
                let in_chunk = row % page;
                by_chunk.entry(chunk_idx).or_default().push((in_chunk, value));
            }

            let chunk_count = self.columns[col].chunk_count();
            for (chunk_idx, updates) in by_chunk {
                if chunk_idx < chunk_count {
                    self.columns[col].apply_overlays_to_chunk(chunk_idx, &updates);
                    let key = CacheKey { col, chunk: chunk_idx };
                    self.cache
                        .lock()
                        .unwrap_or_else(|poisoned| poisoned.into_inner())
                        .remove(&key);
                } else if chunk_idx == chunk_count {
                    self.columns[col].apply_overlays_to_current(&updates);
                }
            }
        }
    }

    /// Produce a compact immutable snapshot of the table.
    ///
    /// This merges any overlay updates into the base pages (like [`Self::compact_in_place`]) and
    /// returns an immutable [`ColumnarTable`] containing all current rows (including the current
    /// tail append buffer).
    ///
    /// The mutable table remains appendable; its tail buffer is intentionally *not* flushed into
    /// `chunks`, because that would break the fixed `page_size_rows` page alignment used for
    /// subsequent appends.
    pub fn compact(&mut self) -> ColumnarTable {
        self.compact_in_place();
        self.to_columnar_table()
    }

    /// Consume the mutable table and return a compact immutable [`ColumnarTable`].
    pub fn freeze(mut self) -> ColumnarTable {
        self.compact_in_place();
        self.flush_all();
        self.into_columnar_table()
    }

    fn to_columnar_table(&self) -> ColumnarTable {
        let columns = self
            .columns
            .iter()
            .map(|c| c.as_column_snapshot())
            .collect();
        ColumnarTable {
            schema: self.schema.clone(),
            columns,
            rows: self.rows,
            options: self.options,
            cache: Arc::new(Mutex::new(LruCache::new(self.options.cache.max_entries))),
        }
    }

    fn into_columnar_table(self) -> ColumnarTable {
        let columns = self.columns.into_iter().map(|c| c.into_column()).collect();
        ColumnarTable {
            schema: self.schema,
            columns,
            rows: self.rows,
            options: self.options,
            cache: Arc::new(Mutex::new(LruCache::new(self.options.cache.max_entries))),
        }
    }

    pub fn get_range(
        &self,
        row_start: usize,
        row_end: usize,
        col_start: usize,
        col_end: usize,
    ) -> ColumnarRange {
        let row_end = row_end.min(self.rows);
        let col_end = col_end.min(self.columns.len());
        let row_start = row_start.min(row_end);
        let col_start = col_start.min(col_end);

        let rows = row_end - row_start;
        let cols = col_end - col_start;

        let mut out_columns: Vec<Vec<Value>> = Vec::new();
        let _ = out_columns.try_reserve_exact(cols);
        for col in col_start..col_end {
            let mut values = Vec::new();
            let _ = values.try_reserve_exact(rows);
            let column_type = self.columns.get(col).map(|c| c.column_type());
            let overlay = self.overlays.get(col);
            let has_overlay = overlay.is_some_and(|m| !m.is_empty());

            let mut r = row_start;
            while r < row_end {
                let chunk_idx = r / self.options.page_size_rows;
                let chunk_row_start = chunk_idx * self.options.page_size_rows;
                let in_chunk_start = r - chunk_row_start;
                let remaining_in_chunk = self.options.page_size_rows - in_chunk_start;
                let take = (row_end - r).min(remaining_in_chunk);

                if chunk_idx < self.columns[col].chunk_count() {
                    if let Some(decoded) = self.decoded_chunk_cached(col, chunk_idx) {
                        for i in 0..take {
                            let row_idx = r + i;
                            if has_overlay {
                                if let Some(v) = overlay.and_then(|m| m.get(&row_idx)) {
                                    values.push(v.clone());
                                    continue;
                                }
                            }
                            let idx = in_chunk_start + i;
                            values.push(match column_type {
                                Some(ColumnType::Number) => decoded
                                    .get_f64(idx)
                                    .map(Value::Number)
                                    .unwrap_or(Value::Null),
                                Some(ColumnType::String) => decoded
                                    .get_string(idx)
                                    .map(Value::String)
                                    .unwrap_or(Value::Null),
                                Some(ColumnType::Boolean) => decoded
                                    .get_bool(idx)
                                    .map(Value::Boolean)
                                    .unwrap_or(Value::Null),
                                Some(ColumnType::DateTime) => decoded
                                    .get_i64(idx)
                                    .map(Value::DateTime)
                                    .unwrap_or(Value::Null),
                                Some(ColumnType::Currency { .. }) => decoded
                                    .get_i64(idx)
                                    .map(Value::Currency)
                                    .unwrap_or(Value::Null),
                                Some(ColumnType::Percentage { .. }) => decoded
                                    .get_i64(idx)
                                    .map(Value::Percentage)
                                    .unwrap_or(Value::Null),
                                None => Value::Null,
                            });
                        }
                    } else {
                        for i in 0..take {
                            let row_idx = r + i;
                            if has_overlay {
                                if let Some(v) = overlay.and_then(|m| m.get(&row_idx)) {
                                    values.push(v.clone());
                                    continue;
                                }
                            }
                            values.push(Value::Null);
                        }
                    }
                } else {
                    for i in 0..take {
                        let row_idx = r + i;
                        if has_overlay {
                            if let Some(v) = overlay.and_then(|m| m.get(&row_idx)) {
                                values.push(v.clone());
                                continue;
                            }
                        }
                        values.push(self.columns[col].get_cell(row_idx, self.options.page_size_rows));
                    }
                }

                r += take;
            }

            out_columns.push(values);
        }

        ColumnarRange {
            row_start,
            row_end,
            col_start,
            col_end,
            columns: out_columns,
        }
    }
}

impl ColumnarTable {
    pub fn into_mutable(self) -> MutableColumnarTable {
        MutableColumnarTable::from(self)
    }
}

impl From<ColumnarTable> for MutableColumnarTable {
    fn from(table: ColumnarTable) -> Self {
        let overlays = (0..table.columns.len()).map(|_| HashMap::new()).collect();
        let page_size = table.options.page_size_rows;
        let remainder = if page_size == 0 {
            0
        } else {
            table.rows % page_size
        };

        let mut columns: Vec<MutableColumn> = table
            .columns
            .into_iter()
            .map(|col| MutableColumn::from_column(col, page_size))
            .collect();

        if remainder != 0 {
            for column in &mut columns {
                column.take_tail_chunk_into_current(remainder);
            }
        }

        Self {
            schema: table.schema,
            columns,
            rows: table.rows,
            options: table.options,
            cache: Arc::new(Mutex::new(LruCache::new(table.options.cache.max_entries))),
            overlays,
        }
    }
}

impl MutableColumn {
    fn column_type(&self) -> ColumnType {
        match self {
            MutableColumn::Int(c) => c.schema.column_type,
            MutableColumn::Float(c) => c.schema.column_type,
            MutableColumn::Bool(c) => c.schema.column_type,
            MutableColumn::Dict(c) => c.schema.column_type,
        }
    }

    fn chunk_count(&self) -> usize {
        match self {
            MutableColumn::Int(c) => c.chunks.len(),
            MutableColumn::Float(c) => c.chunks.len(),
            MutableColumn::Bool(c) => c.chunks.len(),
            MutableColumn::Dict(c) => c.chunks.len(),
        }
    }

    fn get_cell(&self, row: usize, page_size: usize) -> Value {
        match self {
            MutableColumn::Int(c) => c.get_cell(row, page_size),
            MutableColumn::Float(c) => c.get_cell(row, page_size),
            MutableColumn::Bool(c) => c.get_cell(row, page_size),
            MutableColumn::Dict(c) => c.get_cell(row, page_size),
        }
    }

    fn decode_chunk(&self, chunk_idx: usize) -> Option<DecodedChunk> {
        match self {
            MutableColumn::Int(c) => c.decode_chunk(chunk_idx),
            MutableColumn::Float(c) => c.decode_chunk(chunk_idx),
            MutableColumn::Bool(c) => c.decode_chunk(chunk_idx),
            MutableColumn::Dict(c) => c.decode_chunk(chunk_idx),
        }
    }

    fn push(&mut self, value: &Value) {
        match self {
            MutableColumn::Int(c) => c.push(value),
            MutableColumn::Float(c) => c.push(value),
            MutableColumn::Bool(c) => c.push(value),
            MutableColumn::Dict(c) => c.push(value),
        }
    }

    fn flush(&mut self) {
        match self {
            MutableColumn::Int(c) => c.flush(),
            MutableColumn::Float(c) => c.flush(),
            MutableColumn::Bool(c) => c.flush(),
            MutableColumn::Dict(c) => c.flush(),
        }
    }

    fn stats(&self) -> ColumnStats {
        match self {
            MutableColumn::Int(c) => c.stats(),
            MutableColumn::Float(c) => c.stats(),
            MutableColumn::Bool(c) => c.stats(),
            MutableColumn::Dict(c) => c.stats(),
        }
    }

    fn apply_update(&mut self, old: &Value, new: &Value) -> bool {
        match self {
            MutableColumn::Int(c) => c.apply_update(old, new),
            MutableColumn::Float(c) => c.apply_update(old, new),
            MutableColumn::Bool(c) => c.apply_update(old, new),
            MutableColumn::Dict(c) => c.apply_update(old, new),
        }
    }

    fn apply_overlays_to_chunk(&mut self, chunk_idx: usize, updates: &[(usize, Value)]) {
        match self {
            MutableColumn::Int(c) => c.apply_overlays_to_chunk(chunk_idx, updates),
            MutableColumn::Float(c) => c.apply_overlays_to_chunk(chunk_idx, updates),
            MutableColumn::Bool(c) => c.apply_overlays_to_chunk(chunk_idx, updates),
            MutableColumn::Dict(c) => c.apply_overlays_to_chunk(chunk_idx, updates),
        }
    }

    fn apply_overlays_to_current(&mut self, updates: &[(usize, Value)]) {
        match self {
            MutableColumn::Int(c) => c.apply_overlays_to_current(updates),
            MutableColumn::Float(c) => c.apply_overlays_to_current(updates),
            MutableColumn::Bool(c) => c.apply_overlays_to_current(updates),
            MutableColumn::Dict(c) => c.apply_overlays_to_current(updates),
        }
    }

    fn take_tail_chunk_into_current(&mut self, expected_len: usize) {
        if expected_len == 0 {
            return;
        }

        match self {
            MutableColumn::Int(c) => c.take_tail_chunk_into_current(expected_len),
            MutableColumn::Float(c) => c.take_tail_chunk_into_current(expected_len),
            MutableColumn::Bool(c) => c.take_tail_chunk_into_current(expected_len),
            MutableColumn::Dict(c) => c.take_tail_chunk_into_current(expected_len),
        }
    }

    fn as_column_snapshot(&self) -> Column {
        match self {
            MutableColumn::Int(c) => c.as_column_snapshot(),
            MutableColumn::Float(c) => c.as_column_snapshot(),
            MutableColumn::Bool(c) => c.as_column_snapshot(),
            MutableColumn::Dict(c) => c.as_column_snapshot(),
        }
    }

    fn into_column(self) -> Column {
        match self {
            MutableColumn::Int(c) => c.into_column(),
            MutableColumn::Float(c) => c.into_column(),
            MutableColumn::Bool(c) => c.into_column(),
            MutableColumn::Dict(c) => c.into_column(),
        }
    }

    fn from_column(col: Column, page_size: usize) -> Self {
        match col.schema.column_type {
            ColumnType::Number => MutableColumn::Float(MutableFloatColumn::from_column(col, page_size)),
            ColumnType::String => MutableColumn::Dict(MutableDictColumn::from_column(col, page_size)),
            ColumnType::Boolean => MutableColumn::Bool(MutableBoolColumn::from_column(col, page_size)),
            ColumnType::DateTime
            | ColumnType::Currency { .. }
            | ColumnType::Percentage { .. } => {
                MutableColumn::Int(MutableIntColumn::from_column(col, page_size))
            }
        }
    }
}

impl MutableIntColumn {
    fn new(schema: ColumnSchema, page_size: usize) -> Self {
        let mut current = Vec::new();
        let _ = current.try_reserve_exact(page_size);
        Self {
            schema,
            page_size,
            current,
            validity: BitVec::with_capacity_bits(page_size),
            chunks: Vec::new(),
            distinct_base: 0,
            distinct: DistinctCounter::new(),
            null_count: 0,
            min: None,
            max: None,
            sum: 0,
        }
    }

    fn from_column(col: Column, page_size: usize) -> Self {
        let Column {
            schema,
            chunks,
            stats,
            dictionary: _,
            distinct,
        } = col;

        let chunks = match Arc::try_unwrap(chunks) {
            Ok(chunks) => chunks,
            Err(chunks) => (*chunks).clone(),
        };

        let (distinct_base, distinct) = match distinct {
            Some(counter) => {
                let counter = match Arc::try_unwrap(counter) {
                    Ok(counter) => counter,
                    Err(counter) => (*counter).clone(),
                };
                (0, counter)
            }
            None => {
                // We do not have the base distinct sketch (e.g. the table came from
                // `ColumnarTable::from_encoded`). Reconstruct it from the encoded pages so
                // subsequent appends/updates can maintain a consistent estimate.
                let mut counter = DistinctCounter::new();
                for chunk in &chunks {
                    let EncodedChunk::Int(c) = chunk else {
                        continue;
                    };
                    let values = c.decode_i64();
                    if let Some(validity) = &c.validity {
                        for (idx, v) in values.iter().enumerate() {
                            if validity.get(idx) {
                                counter.insert_i64(*v);
                            }
                        }
                    } else {
                        for v in values {
                            counter.insert_i64(v);
                        }
                    }
                }
                (0, counter)
            }
        };

        let min = stats.min.as_ref().and_then(i64_from_value);
        let max = stats.max.as_ref().and_then(i64_from_value);
        let sum = stats.sum.unwrap_or(0.0) as i128;
        let null_count = stats.null_count;

        let mut current = Vec::new();
        let _ = current.try_reserve_exact(page_size);
        Self {
            schema,
            page_size,
            current,
            validity: BitVec::with_capacity_bits(page_size),
            chunks,
            distinct_base,
            distinct,
            null_count,
            min,
            max,
            sum,
        }
    }

    fn chunk_index(&self, row: usize, page_size: usize) -> (usize, usize) {
        (row / page_size, row % page_size)
    }

    fn get_cell(&self, row: usize, page_size: usize) -> Value {
        let (chunk_idx, in_chunk) = self.chunk_index(row, page_size);
        if let Some(chunk) = self.chunks.get(chunk_idx) {
            let EncodedChunk::Int(c) = chunk else {
                return Value::Null;
            };
            return c
                .get_i64(in_chunk)
                .map(|v| value_from_i64(self.schema.column_type, v))
                .unwrap_or(Value::Null);
        }

        if chunk_idx == self.chunks.len() {
            if in_chunk < self.current.len() {
                if self.validity.get(in_chunk) {
                    return value_from_i64(self.schema.column_type, self.current[in_chunk]);
                }
            }
        }

        Value::Null
    }

    fn decode_chunk(&self, chunk_idx: usize) -> Option<DecodedChunk> {
        let chunk = self.chunks.get(chunk_idx)?;
        let EncodedChunk::Int(c) = chunk else {
            return None;
        };
        Some(DecodedChunk::Int {
            values: c.decode_i64(),
            validity: c.validity.clone(),
        })
    }

    fn push(&mut self, value: &Value) {
        let pushed = match value {
            Value::Null => {
                self.current.push(0);
                self.validity.push(false);
                self.null_count += 1;
                return;
            }
            Value::DateTime(v) | Value::Currency(v) | Value::Percentage(v) => Some(*v),
            Value::Number(v) => Some(*v as i64),
            _ => None,
        };

        match pushed {
            Some(v) => {
                self.current.push(v);
                self.validity.push(true);
                self.distinct.insert_i64(v);
                self.sum += v as i128;
                self.min = Some(self.min.map(|m| m.min(v)).unwrap_or(v));
                self.max = Some(self.max.map(|m| m.max(v)).unwrap_or(v));
            }
            None => {
                self.current.push(0);
                self.validity.push(false);
                self.null_count += 1;
            }
        }
    }

    fn flush(&mut self) {
        if self.current.is_empty() {
            return;
        }

        let mut min_valid: Option<i64> = None;
        for (idx, v) in self.current.iter().enumerate() {
            if self.validity.get(idx) {
                min_valid = Some(min_valid.map(|m| m.min(*v)).unwrap_or(*v));
            }
        }
        let min = min_valid.unwrap_or(0);
        let offsets: Vec<u64> = self
            .current
            .iter()
            .enumerate()
            .map(|(idx, v)| {
                if self.validity.get(idx) {
                    (*v as i128 - min as i128) as u64
                } else {
                    0
                }
            })
            .collect();
        let offsets = U64SequenceEncoding::encode(&offsets);

        let validity = if self.validity.all_true() {
            None
        } else {
            Some(self.validity.clone())
        };

        self.chunks.push(EncodedChunk::Int(ValueEncodedChunk {
            min,
            len: self.current.len(),
            offsets,
            validity,
        }));

        self.current.clear();
        self.validity = BitVec::with_capacity_bits(self.page_size);
    }

    fn stats(&self) -> ColumnStats {
        ColumnStats {
            column_type: self.schema.column_type,
            distinct_count: self.distinct_base.saturating_add(self.distinct.estimate()),
            null_count: self.null_count,
            min: self.min.map(|v| value_from_i64(self.schema.column_type, v)),
            max: self.max.map(|v| value_from_i64(self.schema.column_type, v)),
            sum: Some(self.sum as f64),
            avg_length: None,
        }
    }

    fn apply_update(&mut self, old: &Value, new: &Value) -> bool {
        let old_i = i64_from_value(old);
        let new_i = i64_from_value(new);

        match (old_i, new_i) {
            (None, None) => return false,
            (None, Some(v)) => {
                self.null_count = self.null_count.saturating_sub(1);
                self.sum += v as i128;
                self.distinct.insert_i64(v);
                self.min = Some(self.min.map(|m| m.min(v)).unwrap_or(v));
                self.max = Some(self.max.map(|m| m.max(v)).unwrap_or(v));
                return false;
            }
            (Some(v), None) => {
                self.null_count += 1;
                self.sum -= v as i128;
                let needs_recompute = self.min == Some(v) || self.max == Some(v);
                return needs_recompute;
            }
            (Some(old_v), Some(new_v)) => {
                self.sum += new_v as i128 - old_v as i128;
                self.distinct.insert_i64(new_v);
                self.min = Some(self.min.map(|m| m.min(new_v)).unwrap_or(new_v));
                self.max = Some(self.max.map(|m| m.max(new_v)).unwrap_or(new_v));
                let needs_recompute =
                    (self.min == Some(old_v) && new_v != old_v) || (self.max == Some(old_v) && new_v != old_v);
                return needs_recompute;
            }
        }
    }

    fn apply_overlays_to_chunk(&mut self, chunk_idx: usize, updates: &[(usize, Value)]) {
        let (len, mut values, mut validity) = match self.chunks.get(chunk_idx) {
            Some(EncodedChunk::Int(c)) => (c.len, c.decode_i64(), c.validity.clone()),
            _ => return,
        };

        for (in_chunk, value) in updates {
            if *in_chunk >= len {
                continue;
            }
            if let Some(v) = i64_from_value(value) {
                values[*in_chunk] = v;
                if let Some(validity) = &mut validity {
                    validity.set(*in_chunk, true);
                }
            } else {
                values[*in_chunk] = 0;
                let validity = validity.get_or_insert_with(|| bitvec_all_true(len));
                validity.set(*in_chunk, false);
            }
        }

        let mut min_valid: Option<i64> = None;
        for (idx, v) in values.iter().enumerate() {
            let is_valid = validity.as_ref().map(|b| b.get(idx)).unwrap_or(true);
            if is_valid {
                min_valid = Some(min_valid.map(|m| m.min(*v)).unwrap_or(*v));
            }
        }
        let min = min_valid.unwrap_or(0);

        let offsets: Vec<u64> = values
            .iter()
            .enumerate()
            .map(|(idx, v)| {
                let is_valid = validity.as_ref().map(|b| b.get(idx)).unwrap_or(true);
                if is_valid {
                    (*v as i128 - min as i128) as u64
                } else {
                    0
                }
            })
            .collect();

        let offsets = U64SequenceEncoding::encode(&offsets);
        let validity = validity.and_then(|v| (!v.all_true()).then_some(v));

        self.chunks[chunk_idx] = EncodedChunk::Int(ValueEncodedChunk {
            min,
            len,
            offsets,
            validity,
        });
    }

    fn apply_overlays_to_current(&mut self, updates: &[(usize, Value)]) {
        for (in_chunk, value) in updates {
            if *in_chunk >= self.current.len() {
                continue;
            }
            if let Some(v) = i64_from_value(value) {
                self.current[*in_chunk] = v;
                self.validity.set(*in_chunk, true);
            } else {
                self.current[*in_chunk] = 0;
                self.validity.set(*in_chunk, false);
            }
        }
    }

    fn take_tail_chunk_into_current(&mut self, expected_len: usize) {
        let Some(last) = self.chunks.last() else {
            return;
        };
        if last.len() != expected_len {
            return;
        }

        let Some(EncodedChunk::Int(chunk)) = self.chunks.pop() else {
            return;
        };

        let values = chunk.decode_i64();
        let validity = chunk
            .validity
            .unwrap_or_else(|| BitVec::with_len_all_true(expected_len));
        self.current = values;
        self.validity = validity;
    }

    fn encoded_current_chunk(&self) -> Option<EncodedChunk> {
        if self.current.is_empty() {
            return None;
        }

        let len = self.current.len();
        let mut min_valid: Option<i64> = None;
        for (idx, v) in self.current.iter().enumerate() {
            if self.validity.get(idx) {
                min_valid = Some(min_valid.map(|m| m.min(*v)).unwrap_or(*v));
            }
        }
        let min = min_valid.unwrap_or(0);

        let offsets: Vec<u64> = self
            .current
            .iter()
            .enumerate()
            .map(|(idx, v)| {
                if self.validity.get(idx) {
                    (*v as i128 - min as i128) as u64
                } else {
                    0
                }
            })
            .collect();
        let offsets = U64SequenceEncoding::encode(&offsets);

        let validity = if self.validity.all_true() {
            None
        } else {
            Some(self.validity.clone())
        };

        Some(EncodedChunk::Int(ValueEncodedChunk {
            min,
            len,
            offsets,
            validity,
        }))
    }

    fn as_column_snapshot(&self) -> Column {
        let mut chunks = self.chunks.clone();
        if let Some(chunk) = self.encoded_current_chunk() {
            chunks.push(chunk);
        }
        let distinct = (self.distinct_base == 0).then(|| Arc::new(self.distinct.clone()));
        Column {
            schema: self.schema.clone(),
            chunks: Arc::new(chunks),
            stats: self.stats(),
            dictionary: None,
            distinct,
        }
    }

    fn into_column(self) -> Column {
        let stats = self.stats();
        let distinct = (self.distinct_base == 0).then(|| Arc::new(self.distinct));
        Column {
            schema: self.schema,
            chunks: Arc::new(self.chunks),
            stats,
            dictionary: None,
            distinct,
        }
    }
}

impl MutableFloatColumn {
    fn new(schema: ColumnSchema, page_size: usize) -> Self {
        let mut current = Vec::new();
        let _ = current.try_reserve_exact(page_size);
        Self {
            schema,
            page_size,
            current,
            validity: BitVec::with_capacity_bits(page_size),
            chunks: Vec::new(),
            distinct_base: 0,
            distinct: DistinctCounter::new(),
            null_count: 0,
            min: None,
            max: None,
            sum: 0.0,
        }
    }

    fn from_column(col: Column, page_size: usize) -> Self {
        let Column {
            schema,
            chunks,
            stats,
            dictionary: _,
            distinct,
        } = col;

        let chunks = match Arc::try_unwrap(chunks) {
            Ok(chunks) => chunks,
            Err(chunks) => (*chunks).clone(),
        };

        let (distinct_base, distinct) = match distinct {
            Some(counter) => {
                let counter = match Arc::try_unwrap(counter) {
                    Ok(counter) => counter,
                    Err(counter) => (*counter).clone(),
                };
                (0, counter)
            }
            None => {
                let mut counter = DistinctCounter::new();
                for chunk in &chunks {
                    let EncodedChunk::Float(c) = chunk else {
                        continue;
                    };
                    if let Some(validity) = &c.validity {
                        for (idx, v) in c.values.iter().enumerate() {
                            if validity.get(idx) {
                                counter.insert_i64(canonical_f64_bits(*v) as i64);
                            }
                        }
                    } else {
                        for v in &c.values {
                            counter.insert_i64(canonical_f64_bits(*v) as i64);
                        }
                    }
                }
                (0, counter)
            }
        };

        let min = match &stats.min {
            Some(Value::Number(v)) => Some(*v),
            _ => None,
        };
        let max = match &stats.max {
            Some(Value::Number(v)) => Some(*v),
            _ => None,
        };
        let sum = stats.sum.unwrap_or(0.0);

        let mut current = Vec::new();
        let _ = current.try_reserve_exact(page_size);
        Self {
            schema,
            page_size,
            current,
            validity: BitVec::with_capacity_bits(page_size),
            chunks,
            distinct_base,
            distinct,
            null_count: stats.null_count,
            min,
            max,
            sum,
        }
    }

    fn chunk_index(&self, row: usize, page_size: usize) -> (usize, usize) {
        (row / page_size, row % page_size)
    }

    fn get_cell(&self, row: usize, page_size: usize) -> Value {
        let (chunk_idx, in_chunk) = self.chunk_index(row, page_size);
        if let Some(chunk) = self.chunks.get(chunk_idx) {
            let EncodedChunk::Float(c) = chunk else {
                return Value::Null;
            };
            return c
                .get_f64(in_chunk)
                .map(Value::Number)
                .unwrap_or(Value::Null);
        }

        if chunk_idx == self.chunks.len() {
            if in_chunk < self.current.len() {
                if self.validity.get(in_chunk) {
                    return Value::Number(self.current[in_chunk]);
                }
            }
        }

        Value::Null
    }

    fn decode_chunk(&self, chunk_idx: usize) -> Option<DecodedChunk> {
        let chunk = self.chunks.get(chunk_idx)?;
        let EncodedChunk::Float(c) = chunk else {
            return None;
        };
        Some(DecodedChunk::Float {
            values: c.values.clone(),
            validity: c.validity.clone(),
        })
    }

    fn push(&mut self, value: &Value) {
        match value {
            Value::Null => {
                self.current.push(0.0);
                self.validity.push(false);
                self.null_count += 1;
            }
            Value::Number(v) => {
                self.current.push(*v);
                self.validity.push(true);
                self.distinct.insert_i64(canonical_f64_bits(*v) as i64);
                self.sum += *v;
                self.min = Some(self.min.map(|m| m.min(*v)).unwrap_or(*v));
                self.max = Some(self.max.map(|m| m.max(*v)).unwrap_or(*v));
            }
            _ => {
                self.current.push(0.0);
                self.validity.push(false);
                self.null_count += 1;
            }
        }
    }

    fn flush(&mut self) {
        if self.current.is_empty() {
            return;
        }

        let validity = if self.validity.all_true() {
            None
        } else {
            Some(self.validity.clone())
        };

        self.chunks.push(EncodedChunk::Float(FloatChunk {
            values: std::mem::take(&mut self.current),
            validity,
        }));

        self.validity = BitVec::with_capacity_bits(self.page_size);
    }

    fn stats(&self) -> ColumnStats {
        ColumnStats {
            column_type: self.schema.column_type,
            distinct_count: self.distinct_base.saturating_add(self.distinct.estimate()),
            null_count: self.null_count,
            min: self.min.map(Value::Number),
            max: self.max.map(Value::Number),
            sum: Some(self.sum),
            avg_length: None,
        }
    }

    fn apply_update(&mut self, old: &Value, new: &Value) -> bool {
        let old_v = match old {
            Value::Number(v) => Some(*v),
            _ => None,
        };
        let new_v = match new {
            Value::Number(v) => Some(*v),
            _ => None,
        };

        match (old_v, new_v) {
            (None, None) => false,
            (None, Some(v)) => {
                self.null_count = self.null_count.saturating_sub(1);
                self.sum += v;
                self.distinct.insert_i64(canonical_f64_bits(v) as i64);
                self.min = Some(self.min.map(|m| m.min(v)).unwrap_or(v));
                self.max = Some(self.max.map(|m| m.max(v)).unwrap_or(v));
                false
            }
            (Some(v), None) => {
                self.null_count += 1;
                self.sum -= v;
                self.min == Some(v) || self.max == Some(v)
            }
            (Some(old_v), Some(new_v)) => {
                self.sum += new_v - old_v;
                self.distinct.insert_i64(canonical_f64_bits(new_v) as i64);
                self.min = Some(self.min.map(|m| m.min(new_v)).unwrap_or(new_v));
                self.max = Some(self.max.map(|m| m.max(new_v)).unwrap_or(new_v));
                (self.min == Some(old_v) && new_v != old_v) || (self.max == Some(old_v) && new_v != old_v)
            }
        }
    }

    fn apply_overlays_to_chunk(&mut self, chunk_idx: usize, updates: &[(usize, Value)]) {
        let Some(chunk) = self.chunks.get_mut(chunk_idx) else {
            return;
        };
        let EncodedChunk::Float(chunk) = chunk else {
            return;
        };

        let len = chunk.values.len();
        for (in_chunk, value) in updates {
            if *in_chunk >= len {
                continue;
            }
            if let Value::Number(v) = value {
                chunk.values[*in_chunk] = *v;
                if let Some(validity) = &mut chunk.validity {
                    validity.set(*in_chunk, true);
                }
            } else {
                chunk.values[*in_chunk] = 0.0;
                let validity = chunk.validity.get_or_insert_with(|| bitvec_all_true(len));
                validity.set(*in_chunk, false);
            }
        }

        if chunk.validity.as_ref().is_some_and(|v| v.all_true()) {
            chunk.validity = None;
        }
    }

    fn apply_overlays_to_current(&mut self, updates: &[(usize, Value)]) {
        for (in_chunk, value) in updates {
            if *in_chunk >= self.current.len() {
                continue;
            }
            if let Value::Number(v) = value {
                self.current[*in_chunk] = *v;
                self.validity.set(*in_chunk, true);
            } else {
                self.current[*in_chunk] = 0.0;
                self.validity.set(*in_chunk, false);
            }
        }
    }

    fn take_tail_chunk_into_current(&mut self, expected_len: usize) {
        let Some(last) = self.chunks.last() else {
            return;
        };
        if last.len() != expected_len {
            return;
        }

        let Some(EncodedChunk::Float(chunk)) = self.chunks.pop() else {
            return;
        };

        let validity = chunk
            .validity
            .unwrap_or_else(|| BitVec::with_len_all_true(expected_len));
        self.current = chunk.values;
        self.validity = validity;
    }

    fn encoded_current_chunk(&self) -> Option<EncodedChunk> {
        if self.current.is_empty() {
            return None;
        }

        let validity = if self.validity.all_true() {
            None
        } else {
            Some(self.validity.clone())
        };

        Some(EncodedChunk::Float(FloatChunk {
            values: self.current.clone(),
            validity,
        }))
    }

    fn as_column_snapshot(&self) -> Column {
        let mut chunks = self.chunks.clone();
        if let Some(chunk) = self.encoded_current_chunk() {
            chunks.push(chunk);
        }
        let distinct = (self.distinct_base == 0).then(|| Arc::new(self.distinct.clone()));
        Column {
            schema: self.schema.clone(),
            chunks: Arc::new(chunks),
            stats: self.stats(),
            dictionary: None,
            distinct,
        }
    }

    fn into_column(self) -> Column {
        let stats = self.stats();
        let distinct = (self.distinct_base == 0).then(|| Arc::new(self.distinct));
        Column {
            schema: self.schema,
            chunks: Arc::new(self.chunks),
            stats,
            dictionary: None,
            distinct,
        }
    }
}

impl MutableBoolColumn {
    fn new(schema: ColumnSchema, page_size: usize) -> Self {
        Self {
            schema,
            page_size,
            current: BitVec::with_capacity_bits(page_size),
            validity: BitVec::with_capacity_bits(page_size),
            chunks: Vec::new(),
            null_count: 0,
            true_count: 0,
        }
    }

    fn from_column(col: Column, page_size: usize) -> Self {
        let Column {
            schema,
            chunks,
            stats,
            dictionary: _,
            distinct: _,
        } = col;

        let chunks = match Arc::try_unwrap(chunks) {
            Ok(chunks) => chunks,
            Err(chunks) => (*chunks).clone(),
        };

        let true_count = stats.sum.unwrap_or(0.0).round().max(0.0) as u64;
        Self {
            schema,
            page_size,
            current: BitVec::with_capacity_bits(page_size),
            validity: BitVec::with_capacity_bits(page_size),
            chunks,
            null_count: stats.null_count,
            true_count,
        }
    }

    fn chunk_index(&self, row: usize, page_size: usize) -> (usize, usize) {
        (row / page_size, row % page_size)
    }

    fn get_cell(&self, row: usize, page_size: usize) -> Value {
        let (chunk_idx, in_chunk) = self.chunk_index(row, page_size);
        if let Some(chunk) = self.chunks.get(chunk_idx) {
            let EncodedChunk::Bool(c) = chunk else {
                return Value::Null;
            };
            return c
                .get_bool(in_chunk)
                .map(Value::Boolean)
                .unwrap_or(Value::Null);
        }

        if chunk_idx == self.chunks.len() {
            if in_chunk < self.current.len() {
                if self.validity.get(in_chunk) {
                    return Value::Boolean(self.current.get(in_chunk));
                }
            }
        }

        Value::Null
    }

    fn decode_chunk(&self, chunk_idx: usize) -> Option<DecodedChunk> {
        let chunk = self.chunks.get(chunk_idx)?;
        let EncodedChunk::Bool(c) = chunk else {
            return None;
        };
        Some(DecodedChunk::Bool {
            values: c.decode_bools(),
            validity: c.validity.clone(),
        })
    }

    fn push(&mut self, value: &Value) {
        match value {
            Value::Null => {
                self.current.push(false);
                self.validity.push(false);
                self.null_count += 1;
            }
            Value::Boolean(v) => {
                self.current.push(*v);
                self.validity.push(true);
                if *v {
                    self.true_count += 1;
                }
            }
            _ => {
                self.current.push(false);
                self.validity.push(false);
                self.null_count += 1;
            }
        }
    }

    fn flush(&mut self) {
        if self.current.is_empty() {
            return;
        }

        let len = self.current.len();
        let mut data = vec![0u8; (len + 7) / 8];
        for i in 0..len {
            if self.current.get(i) {
                data[i / 8] |= 1u8 << (i % 8);
            }
        }

        let validity = if self.validity.all_true() {
            None
        } else {
            Some(self.validity.clone())
        };

        self.chunks.push(EncodedChunk::Bool(BoolChunk {
            len,
            data,
            validity,
        }));

        self.current = BitVec::with_capacity_bits(self.page_size);
        self.validity = BitVec::with_capacity_bits(self.page_size);
    }

    fn stats(&self) -> ColumnStats {
        let non_null = (self
            .chunks
            .iter()
            .map(|c| c.len() as u64)
            .sum::<u64>()
            + self.current.len() as u64)
        .saturating_sub(self.null_count);
        let false_count = non_null.saturating_sub(self.true_count);
        let distinct_count = match (self.true_count > 0, false_count > 0) {
            (false, false) => 0,
            (true, true) => 2,
            _ => 1,
        };

        ColumnStats {
            column_type: self.schema.column_type,
            distinct_count,
            null_count: self.null_count,
            min: None,
            max: None,
            sum: Some(self.true_count as f64),
            avg_length: None,
        }
    }

    fn apply_update(&mut self, old: &Value, new: &Value) -> bool {
        match (old, new) {
            (Value::Null, Value::Null) => {}
            (Value::Null, Value::Boolean(v)) => {
                self.null_count = self.null_count.saturating_sub(1);
                if *v {
                    self.true_count += 1;
                }
            }
            (Value::Boolean(v), Value::Null) => {
                self.null_count += 1;
                if *v {
                    self.true_count = self.true_count.saturating_sub(1);
                }
            }
            (Value::Boolean(old_v), Value::Boolean(new_v)) => {
                match (*old_v, *new_v) {
                    (true, false) => self.true_count = self.true_count.saturating_sub(1),
                    (false, true) => self.true_count += 1,
                    _ => {}
                }
            }
            // Any mismatched types were coerced to null already.
            _ => {}
        }
        false
    }

    fn apply_overlays_to_chunk(&mut self, chunk_idx: usize, updates: &[(usize, Value)]) {
        let (len, mut values, mut validity) = match self.chunks.get(chunk_idx) {
            Some(EncodedChunk::Bool(c)) => (c.len, c.decode_bools(), c.validity.clone()),
            _ => return,
        };

        for (in_chunk, value) in updates {
            if *in_chunk >= len {
                continue;
            }
            if let Value::Boolean(v) = value {
                values.set(*in_chunk, *v);
                if let Some(validity) = &mut validity {
                    validity.set(*in_chunk, true);
                }
            } else {
                values.set(*in_chunk, false);
                let validity = validity.get_or_insert_with(|| bitvec_all_true(len));
                validity.set(*in_chunk, false);
            }
        }

        let mut data = vec![0u8; (len + 7) / 8];
        for i in 0..len {
            if values.get(i) {
                data[i / 8] |= 1u8 << (i % 8);
            }
        }

        let validity = validity.and_then(|v| (!v.all_true()).then_some(v));
        self.chunks[chunk_idx] = EncodedChunk::Bool(BoolChunk { len, data, validity });
    }

    fn apply_overlays_to_current(&mut self, updates: &[(usize, Value)]) {
        for (in_chunk, value) in updates {
            if *in_chunk >= self.current.len() {
                continue;
            }
            if let Value::Boolean(v) = value {
                self.current.set(*in_chunk, *v);
                self.validity.set(*in_chunk, true);
            } else {
                self.current.set(*in_chunk, false);
                self.validity.set(*in_chunk, false);
            }
        }
    }

    fn take_tail_chunk_into_current(&mut self, expected_len: usize) {
        let Some(last) = self.chunks.last() else {
            return;
        };
        if last.len() != expected_len {
            return;
        }

        let Some(EncodedChunk::Bool(chunk)) = self.chunks.pop() else {
            return;
        };

        let values = chunk.decode_bools();
        let validity = chunk
            .validity
            .unwrap_or_else(|| BitVec::with_len_all_true(expected_len));
        self.current = values;
        self.validity = validity;
    }

    fn encoded_current_chunk(&self) -> Option<EncodedChunk> {
        if self.current.is_empty() {
            return None;
        }

        let len = self.current.len();
        let mut data = vec![0u8; (len + 7) / 8];
        for i in 0..len {
            if self.current.get(i) {
                data[i / 8] |= 1u8 << (i % 8);
            }
        }

        let validity = if self.validity.all_true() {
            None
        } else {
            Some(self.validity.clone())
        };

        Some(EncodedChunk::Bool(BoolChunk { len, data, validity }))
    }

    fn as_column_snapshot(&self) -> Column {
        let mut chunks = self.chunks.clone();
        if let Some(chunk) = self.encoded_current_chunk() {
            chunks.push(chunk);
        }
        Column {
            schema: self.schema.clone(),
            chunks: Arc::new(chunks),
            stats: self.stats(),
            dictionary: None,
            distinct: None,
        }
    }

    fn into_column(self) -> Column {
        let stats = self.stats();
        Column {
            schema: self.schema,
            chunks: Arc::new(self.chunks),
            stats,
            dictionary: None,
            distinct: None,
        }
    }
}

impl MutableDictColumn {
    fn new(schema: ColumnSchema, page_size: usize) -> Self {
        let mut current = Vec::new();
        let _ = current.try_reserve_exact(page_size);
        Self {
            schema,
            page_size,
            dictionary: Arc::new(Vec::new()),
            dict_map: HashMap::new(),
            current,
            validity: BitVec::with_capacity_bits(page_size),
            chunks: Vec::new(),
            null_count: 0,
            min: None,
            max: None,
            total_len: 0,
        }
    }

    fn from_column(col: Column, page_size: usize) -> Self {
        let Column {
            schema,
            chunks,
            stats,
            dictionary,
            distinct: _,
        } = col;

        let chunks = match Arc::try_unwrap(chunks) {
            Ok(chunks) => chunks,
            Err(chunks) => (*chunks).clone(),
        };

        let dict = dictionary.unwrap_or_else(|| Arc::new(Vec::new()));
        let mut dict_map = HashMap::new();
        let _ = dict_map.try_reserve(dict.len());
        for (idx, s) in dict.iter().cloned().enumerate() {
            dict_map.insert(s, idx as u32);
        }

        let (min, max, total_len) = {
            let min = stats.min.as_ref().and_then(|v| match v {
                Value::String(s) => Some(s.clone()),
                _ => None,
            });
            let max = stats.max.as_ref().and_then(|v| match v {
                Value::String(s) => Some(s.clone()),
                _ => None,
            });
            let non_null =
                (chunks.iter().map(|c| c.len() as u64).sum::<u64>()).saturating_sub(stats.null_count)
                .max(1);
            let avg = stats.avg_length.unwrap_or(0.0);
            let total_len = (avg * non_null as f64).round().max(0.0) as u64;
            (min, max, total_len)
        };

        let mut current = Vec::new();
        let _ = current.try_reserve_exact(page_size);
        Self {
            schema,
            page_size,
            dictionary: dict,
            dict_map,
            current,
            validity: BitVec::with_capacity_bits(page_size),
            chunks,
            null_count: stats.null_count,
            min,
            max,
            total_len,
        }
    }

    fn intern(&mut self, s: Arc<str>) -> u32 {
        if let Some(idx) = self.dict_map.get(s.as_ref()) {
            return *idx;
        }

        let dict = Arc::make_mut(&mut self.dictionary);
        let idx = dict.len() as u32;
        dict.push(s.clone());
        self.dict_map.insert(s, idx);
        idx
    }

    fn chunk_index(&self, row: usize, page_size: usize) -> (usize, usize) {
        (row / page_size, row % page_size)
    }

    fn get_cell(&self, row: usize, page_size: usize) -> Value {
        let (chunk_idx, in_chunk) = self.chunk_index(row, page_size);
        if let Some(chunk) = self.chunks.get(chunk_idx) {
            let EncodedChunk::Dict(c) = chunk else {
                return Value::Null;
            };
            return c
                .get_index(in_chunk)
                .and_then(|idx| self.dictionary.get(idx as usize).cloned())
                .map(Value::String)
                .unwrap_or(Value::Null);
        }

        if chunk_idx == self.chunks.len() {
            if in_chunk < self.current.len() {
                if self.validity.get(in_chunk) {
                    let idx = self.current[in_chunk] as usize;
                    if let Some(s) = self.dictionary.get(idx) {
                        return Value::String(s.clone());
                    }
                }
            }
        }

        Value::Null
    }

    fn decode_chunk(&self, chunk_idx: usize) -> Option<DecodedChunk> {
        let chunk = self.chunks.get(chunk_idx)?;
        let EncodedChunk::Dict(c) = chunk else {
            return None;
        };
        Some(DecodedChunk::Dict {
            indices: c.decode_indices(),
            validity: c.validity.clone(),
            dictionary: self.dictionary.clone(),
        })
    }

    fn push(&mut self, value: &Value) {
        match value {
            Value::Null => {
                self.current.push(0);
                self.validity.push(false);
                self.null_count += 1;
            }
            Value::String(s) => {
                let idx = self.intern(s.clone());
                self.current.push(idx);
                self.validity.push(true);
                self.total_len += s.len() as u64;

                self.min = match &self.min {
                    Some(m) if m.as_ref() <= s.as_ref() => Some(m.clone()),
                    _ => Some(s.clone()),
                };
                self.max = match &self.max {
                    Some(m) if m.as_ref() >= s.as_ref() => Some(m.clone()),
                    _ => Some(s.clone()),
                };
            }
            _ => {
                self.current.push(0);
                self.validity.push(false);
                self.null_count += 1;
            }
        }
    }

    fn flush(&mut self) {
        if self.current.is_empty() {
            return;
        }

        let indices = U32SequenceEncoding::encode(&self.current);
        let validity = if self.validity.all_true() {
            None
        } else {
            Some(self.validity.clone())
        };

        self.chunks.push(EncodedChunk::Dict(DictionaryEncodedChunk {
            len: self.current.len(),
            indices,
            validity,
        }));

        self.current.clear();
        self.validity = BitVec::with_capacity_bits(self.page_size);
    }

    fn stats(&self) -> ColumnStats {
        let total_rows: u64 =
            self.chunks.iter().map(|c| c.len() as u64).sum::<u64>() + self.current.len() as u64;
        let non_null = total_rows.saturating_sub(self.null_count);
        let non_null_f = non_null.max(1) as f64;
        ColumnStats {
            column_type: self.schema.column_type,
            distinct_count: self.dictionary.len() as u64,
            null_count: self.null_count,
            min: self.min.as_ref().map(|s| Value::String(s.clone())),
            max: self.max.as_ref().map(|s| Value::String(s.clone())),
            sum: None,
            avg_length: Some(self.total_len as f64 / non_null_f),
        }
    }

    fn apply_update(&mut self, old: &Value, new: &Value) -> bool {
        let old_s = match old {
            Value::String(s) => Some(s.clone()),
            _ => None,
        };
        let new_s = match new {
            Value::String(s) => Some(s.clone()),
            _ => None,
        };

        match (old_s, new_s) {
            (None, None) => false,
            (None, Some(s)) => {
                self.null_count = self.null_count.saturating_sub(1);
                self.total_len += s.len() as u64;
                let _ = self.intern(s.clone());
                self.min = match &self.min {
                    Some(m) if m.as_ref() <= s.as_ref() => Some(m.clone()),
                    _ => Some(s.clone()),
                };
                self.max = match &self.max {
                    Some(m) if m.as_ref() >= s.as_ref() => Some(m.clone()),
                    _ => Some(s.clone()),
                };
                false
            }
            (Some(s), None) => {
                self.null_count += 1;
                self.total_len = self.total_len.saturating_sub(s.len() as u64);
                self.min == Some(s.clone()) || self.max == Some(s)
            }
            (Some(old_s), Some(new_s)) => {
                self.total_len = self
                    .total_len
                    .saturating_add(new_s.len() as u64)
                    .saturating_sub(old_s.len() as u64);
                let _ = self.intern(new_s.clone());
                self.min = match &self.min {
                    Some(m) if m.as_ref() <= new_s.as_ref() => Some(m.clone()),
                    _ => Some(new_s.clone()),
                };
                self.max = match &self.max {
                    Some(m) if m.as_ref() >= new_s.as_ref() => Some(m.clone()),
                    _ => Some(new_s.clone()),
                };
                (self.min == Some(old_s.clone()) && new_s.as_ref() != old_s.as_ref())
                    || (self.max == Some(old_s.clone()) && new_s.as_ref() != old_s.as_ref())
            }
        }
    }

    fn apply_overlays_to_chunk(&mut self, chunk_idx: usize, updates: &[(usize, Value)]) {
        let (len, mut indices, mut validity) = match self.chunks.get(chunk_idx) {
            Some(EncodedChunk::Dict(c)) => (c.len, c.decode_indices(), c.validity.clone()),
            _ => return,
        };

        for (in_chunk, value) in updates {
            if *in_chunk >= len {
                continue;
            }
            if let Value::String(s) = value {
                let idx = self.intern(s.clone());
                indices[*in_chunk] = idx;
                if let Some(validity) = &mut validity {
                    validity.set(*in_chunk, true);
                }
            } else {
                indices[*in_chunk] = 0;
                let validity = validity.get_or_insert_with(|| bitvec_all_true(len));
                validity.set(*in_chunk, false);
            }
        }

        let indices = U32SequenceEncoding::encode(&indices);
        let validity = validity.and_then(|v| (!v.all_true()).then_some(v));

        self.chunks[chunk_idx] = EncodedChunk::Dict(DictionaryEncodedChunk { len, indices, validity });
    }

    fn apply_overlays_to_current(&mut self, updates: &[(usize, Value)]) {
        for (in_chunk, value) in updates {
            if *in_chunk >= self.current.len() {
                continue;
            }
            if let Value::String(s) = value {
                let idx = self.intern(s.clone());
                self.current[*in_chunk] = idx;
                self.validity.set(*in_chunk, true);
            } else {
                self.current[*in_chunk] = 0;
                self.validity.set(*in_chunk, false);
            }
        }
    }

    fn take_tail_chunk_into_current(&mut self, expected_len: usize) {
        let Some(last) = self.chunks.last() else {
            return;
        };
        if last.len() != expected_len {
            return;
        }

        let Some(EncodedChunk::Dict(chunk)) = self.chunks.pop() else {
            return;
        };

        let values = chunk.decode_indices();
        let validity = chunk
            .validity
            .unwrap_or_else(|| BitVec::with_len_all_true(expected_len));
        self.current = values;
        self.validity = validity;
    }

    fn encoded_current_chunk(&self) -> Option<EncodedChunk> {
        if self.current.is_empty() {
            return None;
        }

        let len = self.current.len();
        let indices = U32SequenceEncoding::encode(&self.current);
        let validity = if self.validity.all_true() {
            None
        } else {
            Some(self.validity.clone())
        };

        Some(EncodedChunk::Dict(DictionaryEncodedChunk { len, indices, validity }))
    }

    fn as_column_snapshot(&self) -> Column {
        let mut chunks = self.chunks.clone();
        if let Some(chunk) = self.encoded_current_chunk() {
            chunks.push(chunk);
        }
        Column {
            schema: self.schema.clone(),
            chunks: Arc::new(chunks),
            stats: self.stats(),
            dictionary: Some(self.dictionary.clone()),
            distinct: None,
        }
    }

    fn into_column(self) -> Column {
        let stats = self.stats();
        Column {
            schema: self.schema,
            chunks: Arc::new(self.chunks),
            stats,
            dictionary: Some(self.dictionary),
            distinct: None,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct ColumnarRange {
    pub row_start: usize,
    pub row_end: usize,
    pub col_start: usize,
    pub col_end: usize,
    /// Column-major output: `columns[c][r]`.
    pub columns: Vec<Vec<Value>>,
}

impl ColumnarRange {
    pub fn rows(&self) -> usize {
        self.row_end - self.row_start
    }

    pub fn cols(&self) -> usize {
        self.col_end - self.col_start
    }

    pub fn get(&self, row: usize, col: usize) -> Option<&Value> {
        self.columns.get(col)?.get(row)
    }
}

pub struct TableScan<'a> {
    table: &'a ColumnarTable,
}

fn canonical_f64_bits(v: f64) -> u64 {
    // Keep float grouping/equality consistent with `formula-columnar`'s `GROUP BY` implementation:
    // - Canonicalize `-0.0` to `0.0`
    // - Canonicalize NaNs so they compare/group together
    if v == 0.0 {
        0.0f64.to_bits()
    } else if v.is_nan() {
        f64::NAN.to_bits()
    } else {
        v.to_bits()
    }
}

impl<'a> TableScan<'a> {
    pub fn stats(&self, col: usize) -> Option<&'a ColumnStats> {
        self.table.columns.get(col).map(|c| &c.stats)
    }

    pub fn count_non_null(&self, col: usize) -> u64 {
        let Some(column) = self.table.columns.get(col) else {
            return 0;
        };
        let nulls = column.stats.null_count;
        self.table.rows as u64 - nulls
    }

    /// Scan a column and compute the sum (ignoring nulls).
    ///
    /// This is intended for analytics-style workloads where the UI or formula engine
    /// needs to aggregate over a large table without expanding it into per-cell maps.
    pub fn sum_f64(&self, col: usize) -> Option<f64> {
        let column = self.table.columns.get(col)?;
        match column.schema.column_type {
            ColumnType::Number
            | ColumnType::DateTime
            | ColumnType::Currency { .. }
            | ColumnType::Percentage { .. }
            | ColumnType::Boolean => {}
            ColumnType::String => return None,
        }

        let mut sum = 0f64;
        for chunk in column.chunks.iter() {
            match chunk {
                EncodedChunk::Float(c) => {
                    if let Some(validity) = &c.validity {
                        for (idx, v) in c.values.iter().enumerate() {
                            if validity.get(idx) {
                                sum += *v;
                            }
                        }
                    } else {
                        sum += c.values.iter().sum::<f64>();
                    }
                }
                EncodedChunk::Int(c) => {
                    let values = c.decode_i64();
                    if let Some(validity) = &c.validity {
                        for (idx, v) in values.iter().enumerate() {
                            if validity.get(idx) {
                                sum += *v as f64;
                            }
                        }
                    } else {
                        sum += values.iter().map(|v| *v as f64).sum::<f64>();
                    }
                }
                EncodedChunk::Bool(c) => {
                    let true_count = if let Some(validity) = &c.validity {
                        let mut cnt: u64 = 0;
                        for idx in 0..c.len {
                            if validity.get(idx) && c.get_bool(idx).unwrap_or(false) {
                                cnt += 1;
                            }
                        }
                        cnt
                    } else {
                        count_true_bits(&c.data, c.len)
                    };
                    sum += true_count as f64;
                }
                EncodedChunk::Dict(_) => {}
            }
        }

        Some(sum)
    }

    pub fn filter_eq_string(&self, col: usize, value: &str) -> Vec<usize> {
        let Some(column) = self.table.columns.get(col) else {
            return Vec::new();
        };
        let Some(dict) = column.dictionary.as_ref() else {
            return Vec::new();
        };

        let mut target: Option<u32> = None;
        for (idx, s) in dict.iter().enumerate() {
            if s.as_ref() == value {
                target = Some(idx as u32);
                break;
            }
        }
        let Some(target) = target else {
            return Vec::new();
        };

        let mut out = Vec::new();
        for (chunk_idx, chunk) in column.chunks.iter().enumerate() {
            let EncodedChunk::Dict(c) = chunk else {
                continue;
            };
            let base = chunk_idx * self.table.options.page_size_rows;
            for i in 0..c.len {
                if c.validity.as_ref().is_some_and(|v| !v.get(i)) {
                    continue;
                }
                if c.indices.get(i) == target {
                    out.push(base + i);
                }
            }
        }
        out
    }

    pub fn filter_in_string(&self, col: usize, values: &[&str]) -> Vec<usize> {
        if values.is_empty() {
            return Vec::new();
        }

        let Some(column) = self.table.columns.get(col) else {
            return Vec::new();
        };
        let Some(dict) = column.dictionary.as_ref() else {
            return Vec::new();
        };

        use std::collections::HashSet;
        let want: HashSet<&str> = values.iter().copied().collect();
        if want.is_empty() {
            return Vec::new();
        }

        let mut targets: HashSet<u32> = HashSet::new();
        for (idx, s) in dict.iter().enumerate() {
            if want.contains(s.as_ref()) {
                targets.insert(idx as u32);
            }
        }
        if targets.is_empty() {
            return Vec::new();
        }

        let mut out = Vec::new();
        for (chunk_idx, chunk) in column.chunks.iter().enumerate() {
            let EncodedChunk::Dict(c) = chunk else {
                continue;
            };
            let base = chunk_idx * self.table.options.page_size_rows;
            for i in 0..c.len {
                if c.validity.as_ref().is_some_and(|v| !v.get(i)) {
                    continue;
                }
                if targets.contains(&c.indices.get(i)) {
                    out.push(base + i);
                }
            }
        }
        out
    }

    /// Filter rows where `col == value` for a numeric (float) column.
    ///
    /// This operates directly on encoded chunks (no per-row `Value` allocations) and respects
    /// nullability (nulls are excluded).
    pub fn filter_eq_number(&self, col: usize, value: f64) -> Vec<usize> {
        let Some(column) = self.table.columns.get(col) else {
            return Vec::new();
        };
        if column.schema.column_type != ColumnType::Number {
            return Vec::new();
        }

        let target = canonical_f64_bits(value);
        let mut out = Vec::new();
        for (chunk_idx, chunk) in column.chunks.iter().enumerate() {
            let EncodedChunk::Float(c) = chunk else {
                continue;
            };
            let base = chunk_idx * self.table.options.page_size_rows;
            if let Some(validity) = &c.validity {
                for (i, v) in c.values.iter().enumerate() {
                    if !validity.get(i) {
                        continue;
                    }
                    if canonical_f64_bits(*v) == target {
                        out.push(base + i);
                    }
                }
            } else {
                for (i, v) in c.values.iter().enumerate() {
                    if canonical_f64_bits(*v) == target {
                        out.push(base + i);
                    }
                }
            }
        }
        out
    }

    /// Filter rows where `col IN (values...)` for a numeric (float) column.
    ///
    /// This operates directly on encoded chunks (no per-row `Value` allocations) and respects
    /// nullability (nulls are excluded).
    pub fn filter_in_number(&self, col: usize, values: &[f64]) -> Vec<usize> {
        if values.is_empty() {
            return Vec::new();
        }

        let Some(column) = self.table.columns.get(col) else {
            return Vec::new();
        };
        if column.schema.column_type != ColumnType::Number {
            return Vec::new();
        }

        use std::collections::HashSet;
        let mut targets: HashSet<u64> = HashSet::new();
        let _ = targets.try_reserve(values.len());
        for v in values {
            targets.insert(canonical_f64_bits(*v));
        }
        if targets.is_empty() {
            return Vec::new();
        }

        let mut out = Vec::new();
        for (chunk_idx, chunk) in column.chunks.iter().enumerate() {
            let EncodedChunk::Float(c) = chunk else {
                continue;
            };
            let base = chunk_idx * self.table.options.page_size_rows;
            if let Some(validity) = &c.validity {
                for (i, v) in c.values.iter().enumerate() {
                    if !validity.get(i) {
                        continue;
                    }
                    if targets.contains(&canonical_f64_bits(*v)) {
                        out.push(base + i);
                    }
                }
            } else {
                for (i, v) in c.values.iter().enumerate() {
                    if targets.contains(&canonical_f64_bits(*v)) {
                        out.push(base + i);
                    }
                }
            }
        }
        out
    }

    /// Filter rows where `col == value` for an int-backed logical column
    /// (`DateTime`/`Currency`/`Percentage`).
    pub fn filter_eq_i64(&self, col: usize, value: i64) -> Vec<usize> {
        let Some(column) = self.table.columns.get(col) else {
            return Vec::new();
        };
        match column.schema.column_type {
            ColumnType::DateTime | ColumnType::Currency { .. } | ColumnType::Percentage { .. } => {}
            _ => return Vec::new(),
        }

        let mut out = Vec::new();
        for (chunk_idx, chunk) in column.chunks.iter().enumerate() {
            let EncodedChunk::Int(c) = chunk else {
                continue;
            };
            let base = chunk_idx * self.table.options.page_size_rows;
            let decoded = c.decode_i64();
            if let Some(validity) = &c.validity {
                for (i, v) in decoded.iter().enumerate() {
                    if !validity.get(i) {
                        continue;
                    }
                    if *v == value {
                        out.push(base + i);
                    }
                }
            } else {
                for (i, v) in decoded.iter().enumerate() {
                    if *v == value {
                        out.push(base + i);
                    }
                }
            }
        }
        out
    }

    /// Filter rows where `col IN (values...)` for an int-backed logical column
    /// (`DateTime`/`Currency`/`Percentage`).
    pub fn filter_in_i64(&self, col: usize, values: &[i64]) -> Vec<usize> {
        if values.is_empty() {
            return Vec::new();
        }

        let Some(column) = self.table.columns.get(col) else {
            return Vec::new();
        };
        match column.schema.column_type {
            ColumnType::DateTime | ColumnType::Currency { .. } | ColumnType::Percentage { .. } => {}
            _ => return Vec::new(),
        }

        use std::collections::HashSet;
        let targets: HashSet<i64> = values.iter().copied().collect();
        if targets.is_empty() {
            return Vec::new();
        }

        let mut out = Vec::new();
        for (chunk_idx, chunk) in column.chunks.iter().enumerate() {
            let EncodedChunk::Int(c) = chunk else {
                continue;
            };
            let base = chunk_idx * self.table.options.page_size_rows;
            let decoded = c.decode_i64();
            if let Some(validity) = &c.validity {
                for (i, v) in decoded.iter().enumerate() {
                    if !validity.get(i) {
                        continue;
                    }
                    if targets.contains(v) {
                        out.push(base + i);
                    }
                }
            } else {
                for (i, v) in decoded.iter().enumerate() {
                    if targets.contains(v) {
                        out.push(base + i);
                    }
                }
            }
        }
        out
    }

    /// Filter rows where `col == value` for a boolean column.
    pub fn filter_eq_bool(&self, col: usize, value: bool) -> Vec<usize> {
        let Some(column) = self.table.columns.get(col) else {
            return Vec::new();
        };
        if column.schema.column_type != ColumnType::Boolean {
            return Vec::new();
        }

        let mut out = Vec::new();
        for (chunk_idx, chunk) in column.chunks.iter().enumerate() {
            let EncodedChunk::Bool(c) = chunk else {
                continue;
            };
            let base = chunk_idx * self.table.options.page_size_rows;
            if let Some(validity) = &c.validity {
                for i in 0..c.len {
                    if !validity.get(i) {
                        continue;
                    }
                    let byte = c.data[i / 8];
                    let bit = i % 8;
                    let v = ((byte >> bit) & 1) == 1;
                    if v == value {
                        out.push(base + i);
                    }
                }
            } else {
                for i in 0..c.len {
                    let byte = c.data[i / 8];
                    let bit = i % 8;
                    let v = ((byte >> bit) & 1) == 1;
                    if v == value {
                        out.push(base + i);
                    }
                }
            }
        }
        out
    }

    /// Filter rows where `col IN (values...)` for a boolean column.
    pub fn filter_in_bool(&self, col: usize, values: &[bool]) -> Vec<usize> {
        if values.is_empty() {
            return Vec::new();
        }

        let mut want_true = false;
        let mut want_false = false;
        for v in values {
            if *v {
                want_true = true;
            } else {
                want_false = true;
            }
        }

        if want_true && want_false {
            // Both true and false are accepted: return all non-null rows.
            let Some(column) = self.table.columns.get(col) else {
                return Vec::new();
            };
            if column.schema.column_type != ColumnType::Boolean {
                return Vec::new();
            }

            let mut out = Vec::new();
            for (chunk_idx, chunk) in column.chunks.iter().enumerate() {
                let EncodedChunk::Bool(c) = chunk else {
                    continue;
                };
                let base = chunk_idx * self.table.options.page_size_rows;
                if let Some(validity) = &c.validity {
                    for i in 0..c.len {
                        if validity.get(i) {
                            out.push(base + i);
                        }
                    }
                } else {
                    out.extend((0..c.len).map(|i| base + i));
                }
            }
            return out;
        }

        if want_true {
            self.filter_eq_bool(col, true)
        } else {
            self.filter_eq_bool(col, false)
        }
    }
}

pub struct ColumnarTableBuilder {
    schema: Vec<ColumnSchema>,
    options: TableOptions,
    builders: Vec<ColumnBuilder>,
    rows: usize,
}

enum ColumnBuilder {
    Int(IntBuilder),
    Float(FloatBuilder),
    Bool(BoolBuilder),
    Dict(DictBuilder),
}

struct IntBuilder {
    schema: ColumnSchema,
    page_size: usize,
    current: Vec<i64>,
    validity: BitVec,
    chunks: Vec<EncodedChunk>,
    stats: ColumnStats,
    distinct: DistinctCounter,
    min: Option<i64>,
    max: Option<i64>,
    sum: i128,
}

struct FloatBuilder {
    schema: ColumnSchema,
    page_size: usize,
    current: Vec<f64>,
    validity: BitVec,
    chunks: Vec<EncodedChunk>,
    stats: ColumnStats,
    distinct: DistinctCounter,
    min: Option<f64>,
    max: Option<f64>,
    sum: f64,
}

struct BoolBuilder {
    schema: ColumnSchema,
    page_size: usize,
    current: BitVec,
    validity: BitVec,
    chunks: Vec<EncodedChunk>,
    stats: ColumnStats,
    distinct: DistinctCounter,
    true_count: u64,
}

struct DictBuilder {
    schema: ColumnSchema,
    page_size: usize,
    dictionary: Vec<Arc<str>>,
    dict_map: std::collections::HashMap<Arc<str>, u32>,
    current: Vec<u32>,
    validity: BitVec,
    chunks: Vec<EncodedChunk>,
    stats: ColumnStats,
    min: Option<Arc<str>>,
    max: Option<Arc<str>>,
    total_len: u64,
}

impl ColumnarTableBuilder {
    pub fn new(schema: Vec<ColumnSchema>, options: TableOptions) -> Self {
        let builders = schema
            .iter()
            .cloned()
            .map(|col| match col.column_type {
                ColumnType::Number => {
                    ColumnBuilder::Float(FloatBuilder::new(col, options.page_size_rows))
                }
                ColumnType::String => {
                    ColumnBuilder::Dict(DictBuilder::new(col, options.page_size_rows))
                }
                ColumnType::Boolean => {
                    ColumnBuilder::Bool(BoolBuilder::new(col, options.page_size_rows))
                }
                ColumnType::DateTime
                | ColumnType::Currency { .. }
                | ColumnType::Percentage { .. } => {
                    ColumnBuilder::Int(IntBuilder::new(col, options.page_size_rows))
                }
            })
            .collect();

        Self {
            schema,
            options,
            builders,
            rows: 0,
        }
    }

    pub fn append_row(&mut self, row: &[Value]) {
        assert_eq!(
            row.len(),
            self.builders.len(),
            "row length must match schema"
        );

        for (builder, value) in self.builders.iter_mut().zip(row.iter()) {
            match builder {
                ColumnBuilder::Int(b) => b.push(value),
                ColumnBuilder::Float(b) => b.push(value),
                ColumnBuilder::Bool(b) => b.push(value),
                ColumnBuilder::Dict(b) => b.push(value),
            }
        }

        self.rows += 1;
        if self.rows % self.options.page_size_rows == 0 {
            for builder in &mut self.builders {
                match builder {
                    ColumnBuilder::Int(b) => b.flush(),
                    ColumnBuilder::Float(b) => b.flush(),
                    ColumnBuilder::Bool(b) => b.flush(),
                    ColumnBuilder::Dict(b) => b.flush(),
                }
            }
        }
    }

    /// Append a single value to a one-column builder.
    ///
    /// This is a micro-optimization for streaming ingestion patterns that construct a
    /// single-column table (e.g. when encoding calculated columns) and want to avoid the
    /// per-row slice creation + loop overhead of [`Self::append_row`].
    ///
    /// # Panics
    /// Panics if the builder schema has more than one column.
    pub fn append_value(&mut self, value: Value) {
        assert_eq!(
            self.builders.len(),
            1,
            "append_value requires a single-column schema"
        );

        if let Some(builder) = self.builders.get_mut(0) {
            match builder {
                ColumnBuilder::Int(b) => b.push(&value),
                ColumnBuilder::Float(b) => b.push(&value),
                ColumnBuilder::Bool(b) => b.push(&value),
                ColumnBuilder::Dict(b) => b.push(&value),
            }
        }

        self.rows += 1;
        if self.rows % self.options.page_size_rows == 0 {
            if let Some(builder) = self.builders.get_mut(0) {
                match builder {
                    ColumnBuilder::Int(b) => b.flush(),
                    ColumnBuilder::Float(b) => b.flush(),
                    ColumnBuilder::Bool(b) => b.flush(),
                    ColumnBuilder::Dict(b) => b.flush(),
                }
            }
        }
    }

    pub fn finalize(mut self) -> ColumnarTable {
        for builder in &mut self.builders {
            match builder {
                ColumnBuilder::Int(b) => b.flush(),
                ColumnBuilder::Float(b) => b.flush(),
                ColumnBuilder::Bool(b) => b.flush(),
                ColumnBuilder::Dict(b) => b.flush(),
            }
        }

        let mut columns: Vec<Column> = Vec::new();
        let _ = columns.try_reserve_exact(self.builders.len());
        for builder in self.builders {
            columns.push(match builder {
                ColumnBuilder::Int(b) => b.finish(),
                ColumnBuilder::Float(b) => b.finish(),
                ColumnBuilder::Bool(b) => b.finish(),
                ColumnBuilder::Dict(b) => b.finish(),
            });
        }

        ColumnarTable {
            schema: self.schema,
            columns,
            rows: self.rows,
            options: self.options,
            cache: Arc::new(Mutex::new(LruCache::new(self.options.cache.max_entries))),
        }
    }
}

impl IntBuilder {
    fn new(schema: ColumnSchema, page_size: usize) -> Self {
        let mut current = Vec::new();
        let _ = current.try_reserve_exact(page_size);
        Self {
            stats: ColumnStats {
                column_type: schema.column_type,
                ..ColumnStats::default()
            },
            schema,
            page_size,
            current,
            validity: BitVec::with_capacity_bits(page_size),
            chunks: Vec::new(),
            distinct: DistinctCounter::new(),
            min: None,
            max: None,
            sum: 0,
        }
    }

    fn push(&mut self, value: &Value) {
        let pushed = match value {
            Value::Null => {
                self.current.push(0);
                self.validity.push(false);
                self.stats.null_count += 1;
                return;
            }
            Value::DateTime(v) => Some(*v),
            Value::Currency(v) => Some(*v),
            Value::Percentage(v) => Some(*v),
            Value::Number(v) => Some(*v as i64),
            _ => {
                // Type mismatch: treat as null.
                self.current.push(0);
                self.validity.push(false);
                self.stats.null_count += 1;
                return;
            }
        };

        if let Some(v) = pushed {
            self.current.push(v);
            self.validity.push(true);
            self.distinct.insert_i64(v);
            self.sum += v as i128;
            self.min = Some(self.min.map(|m| m.min(v)).unwrap_or(v));
            self.max = Some(self.max.map(|m| m.max(v)).unwrap_or(v));
        }
    }

    fn flush(&mut self) {
        if self.current.is_empty() {
            return;
        }

        let mut min_valid: Option<i64> = None;
        for (idx, v) in self.current.iter().enumerate() {
            if self.validity.get(idx) {
                min_valid = Some(min_valid.map(|m| m.min(*v)).unwrap_or(*v));
            }
        }
        let min = min_valid.unwrap_or(0);
        let offsets: Vec<u64> = self
            .current
            .iter()
            .enumerate()
            .map(|(idx, v)| {
                if self.validity.get(idx) {
                    (*v as i128 - min as i128) as u64
                } else {
                    0
                }
            })
            .collect();
        let offsets = U64SequenceEncoding::encode(&offsets);

        let validity = if self.validity.all_true() {
            None
        } else {
            Some(self.validity.clone())
        };

        self.chunks.push(EncodedChunk::Int(ValueEncodedChunk {
            min,
            len: self.current.len(),
            offsets,
            validity,
        }));

        self.current.clear();
        self.validity = BitVec::with_capacity_bits(self.page_size);

        self.stats.distinct_count = self.distinct.estimate();
        self.stats.min = self.min.map(|v| value_from_i64(self.schema.column_type, v));
        self.stats.max = self.max.map(|v| value_from_i64(self.schema.column_type, v));
        self.stats.sum = Some(self.sum as f64);
    }

    fn finish(mut self) -> Column {
        self.flush();
        Column {
            schema: self.schema,
            chunks: Arc::new(self.chunks),
            stats: self.stats,
            dictionary: None,
            distinct: Some(Arc::new(self.distinct)),
        }
    }
}

impl FloatBuilder {
    fn new(schema: ColumnSchema, page_size: usize) -> Self {
        let mut current = Vec::new();
        let _ = current.try_reserve_exact(page_size);
        Self {
            stats: ColumnStats {
                column_type: schema.column_type,
                ..ColumnStats::default()
            },
            schema,
            page_size,
            current,
            validity: BitVec::with_capacity_bits(page_size),
            chunks: Vec::new(),
            distinct: DistinctCounter::new(),
            min: None,
            max: None,
            sum: 0.0,
        }
    }

    fn push(&mut self, value: &Value) {
        match value {
            Value::Null => {
                self.current.push(0.0);
                self.validity.push(false);
                self.stats.null_count += 1;
            }
            Value::Number(v) => {
                self.current.push(*v);
                self.validity.push(true);
                self.distinct.insert_i64(canonical_f64_bits(*v) as i64);
                self.sum += *v;
                self.min = Some(self.min.map(|m| m.min(*v)).unwrap_or(*v));
                self.max = Some(self.max.map(|m| m.max(*v)).unwrap_or(*v));
            }
            _ => {
                self.current.push(0.0);
                self.validity.push(false);
                self.stats.null_count += 1;
            }
        }
    }

    fn flush(&mut self) {
        if self.current.is_empty() {
            return;
        }

        let validity = if self.validity.all_true() {
            None
        } else {
            Some(self.validity.clone())
        };

        self.chunks.push(EncodedChunk::Float(FloatChunk {
            values: std::mem::take(&mut self.current),
            validity,
        }));

        self.validity = BitVec::with_capacity_bits(self.page_size);
        self.stats.distinct_count = self.distinct.estimate();
        self.stats.min = self.min.map(Value::Number);
        self.stats.max = self.max.map(Value::Number);
        self.stats.sum = Some(self.sum);
    }

    fn finish(mut self) -> Column {
        self.flush();
        Column {
            schema: self.schema,
            chunks: Arc::new(self.chunks),
            stats: self.stats,
            dictionary: None,
            distinct: Some(Arc::new(self.distinct)),
        }
    }
}

impl BoolBuilder {
    fn new(schema: ColumnSchema, page_size: usize) -> Self {
        Self {
            stats: ColumnStats {
                column_type: schema.column_type,
                ..ColumnStats::default()
            },
            schema,
            page_size,
            current: BitVec::with_capacity_bits(page_size),
            validity: BitVec::with_capacity_bits(page_size),
            chunks: Vec::new(),
            distinct: DistinctCounter::new(),
            true_count: 0,
        }
    }

    fn push(&mut self, value: &Value) {
        match value {
            Value::Null => {
                self.current.push(false);
                self.validity.push(false);
                self.stats.null_count += 1;
            }
            Value::Boolean(v) => {
                self.current.push(*v);
                self.validity.push(true);
                self.distinct.insert_bool(*v);
                if *v {
                    self.true_count += 1;
                }
            }
            _ => {
                self.current.push(false);
                self.validity.push(false);
                self.stats.null_count += 1;
            }
        }
    }

    fn flush(&mut self) {
        if self.current.is_empty() {
            return;
        }

        let len = self.current.len();
        let mut data = vec![0u8; (len + 7) / 8];
        for i in 0..len {
            if self.current.get(i) {
                data[i / 8] |= 1u8 << (i % 8);
            }
        }

        let validity = if self.validity.all_true() {
            None
        } else {
            Some(self.validity.clone())
        };

        self.chunks.push(EncodedChunk::Bool(BoolChunk {
            len,
            data,
            validity,
        }));

        self.current = BitVec::with_capacity_bits(self.page_size);
        self.validity = BitVec::with_capacity_bits(self.page_size);
        self.stats.distinct_count = self.distinct.estimate();
        self.stats.sum = Some(self.true_count as f64);
    }

    fn finish(mut self) -> Column {
        self.flush();
        Column {
            schema: self.schema,
            chunks: Arc::new(self.chunks),
            stats: self.stats,
            dictionary: None,
            distinct: None,
        }
    }
}

impl DictBuilder {
    fn new(schema: ColumnSchema, page_size: usize) -> Self {
        let mut current = Vec::new();
        let _ = current.try_reserve_exact(page_size);
        Self {
            stats: ColumnStats {
                column_type: schema.column_type,
                ..ColumnStats::default()
            },
            schema,
            page_size,
            dictionary: Vec::new(),
            dict_map: std::collections::HashMap::new(),
            current,
            validity: BitVec::with_capacity_bits(page_size),
            chunks: Vec::new(),
            min: None,
            max: None,
            total_len: 0,
        }
    }

    fn intern(&mut self, s: Arc<str>) -> u32 {
        if let Some(idx) = self.dict_map.get(s.as_ref()) {
            return *idx;
        }

        let idx = self.dictionary.len() as u32;
        self.dictionary.push(s.clone());
        self.dict_map.insert(s, idx);
        idx
    }

    fn push(&mut self, value: &Value) {
        match value {
            Value::Null => {
                self.current.push(0);
                self.validity.push(false);
                self.stats.null_count += 1;
            }
            Value::String(s) => {
                let idx = self.intern(s.clone());
                self.current.push(idx);
                self.validity.push(true);
                self.total_len += s.len() as u64;

                self.min = match &self.min {
                    Some(m) if m.as_ref() <= s.as_ref() => Some(m.clone()),
                    _ => Some(s.clone()),
                };
                self.max = match &self.max {
                    Some(m) if m.as_ref() >= s.as_ref() => Some(m.clone()),
                    _ => Some(s.clone()),
                };
            }
            _ => {
                self.current.push(0);
                self.validity.push(false);
                self.stats.null_count += 1;
            }
        }
    }

    fn flush(&mut self) {
        if self.current.is_empty() {
            return;
        }

        let indices = U32SequenceEncoding::encode(&self.current);
        let validity = if self.validity.all_true() {
            None
        } else {
            Some(self.validity.clone())
        };

        self.chunks.push(EncodedChunk::Dict(DictionaryEncodedChunk {
            len: self.current.len(),
            indices,
            validity,
        }));

        self.current.clear();
        self.validity = BitVec::with_capacity_bits(self.page_size);

        self.stats.distinct_count = self.dictionary.len() as u64;
        self.stats.min = self.min.as_ref().map(|s| Value::String(s.clone()));
        self.stats.max = self.max.as_ref().map(|s| Value::String(s.clone()));
        let non_null = (self.rows_non_null()).max(1) as f64;
        self.stats.avg_length = Some(self.total_len as f64 / non_null);
    }

    fn rows_non_null(&self) -> u64 {
        // Accurate for the rows seen so far: current buffer + flushed buffers.
        let total_rows: u64 =
            self.chunks.iter().map(|c| c.len() as u64).sum::<u64>() + self.current.len() as u64;
        total_rows.saturating_sub(self.stats.null_count)
    }

    fn finish(mut self) -> Column {
        self.flush();
        Column {
            schema: self.schema,
            chunks: Arc::new(self.chunks),
            stats: self.stats,
            dictionary: Some(Arc::new(self.dictionary)),
            distinct: None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn value_encoding_roundtrip() {
        let schema = vec![ColumnSchema {
            name: "id".to_owned(),
            column_type: ColumnType::DateTime,
        }];
        let options = TableOptions {
            page_size_rows: 128,
            cache: PageCacheConfig { max_entries: 8 },
        };

        let mut builder = ColumnarTableBuilder::new(schema, options);
        for i in 0..1000 {
            builder.append_row(&[Value::DateTime(i)]);
        }
        let table = builder.finalize();

        assert_eq!(table.row_count(), 1000);
        assert_eq!(table.get_cell(0, 0), Value::DateTime(0));
        assert_eq!(table.get_cell(999, 0), Value::DateTime(999));
        assert_eq!(table.get_cell(1000, 0), Value::Null);
    }

    #[test]
    fn dictionary_encoding_roundtrip_and_range() {
        let schema = vec![
            ColumnSchema {
                name: "cat".to_owned(),
                column_type: ColumnType::String,
            },
            ColumnSchema {
                name: "flag".to_owned(),
                column_type: ColumnType::Boolean,
            },
        ];

        let options = TableOptions {
            page_size_rows: 16,
            cache: PageCacheConfig { max_entries: 4 },
        };
        let mut builder = ColumnarTableBuilder::new(schema, options);

        let cats = ["A", "B", "A", "C", "C", "C"];
        for (i, c) in cats.iter().enumerate() {
            builder.append_row(&[
                Value::String(Arc::<str>::from(*c)),
                Value::Boolean(i % 2 == 0),
            ]);
        }
        let table = builder.finalize();

        assert_eq!(table.get_cell(0, 0), Value::String(Arc::<str>::from("A")));
        assert_eq!(table.get_cell(3, 0), Value::String(Arc::<str>::from("C")));

        let range = table.get_range(1, 5, 0, 2);
        assert_eq!(range.rows(), 4);
        assert_eq!(range.cols(), 2);
        assert_eq!(range.get(0, 0), Some(&Value::String(Arc::<str>::from("B"))));
        assert_eq!(range.get(3, 0), Some(&Value::String(Arc::<str>::from("C"))));
    }

    #[test]
    fn cache_hits_after_repeated_range() {
        let schema = vec![ColumnSchema {
            name: "x".to_owned(),
            column_type: ColumnType::Number,
        }];

        let options = TableOptions {
            page_size_rows: 64,
            cache: PageCacheConfig { max_entries: 2 },
        };
        let mut builder = ColumnarTableBuilder::new(schema, options);
        for i in 0..256 {
            builder.append_row(&[Value::Number(i as f64)]);
        }
        let table = builder.finalize();

        let _ = table.get_range(0, 32, 0, 1);
        let stats1 = table.cache_stats();
        let _ = table.get_range(0, 32, 0, 1);
        let stats2 = table.cache_stats();

        assert!(stats2.hits > stats1.hits);
    }

    #[test]
    fn scan_sum_and_filter_work() {
        let schema = vec![
            ColumnSchema {
                name: "x".to_owned(),
                column_type: ColumnType::Number,
            },
            ColumnSchema {
                name: "cat".to_owned(),
                column_type: ColumnType::String,
            },
        ];

        let options = TableOptions {
            page_size_rows: 8,
            cache: PageCacheConfig { max_entries: 2 },
        };

        let mut builder = ColumnarTableBuilder::new(schema, options);
        let cats = [
            Arc::<str>::from("A"),
            Arc::<str>::from("B"),
            Arc::<str>::from("C"),
            Arc::<str>::from("C"),
        ];
        for i in 0..12 {
            builder.append_row(&[
                Value::Number((i + 1) as f64),
                Value::String(cats[i % cats.len()].clone()),
            ]);
        }
        let table = builder.finalize();

        let sum = table.scan().sum_f64(0).unwrap();
        assert_eq!(sum, (1..=12).map(|v| v as f64).sum::<f64>());

        let rows = table.scan().filter_eq_string(1, "C");
        assert_eq!(rows, vec![2, 3, 6, 7, 10, 11]);
    }

    #[test]
    fn scan_filter_number_handles_nulls_and_canonicalization() {
        let schema = vec![ColumnSchema {
            name: "n".to_owned(),
            column_type: ColumnType::Number,
        }];
        let options = TableOptions {
            page_size_rows: 4,
            cache: PageCacheConfig { max_entries: 2 },
        };
        let mut builder = ColumnarTableBuilder::new(schema, options);

        let rows = [
            Value::Number(0.0),
            Value::Number(-0.0),
            Value::Number(f64::NAN),
            Value::Null,
            Value::Number(1.0),
            Value::Number(f64::NAN),
        ];
        for v in rows {
            builder.append_row(&[v]);
        }
        let table = builder.finalize();

        // -0.0 and 0.0 are canonicalized together, and nulls are excluded.
        assert_eq!(table.scan().filter_eq_number(0, -0.0), vec![0, 1]);
        assert_eq!(table.scan().filter_eq_number(0, 0.0), vec![0, 1]);

        // All NaNs are canonicalized together.
        assert_eq!(table.scan().filter_eq_number(0, f64::NAN), vec![2, 5]);

        assert_eq!(
            table.scan().filter_in_number(0, &[0.0, f64::NAN]),
            vec![0, 1, 2, 5]
        );
    }

    #[test]
    fn scan_filter_i64_handles_nulls_for_int_backed_columns() {
        let schema = vec![ColumnSchema {
            name: "dt".to_owned(),
            column_type: ColumnType::DateTime,
        }];
        let options = TableOptions {
            page_size_rows: 4,
            cache: PageCacheConfig { max_entries: 2 },
        };
        let mut builder = ColumnarTableBuilder::new(schema, options);
        let rows = [
            Value::DateTime(10),
            Value::Null,
            Value::DateTime(11),
            Value::DateTime(10),
            Value::DateTime(12),
        ];
        for v in rows {
            builder.append_row(&[v]);
        }
        let table = builder.finalize();

        assert_eq!(table.scan().filter_eq_i64(0, 10), vec![0, 3]);
        assert_eq!(table.scan().filter_in_i64(0, &[11, 10]), vec![0, 2, 3]);

    }

    #[test]
    fn cloning_columnar_table_preserves_reads_and_group_by_results() {
        use crate::query::AggSpec;

        let schema = vec![
            ColumnSchema {
                name: "cat".to_owned(),
                column_type: ColumnType::String,
            },
            ColumnSchema {
                name: "x".to_owned(),
                column_type: ColumnType::Number,
            },
        ];

        let options = TableOptions {
            page_size_rows: 4,
            cache: PageCacheConfig { max_entries: 4 },
        };

        let mut builder = ColumnarTableBuilder::new(schema, options);
        for i in 0..10 {
            let cat = if i % 2 == 0 { "A" } else { "B" };
            builder.append_row(&[
                Value::String(Arc::<str>::from(cat)),
                Value::Number(i as f64),
            ]);
        }
        let table = builder.finalize();
        let cloned = table.clone();

        // Ensure the clone shares the encoded chunk backing storage.
        for (orig_col, cloned_col) in table.columns.iter().zip(cloned.columns.iter()) {
            assert!(Arc::ptr_eq(&orig_col.chunks, &cloned_col.chunks));
            match (&orig_col.distinct, &cloned_col.distinct) {
                (Some(a), Some(b)) => assert!(Arc::ptr_eq(a, b)),
                (None, None) => {}
                _ => panic!("cloned table should preserve distinct sketch presence"),
            }
        }

        for row in 0..=table.row_count() {
            for col in 0..table.column_count() {
                assert_eq!(table.get_cell(row, col), cloned.get_cell(row, col));
            }
        }

        let keys = [0usize];
        let aggs = [AggSpec::count_rows(), AggSpec::sum_f64(1)];
        let grouped = table.group_by(&keys, &aggs).unwrap().to_values();
        let grouped_cloned = cloned.group_by(&keys, &aggs).unwrap().to_values();
        assert_eq!(grouped, grouped_cloned);
    }

    #[test]
    fn into_mutable_roundtrip_works_after_clone() {
        let schema = vec![
            ColumnSchema {
                name: "cat".to_owned(),
                column_type: ColumnType::String,
            },
            ColumnSchema {
                name: "x".to_owned(),
                column_type: ColumnType::Number,
            },
        ];

        let options = TableOptions {
            page_size_rows: 4,
            cache: PageCacheConfig { max_entries: 4 },
        };

        let mut builder = ColumnarTableBuilder::new(schema, options);
        builder.append_row(&[Value::String(Arc::<str>::from("A")), Value::Number(1.0)]);
        builder.append_row(&[Value::String(Arc::<str>::from("B")), Value::Number(2.0)]);
        let table = builder.finalize();

        // Converting a cloned table to mutable forces recovery of chunk ownership (may clone the
        // underlying `Vec<EncodedChunk>`), but must remain correct.
        let mut mutable = table.clone().into_mutable();
        mutable.append_row(&[Value::String(Arc::<str>::from("C")), Value::Number(3.0)]);
        let frozen = mutable.freeze();

        assert_eq!(frozen.row_count(), 3);
        assert_eq!(frozen.column_count(), 2);
        assert_eq!(frozen.get_cell(0, 0), table.get_cell(0, 0));
        assert_eq!(frozen.get_cell(1, 0), table.get_cell(1, 0));
        assert_eq!(frozen.get_cell(2, 0), Value::String(Arc::<str>::from("C")));
        assert_eq!(frozen.get_cell(2, 1), Value::Number(3.0));
    }
}
