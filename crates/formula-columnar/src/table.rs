#![forbid(unsafe_code)]

use crate::bitmap::BitVec;
use crate::cache::{CacheStats, LruCache, PageCacheConfig};
use crate::encoding::{
    BoolChunk, DecodedChunk, DictionaryEncodedChunk, EncodedChunk, FloatChunk, U32SequenceEncoding,
    U64SequenceEncoding, ValueEncodedChunk,
};
use crate::stats::{ColumnStats, DistinctCounter};
use crate::types::{ColumnType, Value};
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

#[derive(Clone, Debug)]
pub struct Column {
    schema: ColumnSchema,
    chunks: Vec<EncodedChunk>,
    stats: ColumnStats,
    dictionary: Option<Arc<Vec<Arc<str>>>>,
}

impl Column {
    pub fn name(&self) -> &str {
        &self.schema.name
    }

    pub fn column_type(&self) -> ColumnType {
        self.schema.column_type
    }

    pub fn stats(&self) -> &ColumnStats {
        &self.stats
    }

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
                .map(|idx| Value::String(dict[idx as usize].clone()))
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

impl ColumnarTable {
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
        self.columns.get(col)?.dictionary.clone()
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

        let mut cache = self.cache.lock().expect("columnar page cache poisoned");
        if let Some(hit) = cache.get(&key) {
            return Some(hit);
        }

        drop(cache);
        let decoded = self.columns.get(col)?.decode_chunk(chunk_idx)?;
        let decoded = Arc::new(decoded);

        let mut cache = self.cache.lock().expect("columnar page cache poisoned");
        cache.insert(key, decoded.clone());
        Some(decoded)
    }

    pub fn cache_stats(&self) -> CacheStats {
        self.cache
            .lock()
            .expect("columnar page cache poisoned")
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

        let mut out_columns: Vec<Vec<Value>> = Vec::with_capacity(cols);
        for col in col_start..col_end {
            let mut values = Vec::with_capacity(rows);
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
        for chunk in &column.chunks {
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

    pub fn finalize(mut self) -> ColumnarTable {
        for builder in &mut self.builders {
            match builder {
                ColumnBuilder::Int(b) => b.flush(),
                ColumnBuilder::Float(b) => b.flush(),
                ColumnBuilder::Bool(b) => b.flush(),
                ColumnBuilder::Dict(b) => b.flush(),
            }
        }

        let mut columns: Vec<Column> = Vec::with_capacity(self.builders.len());
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
        Self {
            stats: ColumnStats {
                column_type: schema.column_type,
                ..ColumnStats::default()
            },
            schema,
            page_size,
            current: Vec::with_capacity(page_size),
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
            chunks: self.chunks,
            stats: self.stats,
            dictionary: None,
        }
    }
}

impl FloatBuilder {
    fn new(schema: ColumnSchema, page_size: usize) -> Self {
        Self {
            stats: ColumnStats {
                column_type: schema.column_type,
                ..ColumnStats::default()
            },
            schema,
            page_size,
            current: Vec::with_capacity(page_size),
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
                self.distinct.insert_i64(v.to_bits() as i64);
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
            chunks: self.chunks,
            stats: self.stats,
            dictionary: None,
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
            chunks: self.chunks,
            stats: self.stats,
            dictionary: None,
        }
    }
}

impl DictBuilder {
    fn new(schema: ColumnSchema, page_size: usize) -> Self {
        Self {
            stats: ColumnStats {
                column_type: schema.column_type,
                ..ColumnStats::default()
            },
            schema,
            page_size,
            dictionary: Vec::new(),
            dict_map: std::collections::HashMap::new(),
            current: Vec::with_capacity(page_size),
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
            chunks: self.chunks,
            stats: self.stats,
            dictionary: Some(Arc::new(self.dictionary)),
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
}
