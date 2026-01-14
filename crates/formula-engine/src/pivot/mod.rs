//! Pivot table engine + cache.
//!
//! This module is intentionally self-contained: it operates on a cached
//! rectangular dataset (headers + records) and produces a 2D grid that can be
//! rendered into a worksheet range.
//!
//! The goal is an MVP pivot engine that supports the core spreadsheet workflow:
//! - Create/refresh a `PivotCache` from a source range (headers + rows)
//! - Configure row/column/value/filter fields
//! - Compute aggregations (sum/count/avg/min/max + stddev/var variants)
//! - Produce a table with grand totals and basic subtotals.

#[cfg(test)]
use chrono::NaiveDate;
use formula_columnar::{ColumnarTable, Value as ColumnarValue};
use std::borrow::Cow;
use std::cmp::Ordering;
use std::collections::{BTreeMap, HashMap, HashSet};

pub use formula_model::pivots::{
    AggregationType, CalculatedField, CalculatedItem, FilterField, GrandTotals, Layout, PivotConfig,
    PivotField, PivotFieldRef, PivotKeyPart, PivotValue, ShowAsType, SortOrder, SubtotalPosition,
    ValueField,
};
use serde::{Deserialize, Serialize};
use thiserror::Error;

use std::sync::atomic::{AtomicU64, Ordering as AtomicOrdering};

mod definition;
pub mod source;
pub(crate) use definition::refresh_pivot;
pub(crate) use definition::PivotRefreshContext;
pub use definition::{
    PivotDestination, PivotRefreshError, PivotRefreshOutput, PivotSource, PivotTableDefinition,
    PivotTableId,
};
pub(crate) fn pivot_field_ref_name(field: &PivotFieldRef) -> Cow<'_, str> {
    match field {
        // Cache-backed pivots use the header caption directly.
        PivotFieldRef::CacheFieldName(name) => Cow::Borrowed(name.as_str()),
        // Data Model measures are frequently stored in cache headers without DAX brackets, so
        // prefer the raw measure name and fall back to the `[Measure]` display form elsewhere.
        PivotFieldRef::DataModelMeasure(name) => Cow::Borrowed(name.as_str()),
        // Data Model column refs may be stored in cache headers with or without DAX quoting for
        // table names (e.g. `Dim Product[Category]` vs `'Dim Product'[Category]`). Prefer the
        // unquoted `{table}[{column}]` form and fall back to the quoted DAX form elsewhere.
        PivotFieldRef::DataModelColumn { table, column } => {
            let column = escape_dax_bracket_identifier(column);
            Cow::Owned(format!("{table}[{column}]"))
        }
    }
}

fn escape_dax_bracket_identifier(raw: &str) -> Cow<'_, str> {
    // In DAX, `]` is escaped as `]]` within `[...]`.
    if raw.contains(']') {
        Cow::Owned(raw.replace(']', "]]"))
    } else {
        Cow::Borrowed(raw)
    }
}

fn dax_quoted_table_name(raw: &str) -> Cow<'_, str> {
    // In DAX, table names that require quoting use single quotes, and embedded single quotes are
    // escaped by doubling (e.g. `O'Brien` => `'O''Brien'`).
    if raw.contains('\'') {
        Cow::Owned(raw.replace('\'', "''"))
    } else {
        Cow::Borrowed(raw)
    }
}

fn dax_quoted_column_ref(table: &str, column: &str) -> String {
    let table = dax_quoted_table_name(table);
    let column = escape_dax_bracket_identifier(column);
    format!("'{table}'[{column}]")
}

fn pivot_field_ref_caption(field: &PivotFieldRef) -> Cow<'_, str> {
    // Value field captions should use the human-facing field name (Excel-like), not the DAX
    // reference form. In particular, measures are displayed without surrounding brackets.
    match field {
        PivotFieldRef::DataModelMeasure(measure) => Cow::Borrowed(measure.as_str()),
        _ => pivot_field_ref_name(field),
    }
}

#[allow(dead_code)]
fn pivot_field_ref_from_legacy_string(raw: String) -> PivotFieldRef {
    PivotFieldRef::from_unstructured_owned(raw)
}

mod apply;
pub use apply::{
    apply_pivot_cell_writes_to_worksheet, apply_pivot_result_to_worksheet, PivotApplyError,
    PivotApplyOptions,
};
#[derive(Debug, Error)]
pub enum PivotError {
    #[error("worksheet not found: {0}")]
    SheetNotFound(String),
    #[error("missing field in pivot cache: {0}")]
    MissingField(String),
    #[error("pivot table must have at least one value field")]
    NoValueFields,
    #[error("duplicate calculated field: {0}")]
    DuplicateCalculatedField(String),
    #[error("calculated field name conflicts with source field: {0}")]
    CalculatedFieldNameConflictsWithSource(String),
    #[error("calculated item field not in layout: {0}")]
    CalculatedItemFieldNotInLayout(String),
    #[error("calculated items require a PivotCache-backed record source")]
    CalculatedItemsRequirePivotCache,
    #[error("calculated item name conflicts with existing item in field {field}: {item}")]
    CalculatedItemNameConflictsWithExistingItem { field: String, item: String },
    #[error("invalid calculated field formula for {field}: {message}")]
    InvalidCalculatedFieldFormula { field: String, message: String },
    #[error("invalid calculated item formula for {field}::{item}: {message}")]
    InvalidCalculatedItemFormula {
        field: String,
        item: String,
        message: String,
    },
}

fn pivot_key_part_to_pivot_value(part: &PivotKeyPart) -> PivotValue {
    match part {
        // Excel renders blank pivot items as the literal string "(blank)".
        PivotKeyPart::Blank => PivotValue::Text("(blank)".to_string()),
        PivotKeyPart::Number(bits) => PivotValue::Number(f64::from_bits(*bits)),
        PivotKeyPart::Date(d) => PivotValue::Date(*d),
        PivotKeyPart::Text(s) => PivotValue::Text(s.clone()),
        PivotKeyPart::Bool(b) => PivotValue::Bool(*b),
    }
}
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PivotKey(pub Vec<PivotKeyPart>);

impl PivotKey {
    fn display_strings(&self) -> Vec<String> {
        self.0.iter().map(|p| p.display_string()).collect()
    }
}

impl PartialOrd for PivotKey {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for PivotKey {
    fn cmp(&self, other: &Self) -> Ordering {
        self.0.cmp(&other.0)
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CacheField {
    pub name: String,
    pub index: usize,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotCache {
    pub fields: Vec<CacheField>,
    pub records: Vec<Vec<PivotValue>>,
    pub unique_values: HashMap<String, Vec<PivotValue>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum PivotFieldType {
    Blank,
    Number,
    Date,
    Text,
    Bool,
    Mixed,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotSchemaField {
    pub name: String,
    pub field_type: PivotFieldType,
    pub sample_values: Vec<PivotValue>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotSchema {
    pub fields: Vec<PivotSchemaField>,
    pub record_count: usize,
}

impl PivotCache {
    /// Normalizes header captions into non-empty, unique (case-insensitive) pivot field names.
    ///
    /// Excel pivot field captions must be unique and cannot be blank. When we build a cache from a
    /// worksheet/table range we emulate that behavior:
    /// - Trim leading/trailing whitespace.
    /// - If the caption is empty after trimming, assign `Column{n}` where `n` is the 1-based
    ///   ordinal of the blank column in the header row.
    /// - If a caption collides with a previous caption (case-insensitive), append `" (2)"`,
    ///   `" (3)"`, ... until the name is unique.
    fn normalize_field_names(headers: &[PivotValue]) -> Vec<String> {
        let mut out = Vec::with_capacity(headers.len());
        let mut used_folded: HashSet<String> = HashSet::with_capacity(headers.len());
        let mut blank_counter = 0usize;

        for header in headers {
            let mut base = header.display_string();
            base = base.trim().to_string();

            if base.is_empty() {
                blank_counter += 1;
                base = format!("Column{blank_counter}");
            }

            let mut name = base.clone();
            if used_folded.contains(&fold_text_case_insensitive(&name)) {
                let mut suffix = 2usize;
                loop {
                    name = format!("{base} ({suffix})");
                    let folded = fold_text_case_insensitive(&name);
                    if !used_folded.contains(&folded) {
                        break;
                    }
                    suffix += 1;
                }
            }

            used_folded.insert(fold_text_case_insensitive(&name));
            out.push(name);
        }

        out
    }

    pub fn from_range(range: &[Vec<PivotValue>]) -> Result<Self, PivotError> {
        if range.is_empty() {
            return Ok(Self {
                fields: Vec::new(),
                records: Vec::new(),
                unique_values: HashMap::new(),
            });
        }

        let headers = &range[0];
        let normalized_names = Self::normalize_field_names(headers);
        let mut fields = Vec::with_capacity(headers.len());
        for (idx, name) in normalized_names.into_iter().enumerate() {
            fields.push(CacheField { name, index: idx });
        }

        let records = range[1..].to_vec();

        let mut unique_values: HashMap<String, BTreeMap<PivotKeyPart, PivotValue>> = HashMap::new();

        for row in &records {
            for field in &fields {
                let value = row.get(field.index).cloned().unwrap_or(PivotValue::Blank);
                unique_values
                    .entry(field.name.clone())
                    .or_default()
                    .entry(value.to_key_part())
                    .or_insert(value);
            }
        }

        let mut unique_values_final = HashMap::new();
        for field in &fields {
            let values = unique_values
                .get(&field.name)
                .map(|map| map.values().cloned().collect())
                .unwrap_or_default();
            unique_values_final.insert(field.name.clone(), values);
        }

        Ok(Self {
            fields,
            records,
            unique_values: unique_values_final,
        })
    }

    pub fn field_index(&self, name: &str) -> Option<usize> {
        self.fields.iter().find(|f| f.name == name).map(|f| f.index)
    }

    pub fn field_index_ref(&self, field: &PivotFieldRef) -> Option<usize> {
        if let Some(name) = field.as_cache_field_name() {
            return self.field_index(name);
        }

        // Best-effort: match Data Model refs against cache field captions. Caches may store:
        // - Measures either as the raw name (`Total`) or in DAX bracket form (`[Total]`).
        // - Column refs with or without quoted table names (`Sales[Region]` vs `'Sales Table'[Region]`).
        //
        // Try a few common textual encodings in priority order so callers can resolve fields
        // regardless of how the cache captions were generated.
        let canonical = field.canonical_name();
        if let Some(idx) = self.field_index(canonical.as_ref()) {
            return Some(idx);
        }

        let label = pivot_field_ref_name(field);
        if let Some(idx) = self.field_index(label.as_ref()) {
            return Some(idx);
        }

        match field {
            PivotFieldRef::CacheFieldName(_) => {}
            PivotFieldRef::DataModelMeasure(name) => {
                if let Some(idx) = self.field_index(name) {
                    return Some(idx);
                }
            }
            PivotFieldRef::DataModelColumn { table, column } => {
                // Some caches store DAX-quoted table captions (e.g. `'Sales Table'[Region]`).
                let quoted = dax_quoted_column_ref(table, column);
                if let Some(idx) = self.field_index(&quoted) {
                    return Some(idx);
                }

                // Unquoted table name + escaped bracket identifier.
                let escaped_column = column.replace(']', "]]");
                let unquoted = format!("{table}[{escaped_column}]");
                if let Some(idx) = self.field_index(&unquoted) {
                    return Some(idx);
                }

                // Some caches store the raw column name without DAX escaping.
                let unescaped = format!("{table}[{column}]");
                if unescaped != unquoted {
                    if let Some(idx) = self.field_index(&unescaped) {
                        return Some(idx);
                    }
                }

                // Rare: quoted table name but raw (unescaped) column name.
                let table = dax_quoted_table_name(table);
                let quoted_unescaped = format!("'{table}'[{column}]");
                if quoted_unescaped != quoted {
                    if let Some(idx) = self.field_index(&quoted_unescaped) {
                        return Some(idx);
                    }
                }
            }
        }

        // Final fallback: try the canonical `Display` rendering. This handles Data Model columns
        // whose table names require DAX quoting (e.g. `'Sales Table'[Region]`).
        let display = field.to_string();
        if let Some(idx) = self.field_index(&display) {
            return Some(idx);
        }

        None
    }

    pub fn refresh_from_range(&mut self, range: &[Vec<PivotValue>]) -> Result<(), PivotError> {
        *self = Self::from_range(range)?;
        Ok(())
    }

    /// Returns a lightweight schema suitable for AI tool prompting.
    pub fn schema(&self, sample_size: usize) -> PivotSchema {
        let mut fields = Vec::with_capacity(self.fields.len());
        for field in &self.fields {
            let mut samples = Vec::new();
            let mut saw_number = false;
            let mut saw_date = false;
            let mut saw_text = false;
            let mut saw_bool = false;
            let mut saw_blank = false;

            for row in self.records.iter().take(sample_size) {
                let value = row.get(field.index).cloned().unwrap_or(PivotValue::Blank);
                match &value {
                    PivotValue::Blank => saw_blank = true,
                    PivotValue::Number(_) => saw_number = true,
                    PivotValue::Date(_) => saw_date = true,
                    PivotValue::Text(_) => saw_text = true,
                    PivotValue::Bool(_) => saw_bool = true,
                }
                samples.push(value);
            }

            let field_type = match (saw_number, saw_date, saw_text, saw_bool, saw_blank) {
                (false, false, false, false, true) => PivotFieldType::Blank,
                (true, false, false, false, _) => PivotFieldType::Number,
                (false, true, false, false, _) => PivotFieldType::Date,
                (false, false, true, false, _) => PivotFieldType::Text,
                (false, false, false, true, _) => PivotFieldType::Bool,
                _ => PivotFieldType::Mixed,
            };

            fields.push(PivotSchemaField {
                name: field.name.clone(),
                field_type,
                sample_values: samples,
            });
        }

        PivotSchema {
            fields,
            record_count: self.records.len(),
        }
    }
}

/// A lightweight pivot value view returned by [`PivotRecordSource`].
///
/// The pivot engine only needs short-lived access to each cell to build grouping keys and update
/// aggregations. Returning a `PivotValueRef` lets sources either:
/// - borrow values from an in-memory cache (`Borrowed`)
/// - synthesize values on demand from a columnar/streaming source (`Owned`)
#[derive(Debug)]
pub enum PivotValueRef<'a> {
    Borrowed(&'a PivotValue),
    Owned(PivotValue),
}

impl PivotValueRef<'_> {
    fn as_value(&self) -> &PivotValue {
        match self {
            PivotValueRef::Borrowed(v) => v,
            PivotValueRef::Owned(v) => v,
        }
    }

    fn to_key_part(self) -> PivotKeyPart {
        match self {
            PivotValueRef::Borrowed(v) => v.to_key_part(),
            PivotValueRef::Owned(v) => match v {
                PivotValue::Blank => PivotKeyPart::Blank,
                PivotValue::Number(n) => PivotKeyPart::Number(PivotValue::canonical_number_bits(n)),
                PivotValue::Date(d) => PivotKeyPart::Date(d),
                PivotValue::Text(s) => PivotKeyPart::Text(s),
                PivotValue::Bool(b) => PivotKeyPart::Bool(b),
            },
        }
    }
}

/// Abstraction over pivot cache record storage.
///
/// This allows the pivot engine to aggregate large datasets without requiring a full
/// `Vec<Vec<PivotValue>>` materialization. Implementations can be backed by either the existing
/// row-wise [`PivotCache`] or a columnar store like [`ColumnarTable`].
pub trait PivotRecordSource {
    fn row_count(&self) -> usize;
    fn column_count(&self) -> usize;
    fn field_index(&self, field_name: &str) -> Option<usize>;
    fn value(&self, row: usize, col: usize) -> PivotValueRef<'_>;
    #[inline]
    fn as_pivot_cache(&self) -> Option<&PivotCache> {
        None
    }
}

impl PivotRecordSource for PivotCache {
    fn row_count(&self) -> usize {
        self.records.len()
    }

    fn column_count(&self) -> usize {
        self.fields.len()
    }

    fn field_index(&self, field_name: &str) -> Option<usize> {
        PivotCache::field_index(self, field_name)
    }

    fn value(&self, row: usize, col: usize) -> PivotValueRef<'_> {
        self.records
            .get(row)
            .and_then(|r| r.get(col))
            .map(PivotValueRef::Borrowed)
            .unwrap_or_else(|| PivotValueRef::Owned(PivotValue::Blank))
    }

    fn as_pivot_cache(&self) -> Option<&PivotCache> {
        Some(self)
    }
}

fn columnar_value_to_pivot(value: ColumnarValue) -> PivotValue {
    match value {
        ColumnarValue::Null => PivotValue::Blank,
        ColumnarValue::Number(n) => PivotValue::Number(n),
        ColumnarValue::Boolean(b) => PivotValue::Bool(b),
        ColumnarValue::String(s) => PivotValue::Text(s.to_string()),
        // `ColumnarTable` stores datetime/currency/percentage as i64; render them as numbers for
        // now (matching the worksheet layer's default rendering).
        ColumnarValue::DateTime(v) | ColumnarValue::Currency(v) | ColumnarValue::Percentage(v) => {
            PivotValue::Number(v as f64)
        }
    }
}

impl PivotRecordSource for ColumnarTable {
    fn row_count(&self) -> usize {
        ColumnarTable::row_count(self)
    }

    fn column_count(&self) -> usize {
        ColumnarTable::column_count(self)
    }

    fn field_index(&self, field_name: &str) -> Option<usize> {
        self.schema()
            .iter()
            .position(|c| c.name.as_str() == field_name)
    }

    fn value(&self, row: usize, col: usize) -> PivotValueRef<'_> {
        PivotValueRef::Owned(columnar_value_to_pivot(self.get_cell(row, col)))
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotResult {
    pub data: Vec<Vec<PivotValue>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellRef {
    pub row: u32,
    pub col: u32,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellWrite {
    pub row: u32,
    pub col: u32,
    pub value: PivotValue,
    /// Optional number format code to apply when writing this cell to a worksheet.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
}

impl PivotResult {
    /// Converts the computed pivot into a list of worksheet cell writes.
    pub fn to_cell_writes(&self, destination: CellRef) -> Vec<CellWrite> {
        let mut out = Vec::new();
        for (r, row) in self.data.iter().enumerate() {
            for (c, value) in row.iter().enumerate() {
                out.push(CellWrite {
                    row: destination.row + r as u32,
                    col: destination.col + c as u32,
                    value: value.clone(),
                    number_format: None,
                });
            }
        }
        out
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CreatePivotValueSpec {
    pub field: String,
    pub aggregation: AggregationType,
    pub name: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CreatePivotFilterSpec {
    pub field: String,
    pub allowed: Option<Vec<PivotValue>>,
}

/// Request payload for an AI/tooling layer.
///
/// The orchestrator (or an LLM) can create a `CreatePivotTableRequest` from a
/// natural language prompt using the output of [`PivotCache::schema`]. The core
/// engine then validates field names at compute time (via `PivotEngine`).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CreatePivotTableRequest {
    pub name: Option<String>,
    pub row_fields: Vec<String>,
    pub column_fields: Vec<String>,
    pub value_fields: Vec<CreatePivotValueSpec>,
    pub filter_fields: Vec<CreatePivotFilterSpec>,
    pub calculated_fields: Option<Vec<CalculatedField>>,
    pub calculated_items: Option<Vec<CalculatedItem>>,
    pub layout: Option<Layout>,
    pub subtotals: Option<SubtotalPosition>,
    pub grand_totals: Option<GrandTotals>,
}

impl CreatePivotTableRequest {
    pub fn into_config(self) -> PivotConfig {
        PivotConfig {
            row_fields: self
                .row_fields
                .into_iter()
                .map(|field| PivotField::new(pivot_field_ref_from_legacy_string(field)))
                .collect(),
            column_fields: self
                .column_fields
                .into_iter()
                .map(|field| PivotField::new(pivot_field_ref_from_legacy_string(field)))
                .collect(),
            value_fields: self
                .value_fields
                .into_iter()
                .map(|vf| {
                    let field = vf.field;
                    let source_field = pivot_field_ref_from_legacy_string(field);
                    let aggregation = vf.aggregation;
                    let name = vf
                        .name
                        .unwrap_or_else(|| {
                            // For Data Model measures, prefer a human-friendly name without the
                            // DAX bracket syntax (Excel displays the measure as `Total Sales`, not
                            // `[Total Sales]`, in the default "Sum of ..." label).
                            let label = pivot_field_ref_caption(&source_field);
                            format!(
                                "{:?} of {}",
                                aggregation,
                                label
                            )
                        });
                    ValueField {
                        source_field,
                        name,
                        aggregation,
                        number_format: None,
                        show_as: None,
                        base_field: None,
                        base_item: None,
                    }
                })
                .collect(),
            filter_fields: self
                .filter_fields
                .into_iter()
                .map(|f| {
                    let CreatePivotFilterSpec { field, allowed } = f;
                    FilterField {
                        source_field: pivot_field_ref_from_legacy_string(field),
                        allowed: allowed
                            .map(|vals| vals.into_iter().map(|v| v.to_key_part()).collect()),
                    }
                })
                .collect(),
            calculated_fields: self.calculated_fields.unwrap_or_default(),
            calculated_items: self.calculated_items.unwrap_or_default(),
            layout: self.layout.unwrap_or(Layout::Tabular),
            subtotals: self.subtotals.unwrap_or(SubtotalPosition::None),
            grand_totals: self.grand_totals.unwrap_or_default(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotTable {
    pub id: String,
    pub name: String,
    pub config: PivotConfig,
    pub cache: PivotCache,
}

impl PivotTable {
    pub fn new(
        name: impl Into<String>,
        source: &[Vec<PivotValue>],
        config: PivotConfig,
    ) -> Result<Self, PivotError> {
        Ok(Self {
            id: next_pivot_id(),
            name: name.into(),
            config,
            cache: PivotCache::from_range(source)?,
        })
    }

    pub fn refresh_cache(&mut self, source: &[Vec<PivotValue>]) -> Result<(), PivotError> {
        self.cache.refresh_from_range(source)
    }

    pub fn calculate(&self) -> Result<PivotResult, PivotError> {
        PivotEngine::calculate(&self.cache, &self.config)
    }
}

fn next_pivot_id() -> String {
    static NEXT_ID: AtomicU64 = AtomicU64::new(1);
    format!("pivot-{}", NEXT_ID.fetch_add(1, AtomicOrdering::Relaxed))
}

#[derive(Debug, Clone)]
struct Accumulator {
    count: u64,
    count_numbers: u64,
    sum: f64,
    product: f64,
    min: f64,
    max: f64,
    mean: f64,
    m2: f64,
}

impl Accumulator {
    fn new() -> Self {
        Self {
            count: 0,
            count_numbers: 0,
            sum: 0.0,
            product: 1.0,
            min: f64::INFINITY,
            max: f64::NEG_INFINITY,
            mean: 0.0,
            m2: 0.0,
        }
    }

    fn update(&mut self, value: &PivotValue) {
        if !value.is_blank() {
            self.count += 1;
        }
        if let Some(x) = value.as_number() {
            self.count_numbers += 1;
            self.sum += x;
            self.product *= x;
            if x < self.min {
                self.min = x;
            }
            if x > self.max {
                self.max = x;
            }

            // Welford variance
            let n = self.count_numbers as f64;
            let delta = x - self.mean;
            self.mean += delta / n;
            let delta2 = x - self.mean;
            self.m2 += delta * delta2;
        }
    }

    fn merge(&mut self, other: &Accumulator) {
        // Merge counts that include non-blank values.
        self.count += other.count;

        if other.count_numbers == 0 {
            return;
        }

        if self.count_numbers == 0 {
            *self = other.clone();
            return;
        }

        // Merge numeric aggregates.
        let n1 = self.count_numbers as f64;
        let n2 = other.count_numbers as f64;
        let n = n1 + n2;
        let delta = other.mean - self.mean;

        self.sum += other.sum;
        self.product *= other.product;
        self.min = self.min.min(other.min);
        self.max = self.max.max(other.max);

        self.mean = (n1 * self.mean + n2 * other.mean) / n;
        self.m2 += other.m2 + delta * delta * (n1 * n2) / n;
        self.count_numbers += other.count_numbers;
    }

    fn finalize(&self, agg: AggregationType) -> PivotValue {
        match agg {
            AggregationType::Count => PivotValue::Number(self.count as f64),
            AggregationType::CountNumbers => PivotValue::Number(self.count_numbers as f64),
            AggregationType::Sum => PivotValue::Number(self.sum),
            AggregationType::Product => {
                if self.count_numbers == 0 {
                    PivotValue::Blank
                } else {
                    PivotValue::Number(self.product)
                }
            }
            AggregationType::Average => {
                if self.count_numbers == 0 {
                    PivotValue::Blank
                } else {
                    PivotValue::Number(self.sum / self.count_numbers as f64)
                }
            }
            AggregationType::Min => {
                if self.count_numbers == 0 {
                    PivotValue::Blank
                } else {
                    PivotValue::Number(self.min)
                }
            }
            AggregationType::Max => {
                if self.count_numbers == 0 {
                    PivotValue::Blank
                } else {
                    PivotValue::Number(self.max)
                }
            }
            AggregationType::Var => {
                if self.count_numbers < 2 {
                    PivotValue::Blank
                } else {
                    PivotValue::Number(self.m2 / (self.count_numbers as f64 - 1.0))
                }
            }
            AggregationType::VarP => {
                if self.count_numbers == 0 {
                    PivotValue::Blank
                } else {
                    PivotValue::Number(self.m2 / (self.count_numbers as f64))
                }
            }
            AggregationType::StdDev => {
                if self.count_numbers < 2 {
                    PivotValue::Blank
                } else {
                    PivotValue::Number((self.m2 / (self.count_numbers as f64 - 1.0)).sqrt())
                }
            }
            AggregationType::StdDevP => {
                if self.count_numbers == 0 {
                    PivotValue::Blank
                } else {
                    PivotValue::Number((self.m2 / (self.count_numbers as f64)).sqrt())
                }
            }
        }
    }
}

fn normalize_pivot_item_name(name: &str) -> String {
    fold_text_case_insensitive(name)
}

fn fold_text_case_insensitive(s: &str) -> String {
    if s.is_ascii() {
        s.to_ascii_uppercase()
    } else {
        // Use Unicode-aware uppercasing to approximate Excel-like case-insensitive matching for
        // non-ASCII text (e.g. ß -> SS).
        s.chars().flat_map(|c| c.to_uppercase()).collect()
    }
}

#[derive(Debug, Clone)]
struct FieldItemResolver {
    /// Normalized display string -> concrete key part.
    by_name: HashMap<String, PivotKeyPart>,
    /// Normalized display strings that map to multiple distinct key parts.
    ambiguous: HashSet<String>,
}

impl FieldItemResolver {
    fn from_cache(cache: &PivotCache, field: &str) -> Self {
        let mut by_name: HashMap<String, PivotKeyPart> = HashMap::new();
        let mut ambiguous: HashSet<String> = HashSet::new();

        if let Some(values) = cache.unique_values.get(field) {
            for value in values {
                let part = value.to_key_part();
                let display = part.display_string();
                let normalized = normalize_pivot_item_name(&display);

                if ambiguous.contains(&normalized) {
                    continue;
                }

                match by_name.get(&normalized) {
                    None => {
                        by_name.insert(normalized, part);
                    }
                    Some(existing) if existing == &part => {
                        // Same underlying item; ignore.
                    }
                    Some(_) => {
                        // Multiple distinct key parts share the same display string; item references would be
                        // ambiguous, so mark the name as ambiguous.
                        by_name.remove(&normalized);
                        ambiguous.insert(normalized);
                    }
                }
            }
        }

        Self { by_name, ambiguous }
    }

    fn contains_display_name(&self, name: &str) -> bool {
        let normalized = normalize_pivot_item_name(name);
        self.by_name.contains_key(&normalized) || self.ambiguous.contains(&normalized)
    }

    fn insert_calculated_item(&mut self, name: &str) {
        let normalized = normalize_pivot_item_name(name);
        self.by_name
            .insert(normalized, PivotKeyPart::Text(name.to_string()));
    }

    fn resolve(&self, name: &str) -> Result<PivotKeyPart, String> {
        let normalized = normalize_pivot_item_name(name);
        if self.ambiguous.contains(&normalized) {
            return Err(format!(
                "item reference \"{name}\" is ambiguous (multiple distinct items share that display name)"
            ));
        }
        self.by_name
            .get(&normalized)
            .cloned()
            .ok_or_else(|| format!("unknown item reference \"{name}\""))
    }
}

/// Expression grammar for calculated items.
///
/// Supported syntax (MVP):
/// - Item references are double-quoted display strings, e.g. `"East"` or `"(blank)"`.
/// - Numeric literals (floating point) like `10` or `3.5`.
/// - Operators: `+`, `-`, `*`, `/` with standard precedence.
/// - Parentheses for grouping, e.g. `("East" + "West") / 2`.
#[derive(Debug, Clone, PartialEq)]
enum CalculatedItemExprRaw {
    Number(f64),
    Item(String),
    Neg(Box<CalculatedItemExprRaw>),
    Add(Box<CalculatedItemExprRaw>, Box<CalculatedItemExprRaw>),
    Sub(Box<CalculatedItemExprRaw>, Box<CalculatedItemExprRaw>),
    Mul(Box<CalculatedItemExprRaw>, Box<CalculatedItemExprRaw>),
    Div(Box<CalculatedItemExprRaw>, Box<CalculatedItemExprRaw>),
}

#[derive(Debug, Clone, PartialEq)]
enum CalculatedItemExpr {
    Number(f64),
    Item(PivotKeyPart),
    Neg(Box<CalculatedItemExpr>),
    Add(Box<CalculatedItemExpr>, Box<CalculatedItemExpr>),
    Sub(Box<CalculatedItemExpr>, Box<CalculatedItemExpr>),
    Mul(Box<CalculatedItemExpr>, Box<CalculatedItemExpr>),
    Div(Box<CalculatedItemExpr>, Box<CalculatedItemExpr>),
}

impl CalculatedItemExprRaw {
    fn resolve(self, resolver: &FieldItemResolver) -> Result<CalculatedItemExpr, String> {
        Ok(match self {
            CalculatedItemExprRaw::Number(n) => CalculatedItemExpr::Number(n),
            CalculatedItemExprRaw::Item(name) => CalculatedItemExpr::Item(resolver.resolve(&name)?),
            CalculatedItemExprRaw::Neg(expr) => {
                CalculatedItemExpr::Neg(Box::new(expr.resolve(resolver)?))
            }
            CalculatedItemExprRaw::Add(a, b) => CalculatedItemExpr::Add(
                Box::new(a.resolve(resolver)?),
                Box::new(b.resolve(resolver)?),
            ),
            CalculatedItemExprRaw::Sub(a, b) => CalculatedItemExpr::Sub(
                Box::new(a.resolve(resolver)?),
                Box::new(b.resolve(resolver)?),
            ),
            CalculatedItemExprRaw::Mul(a, b) => CalculatedItemExpr::Mul(
                Box::new(a.resolve(resolver)?),
                Box::new(b.resolve(resolver)?),
            ),
            CalculatedItemExprRaw::Div(a, b) => CalculatedItemExpr::Div(
                Box::new(a.resolve(resolver)?),
                Box::new(b.resolve(resolver)?),
            ),
        })
    }
}

impl CalculatedItemExpr {
    fn eval<F>(&self, lookup: &F) -> Result<f64, String>
    where
        F: Fn(&PivotKeyPart) -> f64,
    {
        Ok(match self {
            CalculatedItemExpr::Number(n) => *n,
            CalculatedItemExpr::Item(part) => lookup(part),
            CalculatedItemExpr::Neg(expr) => -expr.eval(lookup)?,
            CalculatedItemExpr::Add(a, b) => a.eval(lookup)? + b.eval(lookup)?,
            CalculatedItemExpr::Sub(a, b) => a.eval(lookup)? - b.eval(lookup)?,
            CalculatedItemExpr::Mul(a, b) => a.eval(lookup)? * b.eval(lookup)?,
            CalculatedItemExpr::Div(a, b) => {
                let denom = b.eval(lookup)?;
                if denom == 0.0 {
                    return Err("division by zero".to_string());
                }
                a.eval(lookup)? / denom
            }
        })
    }
}

struct CalculatedItemParser<'a> {
    input: &'a [u8],
    pos: usize,
}

impl<'a> CalculatedItemParser<'a> {
    fn parse(formula: &'a str) -> Result<CalculatedItemExprRaw, String> {
        let mut parser = Self {
            input: formula.as_bytes(),
            pos: 0,
        };
        let expr = parser.parse_expr()?;
        parser.skip_ws();
        if parser.pos < parser.input.len() {
            return Err(format!(
                "unexpected token at position {}",
                parser.pos.saturating_add(1)
            ));
        }
        Ok(expr)
    }

    fn skip_ws(&mut self) {
        while let Some(b) = self.peek() {
            if b.is_ascii_whitespace() {
                self.pos += 1;
            } else {
                break;
            }
        }
    }

    fn peek(&self) -> Option<u8> {
        self.input.get(self.pos).copied()
    }

    fn consume(&mut self) -> Option<u8> {
        let b = self.peek()?;
        self.pos += 1;
        Some(b)
    }

    fn parse_expr(&mut self) -> Result<CalculatedItemExprRaw, String> {
        let mut left = self.parse_term()?;
        loop {
            self.skip_ws();
            match self.peek() {
                Some(b'+') => {
                    self.pos += 1;
                    let right = self.parse_term()?;
                    left = CalculatedItemExprRaw::Add(Box::new(left), Box::new(right));
                }
                Some(b'-') => {
                    self.pos += 1;
                    let right = self.parse_term()?;
                    left = CalculatedItemExprRaw::Sub(Box::new(left), Box::new(right));
                }
                _ => break,
            }
        }
        Ok(left)
    }

    fn parse_term(&mut self) -> Result<CalculatedItemExprRaw, String> {
        let mut left = self.parse_factor()?;
        loop {
            self.skip_ws();
            match self.peek() {
                Some(b'*') => {
                    self.pos += 1;
                    let right = self.parse_factor()?;
                    left = CalculatedItemExprRaw::Mul(Box::new(left), Box::new(right));
                }
                Some(b'/') => {
                    self.pos += 1;
                    let right = self.parse_factor()?;
                    left = CalculatedItemExprRaw::Div(Box::new(left), Box::new(right));
                }
                _ => break,
            }
        }
        Ok(left)
    }

    fn parse_factor(&mut self) -> Result<CalculatedItemExprRaw, String> {
        self.skip_ws();
        match self.peek() {
            Some(b'-') => {
                self.pos += 1;
                let inner = self.parse_factor()?;
                Ok(CalculatedItemExprRaw::Neg(Box::new(inner)))
            }
            Some(b'(') => {
                self.pos += 1;
                let expr = self.parse_expr()?;
                self.skip_ws();
                match self.consume() {
                    Some(b')') => Ok(expr),
                    _ => Err("expected ')'".to_string()),
                }
            }
            Some(b'"') => Ok(CalculatedItemExprRaw::Item(self.parse_string()?)),
            Some(b'.') | Some(b'0'..=b'9') => {
                Ok(CalculatedItemExprRaw::Number(self.parse_number()?))
            }
            _ => Err("expected number, item reference, or '('".to_string()),
        }
    }

    fn parse_string(&mut self) -> Result<String, String> {
        match self.consume() {
            Some(b'"') => {}
            _ => return Err("expected '\"'".to_string()),
        }

        let mut out = String::new();
        while let Some(b) = self.consume() {
            match b {
                b'"' => return Ok(out),
                b'\\' => {
                    let escaped = self
                        .consume()
                        .ok_or_else(|| "unterminated escape sequence".to_string())?;
                    match escaped {
                        b'"' => out.push('"'),
                        b'\\' => out.push('\\'),
                        b'n' => out.push('\n'),
                        b't' => out.push('\t'),
                        other => {
                            return Err(format!("unsupported escape sequence: \\{}", other as char))
                        }
                    }
                }
                other => out.push(other as char),
            }
        }

        Err("unterminated string literal".to_string())
    }

    fn parse_number(&mut self) -> Result<f64, String> {
        let start = self.pos;
        // Integer part.
        while matches!(self.peek(), Some(b'0'..=b'9')) {
            self.pos += 1;
        }
        // Fractional part.
        if self.peek() == Some(b'.') {
            self.pos += 1;
            while matches!(self.peek(), Some(b'0'..=b'9')) {
                self.pos += 1;
            }
        }
        // Exponent part.
        if matches!(self.peek(), Some(b'e' | b'E')) {
            let exp_start = self.pos;
            self.pos += 1;
            if matches!(self.peek(), Some(b'+' | b'-')) {
                self.pos += 1;
            }
            let digits_start = self.pos;
            while matches!(self.peek(), Some(b'0'..=b'9')) {
                self.pos += 1;
            }
            if digits_start == self.pos {
                // Roll back to before the exponent to produce a more helpful error below.
                self.pos = exp_start;
            }
        }

        let s = std::str::from_utf8(&self.input[start..self.pos])
            .map_err(|_| "invalid number literal".to_string())?;
        if s.is_empty() || s == "." {
            return Err("invalid number literal".to_string());
        }
        s.parse::<f64>()
            .map_err(|_| format!("invalid number literal: {s}"))
    }
}

fn synthetic_accumulator_from_value(
    value: f64,
    agg: AggregationType,
) -> Result<Accumulator, String> {
    if !value.is_finite() {
        return Err("calculated item evaluated to a non-finite number".to_string());
    }

    match agg {
        AggregationType::Count | AggregationType::CountNumbers => {
            let rounded = value.round();
            if (value - rounded).abs() > 1e-9 {
                return Err(format!(
                    "expected an integer result for {:?}, got {value}",
                    agg
                ));
            }
            if rounded < 0.0 {
                return Err(format!(
                    "expected a non-negative result for {:?}, got {value}",
                    agg
                ));
            }
            if rounded > u64::MAX as f64 {
                return Err(format!(
                    "result for {:?} is too large ({} > {})",
                    agg,
                    rounded,
                    u64::MAX
                ));
            }

            let n = rounded as u64;

            // Treat a calculated Count/CountNumbers result as `n` synthetic numeric observations with
            // value `0`. This keeps the accumulator merge logic stable and ensures the grand totals
            // include the calculated item.
            Ok(Accumulator {
                count: n,
                count_numbers: n,
                sum: 0.0,
                product: if n == 0 { 1.0 } else { 0.0 },
                min: if n == 0 { f64::INFINITY } else { 0.0 },
                max: if n == 0 { f64::NEG_INFINITY } else { 0.0 },
                mean: 0.0,
                m2: 0.0,
            })
        }
        AggregationType::Var => {
            if value < 0.0 {
                return Err(format!("variance cannot be negative: {value}"));
            }
            let a = (value / 2.0).sqrt();
            let mut acc = Accumulator::new();
            acc.update(&PivotValue::Number(a));
            acc.update(&PivotValue::Number(-a));
            Ok(acc)
        }
        AggregationType::StdDev => {
            if value < 0.0 {
                return Err(format!("standard deviation cannot be negative: {value}"));
            }
            // For two observations +/-a, sample stdev = |a| * sqrt(2).
            let a = value / 2.0_f64.sqrt();
            let mut acc = Accumulator::new();
            acc.update(&PivotValue::Number(a));
            acc.update(&PivotValue::Number(-a));
            Ok(acc)
        }
        AggregationType::VarP => {
            if value < 0.0 {
                return Err(format!("variance cannot be negative: {value}"));
            }
            // For two observations +/-a, population variance = a^2.
            let a = value.sqrt();
            let mut acc = Accumulator::new();
            acc.update(&PivotValue::Number(a));
            acc.update(&PivotValue::Number(-a));
            Ok(acc)
        }
        AggregationType::StdDevP => {
            if value < 0.0 {
                return Err(format!("standard deviation cannot be negative: {value}"));
            }
            // For two observations +/-a, population stdev = |a|.
            let a = value;
            let mut acc = Accumulator::new();
            acc.update(&PivotValue::Number(a));
            acc.update(&PivotValue::Number(-a));
            Ok(acc)
        }
        _ => {
            let mut acc = Accumulator::new();
            acc.update(&PivotValue::Number(value));
            Ok(acc)
        }
    }
}

pub struct PivotEngine;

#[derive(Debug, Clone, PartialEq, Eq)]
enum PivotRowKind {
    Header,
    Leaf {
        row_key_idx: usize,
    },
    /// A subtotal row for the prefix `prefix_key` (length = level + 1).
    Subtotal {
        level: usize,
        prefix_key: PivotKey,
    },
    GrandTotal,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CalculatedItemPlacement {
    Row(usize),
    Column(usize),
}

impl PivotEngine {
    pub fn calculate(cache: &PivotCache, cfg: &PivotConfig) -> Result<PivotResult, PivotError> {
        Self::calculate_streaming(cache, cfg)
    }

    /// Compute a pivot table by scanning a [`PivotRecordSource`].
    ///
    /// This enables a streaming/columnar-backed execution path for large datasets where
    /// materializing a `Vec<Vec<PivotValue>>` would be prohibitively expensive.
    pub fn calculate_streaming<S: PivotRecordSource + ?Sized>(
        source: &S,
        cfg: &PivotConfig,
    ) -> Result<PivotResult, PivotError> {
        if cfg.value_fields.is_empty() {
            return Err(PivotError::NoValueFields);
        }

        let indices = FieldIndices::new(source, cfg)?;

        let mut cube: HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>> = HashMap::new();
        let mut row_keys: HashSet<PivotKey> = HashSet::new();
        let mut col_keys: HashSet<PivotKey> = HashSet::new();

        for row in 0..source.row_count() {
            if !indices.passes_filters(source, row) {
                continue;
            }

            let row_key = indices.build_key(source, row, &indices.row_indices);
            let col_key = indices.build_key(source, row, &indices.col_indices);

            row_keys.insert(row_key.clone());
            col_keys.insert(col_key.clone());

            let row_entry = cube.entry(row_key).or_default();
            let cell = row_entry.entry(col_key).or_insert_with(|| {
                (0..cfg.value_fields.len())
                    .map(|_| Accumulator::new())
                    .collect()
            });

            for (vf_idx, _vf) in cfg.value_fields.iter().enumerate() {
                let val = source.value(row, indices.value_indices[vf_idx]);
                cell[vf_idx].update(val.as_value());
            }
        }

        if !cfg.calculated_items.is_empty() {
            let Some(cache) = source.as_pivot_cache() else {
                return Err(PivotError::CalculatedItemsRequirePivotCache);
            };
            Self::apply_calculated_items(cache, cfg, &mut cube, &mut row_keys, &mut col_keys)?;
        }

        let mut row_keys: Vec<PivotKey> = row_keys.into_iter().collect();
        let row_sort_specs = cfg
            .row_fields
            .iter()
            .map(KeySortSpec::for_field)
            .collect::<Vec<_>>();
        row_keys.sort_by(|a, b| compare_pivot_keys(a, b, &row_sort_specs));

        let mut col_keys: Vec<PivotKey> = col_keys.into_iter().collect();
        let col_sort_specs = cfg
            .column_fields
            .iter()
            .map(KeySortSpec::for_field)
            .collect::<Vec<_>>();
        col_keys.sort_by(|a, b| compare_pivot_keys(a, b, &col_sort_specs));

        // Ensure at least one column key exists to simplify output logic.
        if cfg.column_fields.is_empty() && col_keys.is_empty() {
            col_keys.push(PivotKey(Vec::new()));
        }

        let mut data = Vec::new();
        data.push(Self::build_header_row(&col_keys, cfg));
        let mut row_kinds = vec![PivotRowKind::Header];

        // Subtotal accumulators per row-field level (excluding leaf level).
        let subtotal_levels = cfg.row_fields.len().saturating_sub(1);
        let mut grand_acc: Option<GroupAccumulator> = if cfg.grand_totals.rows {
            Some(GroupAccumulator::new())
        } else {
            None
        };

        match cfg.subtotals {
            SubtotalPosition::Top if subtotal_levels > 0 => {
                // For top subtotals we need the totals up front, so we precompute them per prefix.
                let group_totals = Self::precompute_group_totals(
                    &cube,
                    &row_keys,
                    &col_keys,
                    cfg,
                    subtotal_levels,
                );

                let mut prev_row_key: Option<PivotKey> = None;
                for (row_key_idx, row_key) in row_keys.iter().enumerate() {
                    let common_prefix = prev_row_key
                        .as_ref()
                        .map(|prev| common_prefix_len(&prev.0, &row_key.0))
                        .unwrap_or(0);

                    // Open new groups for changed prefixes and emit their subtotal rows.
                    for level in common_prefix..subtotal_levels {
                        let prefix_key = PivotKey(row_key.0[..=level].to_vec());
                        if let Some(totals) = group_totals[level].get(&prefix_key) {
                            data.push(Self::render_subtotal_row(
                                level, &row_key.0, totals, &col_keys, cfg,
                            ));
                            row_kinds.push(PivotRowKind::Subtotal { level, prefix_key });
                        }
                    }

                    let row_map = cube.get(row_key);
                    data.push(Self::render_row(
                        row_key, row_map, &col_keys, cfg, /*label*/ None,
                    ));
                    row_kinds.push(PivotRowKind::Leaf { row_key_idx });

                    if let Some(acc) = grand_acc.as_mut() {
                        acc.merge_row(row_map, &col_keys, cfg.value_fields.len());
                    }

                    prev_row_key = Some(row_key.clone());
                }
            }
            SubtotalPosition::Bottom if subtotal_levels > 0 => {
                let mut group_accs: Vec<Option<GroupAccumulator>> = vec![None; subtotal_levels];

                let mut prev_row_key: Option<PivotKey> = None;
                for (row_key_idx, row_key) in row_keys.iter().enumerate() {
                    let common_prefix = prev_row_key
                        .as_ref()
                        .map(|prev| common_prefix_len(&prev.0, &row_key.0))
                        .unwrap_or(0);

                    // Close groups if needed before emitting leaf row.
                    if let Some(prev) = prev_row_key.as_ref() {
                        Self::close_groups_bottom(
                            cfg,
                            &col_keys,
                            common_prefix,
                            &mut group_accs,
                            prev,
                            &mut data,
                            &mut row_kinds,
                        );
                    }

                    // Open new groups for changed prefixes.
                    for level in common_prefix..subtotal_levels {
                        group_accs[level] = Some(GroupAccumulator::new());
                    }

                    let row_map = cube.get(row_key);
                    data.push(Self::render_row(
                        row_key, row_map, &col_keys, cfg, /*label*/ None,
                    ));
                    row_kinds.push(PivotRowKind::Leaf { row_key_idx });

                    // Update subtotal accumulators & grand accumulator.
                    for level in 0..subtotal_levels {
                        if let Some(acc) = group_accs[level].as_mut() {
                            acc.merge_row(row_map, &col_keys, cfg.value_fields.len());
                        }
                    }
                    if let Some(acc) = grand_acc.as_mut() {
                        acc.merge_row(row_map, &col_keys, cfg.value_fields.len());
                    }

                    prev_row_key = Some(row_key.clone());
                }

                // Close remaining groups.
                if let Some(prev) = prev_row_key.as_ref() {
                    Self::close_groups_bottom(
                        cfg,
                        &col_keys,
                        0,
                        &mut group_accs,
                        prev,
                        &mut data,
                        &mut row_kinds,
                    );
                }
            }
            _ => {
                // No subtotals (or not enough row fields).
                for (row_key_idx, row_key) in row_keys.iter().enumerate() {
                    let row_map = cube.get(row_key);
                    data.push(Self::render_row(
                        row_key, row_map, &col_keys, cfg, /*label*/ None,
                    ));
                    row_kinds.push(PivotRowKind::Leaf { row_key_idx });
                    if let Some(acc) = grand_acc.as_mut() {
                        acc.merge_row(row_map, &col_keys, cfg.value_fields.len());
                    }
                }
            }
        }

        // Grand total row.
        if let Some(grand) = grand_acc {
            data.push(Self::render_totals_row(
                PivotValue::Text("Grand Total".to_string()),
                /*label_column*/ 0,
                /*prefix_parts*/ &[],
                &grand,
                &col_keys,
                cfg,
            ));
            row_kinds.push(PivotRowKind::GrandTotal);
        }

        if cfg
            .value_fields
            .iter()
            .any(|vf| vf.show_as.unwrap_or(ShowAsType::Normal) != ShowAsType::Normal)
        {
            Self::apply_show_as(&mut data, &row_kinds, &cube, &row_keys, &col_keys, cfg);
        }

        Ok(PivotResult { data })
    }

    fn calculated_item_placement(
        cfg: &PivotConfig,
        field: &str,
    ) -> Option<CalculatedItemPlacement> {
        let field = field.trim();
        if field.is_empty() {
            return None;
        }

        cfg.row_fields
            .iter()
            .position(|f| f.source_field.as_cache_field_name() == Some(field))
            .map(CalculatedItemPlacement::Row)
            .or_else(|| {
                cfg.column_fields
                    .iter()
                    .position(|f| f.source_field.as_cache_field_name() == Some(field))
                    .map(CalculatedItemPlacement::Column)
            })
    }

    fn apply_calculated_items(
        cache: &PivotCache,
        cfg: &PivotConfig,
        cube: &mut HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_keys: &mut HashSet<PivotKey>,
        col_keys: &mut HashSet<PivotKey>,
    ) -> Result<(), PivotError> {
        if cfg.calculated_items.is_empty() {
            return Ok(());
        }

        let mut resolvers: HashMap<String, FieldItemResolver> = HashMap::new();

        for item in &cfg.calculated_items {
            let placement = Self::calculated_item_placement(cfg, &item.field)
                .ok_or_else(|| PivotError::CalculatedItemFieldNotInLayout(item.field.clone()))?;

            let resolver = resolvers
                .entry(item.field.clone())
                .or_insert_with(|| FieldItemResolver::from_cache(cache, &item.field));

            // Excel disallows calculated item captions colliding with existing item captions.
            if resolver.contains_display_name(&item.name) {
                return Err(PivotError::CalculatedItemNameConflictsWithExistingItem {
                    field: item.field.clone(),
                    item: item.name.clone(),
                });
            }

            let raw_expr = CalculatedItemParser::parse(&item.formula).map_err(|message| {
                PivotError::InvalidCalculatedItemFormula {
                    field: item.field.clone(),
                    item: item.name.clone(),
                    message,
                }
            })?;
            let expr = raw_expr.resolve(resolver).map_err(|message| {
                PivotError::InvalidCalculatedItemFormula {
                    field: item.field.clone(),
                    item: item.name.clone(),
                    message,
                }
            })?;

            match placement {
                CalculatedItemPlacement::Row(idx) => Self::apply_row_calculated_item(
                    cube,
                    row_keys,
                    col_keys,
                    cfg,
                    idx,
                    &item.field,
                    &item.name,
                    &expr,
                )?,
                CalculatedItemPlacement::Column(idx) => Self::apply_column_calculated_item(
                    cube,
                    row_keys,
                    col_keys,
                    cfg,
                    idx,
                    &item.field,
                    &item.name,
                    &expr,
                )?,
            }

            resolver.insert_calculated_item(&item.name);
        }

        Ok(())
    }

    fn apply_row_calculated_item(
        cube: &mut HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_keys: &mut HashSet<PivotKey>,
        col_keys: &HashSet<PivotKey>,
        cfg: &PivotConfig,
        row_field_idx: usize,
        field: &str,
        item_name: &str,
        expr: &CalculatedItemExpr,
    ) -> Result<(), PivotError> {
        let existing_row_keys: Vec<PivotKey> = row_keys.iter().cloned().collect();
        if existing_row_keys.is_empty() {
            return Ok(());
        }

        #[derive(Debug, Clone)]
        struct RowGroup {
            template: PivotKey,
            items: HashMap<PivotKeyPart, PivotKey>,
        }

        let mut groups: HashMap<PivotKey, RowGroup> = HashMap::new();
        for row_key in &existing_row_keys {
            let Some(item_part) = row_key.0.get(row_field_idx).cloned() else {
                continue;
            };
            let mut base_parts = row_key.0.clone();
            base_parts.remove(row_field_idx);
            let base_key = PivotKey(base_parts);
            let entry = groups.entry(base_key).or_insert_with(|| RowGroup {
                template: row_key.clone(),
                items: HashMap::new(),
            });
            entry.items.insert(item_part, row_key.clone());
        }

        let col_keys_vec: Vec<PivotKey> = col_keys.iter().cloned().collect();
        let new_part = PivotKeyPart::Text(item_name.to_string());

        let mut new_rows: Vec<(PivotKey, HashMap<PivotKey, Vec<Accumulator>>)> = Vec::new();

        for group in groups.values() {
            let mut new_row_key = group.template.clone();
            if row_field_idx < new_row_key.0.len() {
                new_row_key.0[row_field_idx] = new_part.clone();
            } else {
                // Should not happen with valid layout, but keep output deterministic.
                continue;
            }

            let mut new_row_map: HashMap<PivotKey, Vec<Accumulator>> = HashMap::new();

            for col_key in &col_keys_vec {
                let mut cell_accs: Vec<Accumulator> = Vec::with_capacity(cfg.value_fields.len());
                for (vf_idx, vf) in cfg.value_fields.iter().enumerate() {
                    let agg = vf.aggregation;
                    let lookup = |part: &PivotKeyPart| -> f64 {
                        let Some(src_row_key) = group.items.get(part) else {
                            return 0.0;
                        };
                        let Some(src_row_map) = cube.get(src_row_key) else {
                            return 0.0;
                        };
                        let Some(src_cell) = src_row_map.get(col_key) else {
                            return 0.0;
                        };
                        src_cell
                            .get(vf_idx)
                            .map(|acc| acc.finalize(agg).as_number().unwrap_or(0.0))
                            .unwrap_or(0.0)
                    };

                    let value = expr.eval(&lookup).map_err(|message| {
                        PivotError::InvalidCalculatedItemFormula {
                            field: field.to_string(),
                            item: item_name.to_string(),
                            message,
                        }
                    })?;

                    let acc = synthetic_accumulator_from_value(value, agg).map_err(|message| {
                        PivotError::InvalidCalculatedItemFormula {
                            field: field.to_string(),
                            item: item_name.to_string(),
                            message,
                        }
                    })?;
                    cell_accs.push(acc);
                }

                new_row_map.insert(col_key.clone(), cell_accs);
            }

            new_rows.push((new_row_key, new_row_map));
        }

        for (row_key, row_map) in new_rows {
            row_keys.insert(row_key.clone());
            cube.insert(row_key, row_map);
        }

        Ok(())
    }

    fn apply_column_calculated_item(
        cube: &mut HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_keys: &HashSet<PivotKey>,
        col_keys: &mut HashSet<PivotKey>,
        cfg: &PivotConfig,
        col_field_idx: usize,
        field: &str,
        item_name: &str,
        expr: &CalculatedItemExpr,
    ) -> Result<(), PivotError> {
        let existing_col_keys: Vec<PivotKey> = col_keys.iter().cloned().collect();
        if existing_col_keys.is_empty() {
            return Ok(());
        }

        #[derive(Debug, Clone)]
        struct ColGroup {
            template: PivotKey,
            items: HashMap<PivotKeyPart, PivotKey>,
        }

        let mut groups: HashMap<PivotKey, ColGroup> = HashMap::new();
        for col_key in &existing_col_keys {
            let Some(item_part) = col_key.0.get(col_field_idx).cloned() else {
                continue;
            };
            let mut base_parts = col_key.0.clone();
            base_parts.remove(col_field_idx);
            let base_key = PivotKey(base_parts);
            let entry = groups.entry(base_key).or_insert_with(|| ColGroup {
                template: col_key.clone(),
                items: HashMap::new(),
            });
            entry.items.insert(item_part, col_key.clone());
        }

        let row_keys_vec: Vec<PivotKey> = row_keys.iter().cloned().collect();
        let new_part = PivotKeyPart::Text(item_name.to_string());

        for group in groups.values() {
            let mut new_col_key = group.template.clone();
            if col_field_idx < new_col_key.0.len() {
                new_col_key.0[col_field_idx] = new_part.clone();
            } else {
                continue;
            }

            col_keys.insert(new_col_key.clone());

            for row_key in &row_keys_vec {
                let row_map = cube.entry(row_key.clone()).or_default();
                let mut cell_accs: Vec<Accumulator> = Vec::with_capacity(cfg.value_fields.len());
                for (vf_idx, vf) in cfg.value_fields.iter().enumerate() {
                    let agg = vf.aggregation;
                    let value = expr
                        .eval(&|part: &PivotKeyPart| -> f64 {
                            let Some(src_col_key) = group.items.get(part) else {
                                return 0.0;
                            };
                            let Some(src_cell) = row_map.get(src_col_key) else {
                                return 0.0;
                            };
                            src_cell
                                .get(vf_idx)
                                .map(|acc| acc.finalize(agg).as_number().unwrap_or(0.0))
                                .unwrap_or(0.0)
                        })
                        .map_err(|message| PivotError::InvalidCalculatedItemFormula {
                            field: field.to_string(),
                            item: item_name.to_string(),
                            message,
                        })?;

                    let acc = synthetic_accumulator_from_value(value, agg).map_err(|message| {
                        PivotError::InvalidCalculatedItemFormula {
                            field: field.to_string(),
                            item: item_name.to_string(),
                            message,
                        }
                    })?;
                    cell_accs.push(acc);
                }
                row_map.insert(new_col_key.clone(), cell_accs);
            }
        }

        Ok(())
    }

    fn build_header_row(col_keys: &[PivotKey], cfg: &PivotConfig) -> Vec<PivotValue> {
        let mut row = Vec::new();

        match cfg.layout {
            Layout::Compact => {
                row.push(PivotValue::Text("Row Labels".to_string()));
            }
            Layout::Outline | Layout::Tabular => {
                for f in &cfg.row_fields {
                    row.push(PivotValue::Text(
                        pivot_field_ref_name(&f.source_field).into_owned(),
                    ));
                }
            }
        }

        // Flatten column keys × value fields.
        //
        // Note: Excel renders pivot item labels (including column items) as typed cell values
        // (numbers/dates/bools). The current pivot output flattens column-field labels and value
        // field captions into a single header row of text (e.g. "A - Sum of Sales"). Emitting
        // typed column key values would require multi-row headers to avoid losing the value field
        // captions; that is handled in higher-level rendering layers.
        for col_key in col_keys {
            let col_label = if cfg.column_fields.is_empty() {
                String::new()
            } else {
                col_key
                    .display_strings()
                    .into_iter()
                    .filter(|s| !s.is_empty())
                    .collect::<Vec<_>>()
                    .join(" / ")
            };
            for vf in &cfg.value_fields {
                let base = if vf.name.is_empty() {
                    format!(
                        "{:?} of {}",
                        vf.aggregation,
                        pivot_field_ref_caption(&vf.source_field)
                    )
                } else {
                    vf.name.clone()
                };
                let header = if col_label.is_empty() {
                    base
                } else {
                    format!("{col_label} - {base}")
                };
                row.push(PivotValue::Text(header));
            }
        }

        if cfg.grand_totals.columns {
            for vf in &cfg.value_fields {
                let base = if vf.name.is_empty() {
                    format!(
                        "{:?} of {}",
                        vf.aggregation,
                        pivot_field_ref_caption(&vf.source_field)
                    )
                } else {
                    vf.name.clone()
                };
                row.push(PivotValue::Text(format!("Grand Total - {base}")));
            }
        }

        row
    }

    fn render_row(
        row_key: &PivotKey,
        row_map: Option<&HashMap<PivotKey, Vec<Accumulator>>>,
        col_keys: &[PivotKey],
        cfg: &PivotConfig,
        label: Option<PivotValue>,
    ) -> Vec<PivotValue> {
        let mut row = Vec::new();

        match cfg.layout {
            Layout::Compact => {
                if let Some(label) = label {
                    row.push(label);
                } else if row_key.0.len() == 1 {
                    // Preserve typed values when the compact layout includes only a single row
                    // field (Excel-like: the row label cell stores the underlying value rather than
                    // the formatted display string).
                    row.push(pivot_key_part_to_pivot_value(&row_key.0[0]));
                } else {
                    // Compact: join row keys into one cell.
                    let s = row_key
                        .display_strings()
                        .into_iter()
                        .filter(|s| !s.is_empty())
                        .collect::<Vec<_>>()
                        .join(" / ");
                    row.push(PivotValue::Text(s));
                }
            }
            Layout::Outline | Layout::Tabular => {
                for (idx, part) in row_key.0.iter().enumerate() {
                    if idx == 0 {
                        if let Some(l) = label.as_ref() {
                            row.push(l.clone());
                            continue;
                        }
                    }
                    row.push(pivot_key_part_to_pivot_value(part));
                }

                // If row key shorter than row_fields (shouldn't happen), pad.
                while row.len() < cfg.row_fields.len() {
                    row.push(PivotValue::Blank);
                }
            }
        }

        let mut row_total_accs: Vec<Accumulator> = (0..cfg.value_fields.len())
            .map(|_| Accumulator::new())
            .collect();

        for col_key in col_keys {
            let maybe_cell = row_map.and_then(|m| m.get(col_key));
            for (vf_idx, vf) in cfg.value_fields.iter().enumerate() {
                if let Some(cell_accs) = maybe_cell {
                    row_total_accs[vf_idx].merge(&cell_accs[vf_idx]);
                    row.push(cell_accs[vf_idx].finalize(vf.aggregation));
                } else {
                    row.push(PivotValue::Blank);
                }
            }
        }

        if cfg.grand_totals.columns {
            for (vf_idx, vf) in cfg.value_fields.iter().enumerate() {
                row.push(row_total_accs[vf_idx].finalize(vf.aggregation));
            }
        }

        row
    }

    fn render_totals_row(
        label: PivotValue,
        label_column: usize,
        prefix_parts: &[PivotKeyPart],
        totals: &GroupAccumulator,
        col_keys: &[PivotKey],
        cfg: &PivotConfig,
    ) -> Vec<PivotValue> {
        let mut row = Vec::new();

        match cfg.layout {
            Layout::Compact => {
                row.push(label);
            }
            Layout::Outline | Layout::Tabular => {
                if cfg.row_fields.is_empty() {
                    // Preserve previous behavior: still emit a label cell even if there are no row fields.
                    row.push(label);
                } else {
                    for idx in 0..cfg.row_fields.len() {
                        if idx < label_column {
                            if let Some(prefix) = prefix_parts.get(idx) {
                                row.push(pivot_key_part_to_pivot_value(prefix));
                            } else {
                                row.push(PivotValue::Blank);
                            }
                        } else if idx == label_column {
                            row.push(label.clone());
                        } else {
                            row.push(PivotValue::Blank);
                        }
                    }
                }
            }
        }

        let mut row_total_accs: Vec<Accumulator> = (0..cfg.value_fields.len())
            .map(|_| Accumulator::new())
            .collect();

        for col_key in col_keys {
            let cell_accs = totals.cells.get(col_key).cloned().unwrap_or_else(|| {
                (0..cfg.value_fields.len())
                    .map(|_| Accumulator::new())
                    .collect()
            });

            for (vf_idx, vf) in cfg.value_fields.iter().enumerate() {
                row_total_accs[vf_idx].merge(&cell_accs[vf_idx]);
                row.push(cell_accs[vf_idx].finalize(vf.aggregation));
            }
        }

        if cfg.grand_totals.columns {
            for (vf_idx, vf) in cfg.value_fields.iter().enumerate() {
                row.push(row_total_accs[vf_idx].finalize(vf.aggregation));
            }
        }

        row
    }

    fn render_subtotal_row(
        level: usize,
        row_key_parts: &[PivotKeyPart],
        totals: &GroupAccumulator,
        col_keys: &[PivotKey],
        cfg: &PivotConfig,
    ) -> Vec<PivotValue> {
        let base = row_key_parts
            .get(level)
            .map(|p| p.display_string())
            .unwrap_or_default();
        let label = if base.is_empty() {
            PivotValue::Text("Total".to_string())
        } else {
            PivotValue::Text(format!("{base} Total"))
        };

        Self::render_totals_row(label, level, row_key_parts, totals, col_keys, cfg)
    }

    fn precompute_group_totals(
        cube: &HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_keys: &[PivotKey],
        col_keys: &[PivotKey],
        cfg: &PivotConfig,
        subtotal_levels: usize,
    ) -> Vec<HashMap<PivotKey, GroupAccumulator>> {
        let mut out: Vec<HashMap<PivotKey, GroupAccumulator>> =
            (0..subtotal_levels).map(|_| HashMap::new()).collect();

        for row_key in row_keys {
            let row_map = cube.get(row_key);
            for level in 0..subtotal_levels {
                let prefix_key = PivotKey(row_key.0[..=level].to_vec());
                let entry = out[level]
                    .entry(prefix_key)
                    .or_insert_with(GroupAccumulator::new);
                entry.merge_row(row_map, col_keys, cfg.value_fields.len());
            }
        }

        out
    }

    fn close_groups_bottom(
        cfg: &PivotConfig,
        col_keys: &[PivotKey],
        keep_prefix_len: usize,
        group_accs: &mut [Option<GroupAccumulator>],
        prev_row_key: &PivotKey,
        out: &mut Vec<Vec<PivotValue>>,
        row_kinds: &mut Vec<PivotRowKind>,
    ) {
        // Close from deepest to keep_prefix_len.
        for level in (keep_prefix_len..group_accs.len()).rev() {
            if let Some(acc) = group_accs[level].take() {
                out.push(Self::render_subtotal_row(
                    level,
                    &prev_row_key.0,
                    &acc,
                    col_keys,
                    cfg,
                ));
                let prefix_key = PivotKey(prev_row_key.0[..=level].to_vec());
                row_kinds.push(PivotRowKind::Subtotal { level, prefix_key });
            }
        }
    }

    fn apply_show_as(
        data: &mut [Vec<PivotValue>],
        row_kinds: &[PivotRowKind],
        cube: &HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_keys: &[PivotKey],
        col_keys: &[PivotKey],
        cfg: &PivotConfig,
    ) {
        if data.len() <= 1 {
            return;
        }

        let value_field_count = cfg.value_fields.len();
        if value_field_count == 0 {
            return;
        }

        let row_label_width = match cfg.layout {
            Layout::Compact => 1,
            Layout::Outline | Layout::Tabular => cfg.row_fields.len(),
        };

        let regular_column_count = col_keys.len() * value_field_count;
        let row_grand_total_start = row_label_width + regular_column_count;

        let leaf_rows: Vec<(usize, usize)> = row_kinds
            .iter()
            .enumerate()
            .filter_map(|(idx, kind)| match kind {
                PivotRowKind::Leaf { row_key_idx } => Some((idx, *row_key_idx)),
                _ => None,
            })
            .collect();

        let leaf_row_indices: Vec<usize> = leaf_rows.iter().map(|(r, _)| *r).collect();

        let subtotal_rows: Vec<(usize, &PivotKey)> = row_kinds
            .iter()
            .enumerate()
            .filter_map(|(idx, kind)| match kind {
                PivotRowKind::Subtotal { prefix_key, .. } => Some((idx, prefix_key)),
                _ => None,
            })
            .collect();

        let grand_total_row = row_kinds
            .iter()
            .enumerate()
            .find_map(|(idx, kind)| matches!(kind, PivotRowKind::GrandTotal).then_some(idx));

        for vf_idx in 0..value_field_count {
            let show_as = cfg.value_fields[vf_idx]
                .show_as
                .unwrap_or(ShowAsType::Normal);
            if show_as == ShowAsType::Normal {
                continue;
            }

            // All output columns that correspond to this value field:
            // - each column key has `value_field_count` columns (one per value field)
            // - followed by an optional row grand total section (also one per value field)
            let mut cols =
                Vec::with_capacity(col_keys.len() + usize::from(cfg.grand_totals.columns));
            for col_idx in 0..col_keys.len() {
                cols.push(row_label_width + col_idx * value_field_count + vf_idx);
            }
            if cfg.grand_totals.columns {
                cols.push(row_grand_total_start + vf_idx);
            }

            let agg = cfg.value_fields[vf_idx].aggregation;
            match show_as {
                ShowAsType::PercentOfGrandTotal => {
                    let denom = Self::compute_grand_total(cube, row_keys, col_keys, vf_idx, agg);
                    Self::apply_percent_of_total(data, &cols, denom);
                }
                ShowAsType::PercentOfRowTotal => {
                    let row_denoms = Self::compute_row_totals_for_show_as(
                        data,
                        row_label_width,
                        col_keys.len(),
                        value_field_count,
                        vf_idx,
                        cfg.grand_totals.columns,
                    );
                    Self::apply_percent_of_row_total(data, &cols, &row_denoms);
                }
                ShowAsType::PercentOfColumnTotal => {
                    let col_denoms =
                        Self::compute_column_totals(cube, row_keys, col_keys, vf_idx, agg);
                    let grand_denom =
                        Self::compute_grand_total(cube, row_keys, col_keys, vf_idx, agg);
                    Self::apply_percent_of_column_total(
                        data,
                        row_label_width,
                        col_keys.len(),
                        value_field_count,
                        vf_idx,
                        cfg.grand_totals.columns,
                        &col_denoms,
                        grand_denom,
                    );
                }
                ShowAsType::RunningTotal => {
                    if let Some(base_field) = cfg.value_fields[vf_idx].base_field.as_ref() {
                        if let Some(base_row_pos) = cfg
                            .row_fields
                            .iter()
                            .position(|f| &f.source_field == base_field)
                        {
                            let (group_ids, group_count) =
                                Self::group_ids_excluding_pos(row_keys, base_row_pos);
                            Self::apply_running_total_grouped_by_row(
                                data,
                                &leaf_rows,
                                &cols,
                                &group_ids,
                                group_count,
                            );
                        } else if let Some(base_col_pos) = cfg
                            .column_fields
                            .iter()
                            .position(|f| &f.source_field == base_field)
                        {
                            let (group_ids, group_count) =
                                Self::group_ids_excluding_pos(col_keys, base_col_pos);
                            Self::apply_running_total_grouped_by_column(
                                data,
                                &leaf_row_indices,
                                &cols[..col_keys.len()],
                                &group_ids,
                                group_count,
                            );
                        } else {
                            Self::apply_running_total(data, &leaf_row_indices, &cols);
                        }
                    } else {
                        Self::apply_running_total(data, &leaf_row_indices, &cols);
                    }
                }
                ShowAsType::RankAscending | ShowAsType::RankDescending => {
                    let descending = show_as == ShowAsType::RankDescending;
                    if let Some(base_field) = cfg.value_fields[vf_idx].base_field.as_ref() {
                        if let Some(base_row_pos) = cfg
                            .row_fields
                            .iter()
                            .position(|f| &f.source_field == base_field)
                        {
                            let (group_ids, group_count) =
                                Self::group_ids_excluding_pos(row_keys, base_row_pos);
                            Self::apply_rank_grouped_by_row(
                                data,
                                &leaf_rows,
                                &cols,
                                &group_ids,
                                group_count,
                                descending,
                            );
                        } else if let Some(base_col_pos) = cfg
                            .column_fields
                            .iter()
                            .position(|f| &f.source_field == base_field)
                        {
                            let (group_ids, group_count) =
                                Self::group_ids_excluding_pos(col_keys, base_col_pos);
                            Self::apply_rank_grouped_by_column(
                                data,
                                &leaf_row_indices,
                                &cols[..col_keys.len()],
                                &group_ids,
                                group_count,
                                descending,
                            );
                        } else {
                            Self::apply_rank(data, &leaf_row_indices, &cols, descending);
                        }
                    } else {
                        Self::apply_rank(data, &leaf_row_indices, &cols, descending);
                    }
                }
                ShowAsType::PercentOf | ShowAsType::PercentDifferenceFrom => {
                    // Base item semantics:
                    // - Treat the "base value" as the cell value at the same row key / column key,
                    //   except the axis corresponding to `base_field` is replaced with `base_item`.
                    // - If `base_field`/`base_item` is missing or invalid, blank affected cells.
                    // - If the base value is blank or 0, the output is blank.
                    let Some(base_field) = cfg.value_fields[vf_idx].base_field.as_ref() else {
                        Self::blank_numeric_cells(data, &cols);
                        continue;
                    };
                    let Some(base_item) = cfg.value_fields[vf_idx].base_item.as_deref() else {
                        Self::blank_numeric_cells(data, &cols);
                        continue;
                    };
                    let difference = show_as == ShowAsType::PercentDifferenceFrom;

                    let row_total_col = cfg
                        .grand_totals
                        .columns
                        .then_some(row_grand_total_start + vf_idx);

                    if let Some(base_row_pos) = cfg
                        .row_fields
                        .iter()
                        .position(|f| &f.source_field == base_field)
                    {
                        let Some(base_part) = row_keys.iter().find_map(|rk| {
                            rk.0.get(base_row_pos)
                                .filter(|p| p.display_string() == base_item)
                                .cloned()
                        }) else {
                            Self::blank_numeric_cells(data, &cols);
                            continue;
                        };

                        Self::apply_percent_of_base_item_row_field(
                            data,
                            &leaf_rows,
                            &subtotal_rows,
                            grand_total_row,
                            cube,
                            row_keys,
                            col_keys,
                            &cols[..col_keys.len()],
                            row_total_col,
                            vf_idx,
                            agg,
                            base_row_pos,
                            &base_part,
                            difference,
                        );
                    } else if let Some(base_col_pos) = cfg
                        .column_fields
                        .iter()
                        .position(|f| &f.source_field == base_field)
                    {
                        let Some(base_part) = col_keys.iter().find_map(|ck| {
                            ck.0.get(base_col_pos)
                                .filter(|p| p.display_string() == base_item)
                                .cloned()
                        }) else {
                            Self::blank_numeric_cells(data, &cols);
                            continue;
                        };

                        Self::apply_percent_of_base_item_column_field(
                            data,
                            &leaf_rows,
                            &subtotal_rows,
                            grand_total_row,
                            cube,
                            row_keys,
                            col_keys,
                            &cols[..col_keys.len()],
                            row_total_col,
                            vf_idx,
                            agg,
                            base_col_pos,
                            &base_part,
                            difference,
                        );
                    } else {
                        Self::blank_numeric_cells(data, &cols);
                    }
                }
                ShowAsType::Normal => {}
            }
        }
    }

    fn blank_numeric_cells(data: &mut [Vec<PivotValue>], cols: &[usize]) {
        for row in data.iter_mut().skip(1) {
            for &c in cols {
                if let Some(v) = row.get_mut(c) {
                    if matches!(v, PivotValue::Number(_)) {
                        *v = PivotValue::Blank;
                    }
                }
            }
        }
    }

    fn compute_grand_total(
        cube: &HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_keys: &[PivotKey],
        col_keys: &[PivotKey],
        value_field_idx: usize,
        agg: AggregationType,
    ) -> Option<f64> {
        let mut acc = Accumulator::new();
        for row_key in row_keys {
            let Some(row_map) = cube.get(row_key) else {
                continue;
            };
            for col_key in col_keys {
                if let Some(cell_accs) = row_map.get(col_key) {
                    acc.merge(&cell_accs[value_field_idx]);
                }
            }
        }
        acc.finalize(agg).as_number()
    }

    fn compute_column_totals(
        cube: &HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_keys: &[PivotKey],
        col_keys: &[PivotKey],
        value_field_idx: usize,
        agg: AggregationType,
    ) -> Vec<Option<f64>> {
        let mut out = Vec::with_capacity(col_keys.len());
        for col_key in col_keys {
            let mut acc = Accumulator::new();
            for row_key in row_keys {
                let Some(row_map) = cube.get(row_key) else {
                    continue;
                };
                if let Some(cell_accs) = row_map.get(col_key) {
                    acc.merge(&cell_accs[value_field_idx]);
                }
            }
            out.push(acc.finalize(agg).as_number());
        }
        out
    }

    fn apply_percent_of_total(data: &mut [Vec<PivotValue>], cols: &[usize], denom: Option<f64>) {
        let Some(denom) = denom.filter(|d| *d != 0.0) else {
            for r in 1..data.len() {
                for &c in cols {
                    if matches!(data[r].get(c), Some(PivotValue::Number(_))) {
                        data[r][c] = PivotValue::Blank;
                    }
                }
            }
            return;
        };

        for r in 1..data.len() {
            for &c in cols {
                if let Some(n) = data[r][c].as_number() {
                    data[r][c] = PivotValue::Number(n / denom);
                }
            }
        }
    }

    fn compute_row_totals_for_show_as(
        data: &[Vec<PivotValue>],
        row_label_width: usize,
        col_key_count: usize,
        value_field_count: usize,
        value_field_idx: usize,
        has_row_grand_totals: bool,
    ) -> Vec<Option<f64>> {
        let mut out = vec![None; data.len()];
        let row_grand_total_start = row_label_width + col_key_count * value_field_count;

        for r in 1..data.len() {
            if has_row_grand_totals {
                out[r] = data[r][row_grand_total_start + value_field_idx].as_number();
                continue;
            }

            let mut sum = 0.0;
            let mut saw_number = false;
            for col_idx in 0..col_key_count {
                let c = row_label_width + col_idx * value_field_count + value_field_idx;
                if let Some(n) = data[r][c].as_number() {
                    sum += n;
                    saw_number = true;
                }
            }
            out[r] = saw_number.then_some(sum);
        }

        out
    }

    fn apply_percent_of_row_total(
        data: &mut [Vec<PivotValue>],
        cols: &[usize],
        row_denoms: &[Option<f64>],
    ) {
        for r in 1..data.len() {
            let denom = row_denoms.get(r).and_then(|d| d.filter(|d| *d != 0.0));
            for &c in cols {
                if let Some(n) = data[r][c].as_number() {
                    if let Some(d) = denom {
                        data[r][c] = PivotValue::Number(n / d);
                    } else {
                        data[r][c] = PivotValue::Blank;
                    }
                }
            }
        }
    }

    fn apply_percent_of_column_total(
        data: &mut [Vec<PivotValue>],
        row_label_width: usize,
        col_key_count: usize,
        value_field_count: usize,
        value_field_idx: usize,
        has_row_grand_totals: bool,
        col_denoms: &[Option<f64>],
        grand_denom: Option<f64>,
    ) {
        let row_grand_total_start = row_label_width + col_key_count * value_field_count;

        for r in 1..data.len() {
            for col_idx in 0..col_key_count {
                let denom = col_denoms
                    .get(col_idx)
                    .and_then(|d| d.filter(|d| *d != 0.0));
                let c = row_label_width + col_idx * value_field_count + value_field_idx;
                if let Some(n) = data[r][c].as_number() {
                    if let Some(d) = denom {
                        data[r][c] = PivotValue::Number(n / d);
                    } else {
                        data[r][c] = PivotValue::Blank;
                    }
                }
            }

            if has_row_grand_totals {
                let denom = grand_denom.filter(|d| *d != 0.0);
                let c = row_grand_total_start + value_field_idx;
                if let Some(n) = data[r][c].as_number() {
                    if let Some(d) = denom {
                        data[r][c] = PivotValue::Number(n / d);
                    } else {
                        data[r][c] = PivotValue::Blank;
                    }
                }
            }
        }
    }

    fn apply_percent_of_base_item_cell(
        data: &mut [Vec<PivotValue>],
        r: usize,
        c: usize,
        denom: Option<f64>,
        difference: bool,
    ) {
        let Some(n) = data
            .get(r)
            .and_then(|row| row.get(c))
            .and_then(|v| v.as_number())
        else {
            return;
        };
        let Some(d) = denom.filter(|d| *d != 0.0) else {
            data[r][c] = PivotValue::Blank;
            return;
        };

        let out = if difference { (n - d) / d } else { n / d };
        data[r][c] = PivotValue::Number(out);
    }

    fn row_key_has_prefix(row_key: &PivotKey, prefix_key: &PivotKey) -> bool {
        let prefix_len = prefix_key.0.len();
        if prefix_len == 0 {
            return true;
        }
        if row_key.0.len() < prefix_len {
            return false;
        }
        row_key
            .0
            .iter()
            .take(prefix_len)
            .zip(prefix_key.0.iter())
            .all(|(a, b)| a == b)
    }

    fn row_key_has_prefix_with_base_row_part(
        row_key: &PivotKey,
        prefix_key: &PivotKey,
        base_row_pos: usize,
        base_part: &PivotKeyPart,
    ) -> bool {
        let prefix_len = prefix_key.0.len();
        if row_key.0.len() < prefix_len {
            return false;
        }

        for idx in 0..prefix_len {
            if idx == base_row_pos {
                if row_key.0[idx] != *base_part {
                    return false;
                }
            } else if row_key.0[idx] != prefix_key.0[idx] {
                return false;
            }
        }

        if base_row_pos >= prefix_len {
            return row_key.0.get(base_row_pos) == Some(base_part);
        }

        true
    }

    fn cube_cell_number(
        cube: &HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_key: &PivotKey,
        col_key: &PivotKey,
        value_field_idx: usize,
        agg: AggregationType,
    ) -> Option<f64> {
        let row_map = cube.get(row_key)?;
        let cell_accs = row_map.get(col_key)?;
        cell_accs
            .get(value_field_idx)
            .map(|acc| acc.finalize(agg).as_number())
            .flatten()
    }

    fn cube_row_total(
        cube: &HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_key: &PivotKey,
        col_keys: &[PivotKey],
        value_field_idx: usize,
        agg: AggregationType,
    ) -> Option<f64> {
        let row_map = cube.get(row_key)?;
        let mut acc = Accumulator::new();
        let mut saw = false;
        for col_key in col_keys {
            if let Some(cell_accs) = row_map.get(col_key) {
                acc.merge(&cell_accs[value_field_idx]);
                saw = true;
            }
        }
        saw.then(|| acc.finalize(agg).as_number()).flatten()
    }

    fn cube_row_total_filtered_by_col_part(
        cube: &HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_key: &PivotKey,
        col_keys: &[PivotKey],
        value_field_idx: usize,
        agg: AggregationType,
        base_col_pos: usize,
        base_part: &PivotKeyPart,
    ) -> Option<f64> {
        let row_map = cube.get(row_key)?;
        let mut acc = Accumulator::new();
        let mut saw = false;
        for col_key in col_keys {
            if col_key.0.get(base_col_pos) != Some(base_part) {
                continue;
            }
            if let Some(cell_accs) = row_map.get(col_key) {
                acc.merge(&cell_accs[value_field_idx]);
                saw = true;
            }
        }
        saw.then(|| acc.finalize(agg).as_number()).flatten()
    }

    fn cube_col_total(
        cube: &HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_keys: &[PivotKey],
        col_key: &PivotKey,
        value_field_idx: usize,
        agg: AggregationType,
    ) -> Option<f64> {
        let mut acc = Accumulator::new();
        let mut saw = false;
        for row_key in row_keys {
            let Some(row_map) = cube.get(row_key) else {
                continue;
            };
            if let Some(cell_accs) = row_map.get(col_key) {
                acc.merge(&cell_accs[value_field_idx]);
                saw = true;
            }
        }
        saw.then(|| acc.finalize(agg).as_number()).flatten()
    }

    fn cube_col_total_filtered_by_row_part(
        cube: &HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_keys: &[PivotKey],
        col_key: &PivotKey,
        value_field_idx: usize,
        agg: AggregationType,
        base_row_pos: usize,
        base_part: &PivotKeyPart,
    ) -> Option<f64> {
        let mut acc = Accumulator::new();
        let mut saw = false;
        for row_key in row_keys {
            if row_key.0.get(base_row_pos) != Some(base_part) {
                continue;
            }
            let Some(row_map) = cube.get(row_key) else {
                continue;
            };
            if let Some(cell_accs) = row_map.get(col_key) {
                acc.merge(&cell_accs[value_field_idx]);
                saw = true;
            }
        }
        saw.then(|| acc.finalize(agg).as_number()).flatten()
    }

    fn cube_grand_total_filtered_by_row_part(
        cube: &HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_keys: &[PivotKey],
        col_keys: &[PivotKey],
        value_field_idx: usize,
        agg: AggregationType,
        base_row_pos: usize,
        base_part: &PivotKeyPart,
    ) -> Option<f64> {
        let mut acc = Accumulator::new();
        let mut saw = false;
        for row_key in row_keys {
            if row_key.0.get(base_row_pos) != Some(base_part) {
                continue;
            }
            let Some(row_map) = cube.get(row_key) else {
                continue;
            };
            for col_key in col_keys {
                if let Some(cell_accs) = row_map.get(col_key) {
                    acc.merge(&cell_accs[value_field_idx]);
                    saw = true;
                }
            }
        }
        saw.then(|| acc.finalize(agg).as_number()).flatten()
    }

    fn cube_grand_total_filtered_by_col_part(
        cube: &HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_keys: &[PivotKey],
        col_keys: &[PivotKey],
        value_field_idx: usize,
        agg: AggregationType,
        base_col_pos: usize,
        base_part: &PivotKeyPart,
    ) -> Option<f64> {
        let mut acc = Accumulator::new();
        let mut saw = false;
        for row_key in row_keys {
            let Some(row_map) = cube.get(row_key) else {
                continue;
            };
            for col_key in col_keys {
                if col_key.0.get(base_col_pos) != Some(base_part) {
                    continue;
                }
                if let Some(cell_accs) = row_map.get(col_key) {
                    acc.merge(&cell_accs[value_field_idx]);
                    saw = true;
                }
            }
        }
        saw.then(|| acc.finalize(agg).as_number()).flatten()
    }

    fn apply_percent_of_base_item_row_field(
        data: &mut [Vec<PivotValue>],
        leaf_rows: &[(usize, usize)],
        subtotal_rows: &[(usize, &PivotKey)],
        grand_total_row: Option<usize>,
        cube: &HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_keys: &[PivotKey],
        col_keys: &[PivotKey],
        regular_cols: &[usize],
        row_total_col: Option<usize>,
        value_field_idx: usize,
        agg: AggregationType,
        base_row_pos: usize,
        base_part: &PivotKeyPart,
        difference: bool,
    ) {
        for &(r, row_key_idx) in leaf_rows {
            let Some(row_key) = row_keys.get(row_key_idx) else {
                continue;
            };
            if base_row_pos >= row_key.0.len() {
                continue;
            }
            let mut base_key_parts = row_key.0.clone();
            base_key_parts[base_row_pos] = base_part.clone();
            let base_row_key = PivotKey(base_key_parts);

            for (col_idx, col_key) in col_keys.iter().enumerate() {
                let denom =
                    Self::cube_cell_number(cube, &base_row_key, col_key, value_field_idx, agg);
                Self::apply_percent_of_base_item_cell(
                    data,
                    r,
                    regular_cols[col_idx],
                    denom,
                    difference,
                );
            }

            if let Some(total_col) = row_total_col {
                let denom =
                    Self::cube_row_total(cube, &base_row_key, col_keys, value_field_idx, agg);
                Self::apply_percent_of_base_item_cell(data, r, total_col, denom, difference);
            }
        }

        for &(r, prefix_key) in subtotal_rows {
            for (col_idx, col_key) in col_keys.iter().enumerate() {
                let mut acc = Accumulator::new();
                let mut saw = false;
                for row_key in row_keys {
                    if !Self::row_key_has_prefix_with_base_row_part(
                        row_key,
                        prefix_key,
                        base_row_pos,
                        base_part,
                    ) {
                        continue;
                    }
                    let Some(row_map) = cube.get(row_key) else {
                        continue;
                    };
                    if let Some(cell_accs) = row_map.get(col_key) {
                        acc.merge(&cell_accs[value_field_idx]);
                        saw = true;
                    }
                }
                let denom = saw.then(|| acc.finalize(agg).as_number()).flatten();
                Self::apply_percent_of_base_item_cell(
                    data,
                    r,
                    regular_cols[col_idx],
                    denom,
                    difference,
                );
            }

            if let Some(total_col) = row_total_col {
                let mut acc = Accumulator::new();
                let mut saw = false;
                for row_key in row_keys {
                    if !Self::row_key_has_prefix_with_base_row_part(
                        row_key,
                        prefix_key,
                        base_row_pos,
                        base_part,
                    ) {
                        continue;
                    }
                    let Some(row_map) = cube.get(row_key) else {
                        continue;
                    };
                    for col_key in col_keys {
                        if let Some(cell_accs) = row_map.get(col_key) {
                            acc.merge(&cell_accs[value_field_idx]);
                            saw = true;
                        }
                    }
                }
                let denom = saw.then(|| acc.finalize(agg).as_number()).flatten();
                Self::apply_percent_of_base_item_cell(data, r, total_col, denom, difference);
            }
        }

        if let Some(grand_r) = grand_total_row {
            for (col_idx, col_key) in col_keys.iter().enumerate() {
                let denom = Self::cube_col_total_filtered_by_row_part(
                    cube,
                    row_keys,
                    col_key,
                    value_field_idx,
                    agg,
                    base_row_pos,
                    base_part,
                );
                Self::apply_percent_of_base_item_cell(
                    data,
                    grand_r,
                    regular_cols[col_idx],
                    denom,
                    difference,
                );
            }

            if let Some(total_col) = row_total_col {
                let denom = Self::cube_grand_total_filtered_by_row_part(
                    cube,
                    row_keys,
                    col_keys,
                    value_field_idx,
                    agg,
                    base_row_pos,
                    base_part,
                );
                Self::apply_percent_of_base_item_cell(data, grand_r, total_col, denom, difference);
            }
        }
    }

    fn apply_percent_of_base_item_column_field(
        data: &mut [Vec<PivotValue>],
        leaf_rows: &[(usize, usize)],
        subtotal_rows: &[(usize, &PivotKey)],
        grand_total_row: Option<usize>,
        cube: &HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>>,
        row_keys: &[PivotKey],
        col_keys: &[PivotKey],
        regular_cols: &[usize],
        row_total_col: Option<usize>,
        value_field_idx: usize,
        agg: AggregationType,
        base_col_pos: usize,
        base_part: &PivotKeyPart,
        difference: bool,
    ) {
        let base_col_keys = col_keys
            .iter()
            .map(|col_key| {
                let mut parts = col_key.0.clone();
                if base_col_pos < parts.len() {
                    parts[base_col_pos] = base_part.clone();
                }
                PivotKey(parts)
            })
            .collect::<Vec<_>>();

        for &(r, row_key_idx) in leaf_rows {
            let Some(row_key) = row_keys.get(row_key_idx) else {
                continue;
            };

            for col_idx in 0..col_keys.len() {
                let denom = Self::cube_cell_number(
                    cube,
                    row_key,
                    &base_col_keys[col_idx],
                    value_field_idx,
                    agg,
                );
                Self::apply_percent_of_base_item_cell(
                    data,
                    r,
                    regular_cols[col_idx],
                    denom,
                    difference,
                );
            }

            if let Some(total_col) = row_total_col {
                let denom = Self::cube_row_total_filtered_by_col_part(
                    cube,
                    row_key,
                    col_keys,
                    value_field_idx,
                    agg,
                    base_col_pos,
                    base_part,
                );
                Self::apply_percent_of_base_item_cell(data, r, total_col, denom, difference);
            }
        }

        for &(r, prefix_key) in subtotal_rows {
            for col_idx in 0..col_keys.len() {
                let mut acc = Accumulator::new();
                let mut saw = false;
                for row_key in row_keys {
                    if !Self::row_key_has_prefix(row_key, prefix_key) {
                        continue;
                    }
                    let Some(row_map) = cube.get(row_key) else {
                        continue;
                    };
                    if let Some(cell_accs) = row_map.get(&base_col_keys[col_idx]) {
                        acc.merge(&cell_accs[value_field_idx]);
                        saw = true;
                    }
                }
                let denom = saw.then(|| acc.finalize(agg).as_number()).flatten();
                Self::apply_percent_of_base_item_cell(
                    data,
                    r,
                    regular_cols[col_idx],
                    denom,
                    difference,
                );
            }

            if let Some(total_col) = row_total_col {
                let mut acc = Accumulator::new();
                let mut saw = false;
                for row_key in row_keys {
                    if !Self::row_key_has_prefix(row_key, prefix_key) {
                        continue;
                    }
                    let Some(row_map) = cube.get(row_key) else {
                        continue;
                    };
                    for col_key in col_keys {
                        if col_key.0.get(base_col_pos) != Some(base_part) {
                            continue;
                        }
                        if let Some(cell_accs) = row_map.get(col_key) {
                            acc.merge(&cell_accs[value_field_idx]);
                            saw = true;
                        }
                    }
                }
                let denom = saw.then(|| acc.finalize(agg).as_number()).flatten();
                Self::apply_percent_of_base_item_cell(data, r, total_col, denom, difference);
            }
        }

        if let Some(grand_r) = grand_total_row {
            for col_idx in 0..col_keys.len() {
                let denom = Self::cube_col_total(
                    cube,
                    row_keys,
                    &base_col_keys[col_idx],
                    value_field_idx,
                    agg,
                );
                Self::apply_percent_of_base_item_cell(
                    data,
                    grand_r,
                    regular_cols[col_idx],
                    denom,
                    difference,
                );
            }

            if let Some(total_col) = row_total_col {
                let denom = Self::cube_grand_total_filtered_by_col_part(
                    cube,
                    row_keys,
                    col_keys,
                    value_field_idx,
                    agg,
                    base_col_pos,
                    base_part,
                );
                Self::apply_percent_of_base_item_cell(data, grand_r, total_col, denom, difference);
            }
        }
    }

    fn group_ids_excluding_pos(keys: &[PivotKey], exclude_pos: usize) -> (Vec<usize>, usize) {
        let mut id_by_key: HashMap<PivotKey, usize> = HashMap::new();
        let mut out = Vec::with_capacity(keys.len());

        for key in keys {
            let group_parts = key
                .0
                .iter()
                .enumerate()
                .filter_map(|(idx, part)| (idx != exclude_pos).then_some(part.clone()))
                .collect::<Vec<_>>();
            let group_key = PivotKey(group_parts);

            let next_id = id_by_key.len();
            let id = *id_by_key.entry(group_key).or_insert(next_id);
            out.push(id);
        }

        (out, id_by_key.len())
    }

    fn apply_running_total_grouped_by_row(
        data: &mut [Vec<PivotValue>],
        leaf_rows: &[(usize, usize)],
        cols: &[usize],
        row_group_ids: &[usize],
        group_count: usize,
    ) {
        if group_count == 0 || leaf_rows.is_empty() {
            return;
        }

        for &c in cols {
            let mut running_by_group = vec![0.0; group_count];
            for &(r, row_key_idx) in leaf_rows {
                let Some(group_id) = row_group_ids.get(row_key_idx).copied() else {
                    continue;
                };
                let Some(n) = data
                    .get(r)
                    .and_then(|row| row.get(c))
                    .and_then(|v| v.as_number())
                else {
                    continue;
                };
                if let Some(running) = running_by_group.get_mut(group_id) {
                    *running += n;
                    data[r][c] = PivotValue::Number(*running);
                }
            }
        }
    }

    fn apply_running_total_grouped_by_column(
        data: &mut [Vec<PivotValue>],
        leaf_rows: &[usize],
        cols: &[usize],
        col_group_ids: &[usize],
        group_count: usize,
    ) {
        if group_count == 0 || leaf_rows.is_empty() {
            return;
        }

        for &r in leaf_rows {
            let mut running_by_group = vec![0.0; group_count];
            for (col_idx, &c) in cols.iter().enumerate() {
                let Some(group_id) = col_group_ids.get(col_idx).copied() else {
                    continue;
                };
                let Some(n) = data
                    .get(r)
                    .and_then(|row| row.get(c))
                    .and_then(|v| v.as_number())
                else {
                    continue;
                };
                if let Some(running) = running_by_group.get_mut(group_id) {
                    *running += n;
                    data[r][c] = PivotValue::Number(*running);
                }
            }
        }
    }

    fn apply_running_total(data: &mut [Vec<PivotValue>], leaf_rows: &[usize], cols: &[usize]) {
        for &c in cols {
            let mut running = 0.0;
            for &r in leaf_rows {
                if let Some(n) = data[r][c].as_number() {
                    running += n;
                    data[r][c] = PivotValue::Number(running);
                }
            }
        }
    }

    fn apply_rank_grouped_by_row(
        data: &mut [Vec<PivotValue>],
        leaf_rows: &[(usize, usize)],
        cols: &[usize],
        row_group_ids: &[usize],
        group_count: usize,
        descending: bool,
    ) {
        if group_count == 0 || leaf_rows.is_empty() {
            return;
        }

        for &c in cols {
            let mut values_by_group: Vec<Vec<(usize, f64)>> = vec![Vec::new(); group_count];
            for &(r, row_key_idx) in leaf_rows {
                let Some(group_id) = row_group_ids.get(row_key_idx).copied() else {
                    continue;
                };
                let Some(n) = data
                    .get(r)
                    .and_then(|row| row.get(c))
                    .and_then(|v| v.as_number())
                else {
                    continue;
                };
                if let Some(group) = values_by_group.get_mut(group_id) {
                    group.push((r, n));
                }
            }

            for mut values in values_by_group {
                if values.is_empty() {
                    continue;
                }

                values.sort_by(|(_, a), (_, b)| {
                    if descending {
                        b.total_cmp(a)
                    } else {
                        a.total_cmp(b)
                    }
                });

                let mut next_rank = 1usize;
                let mut i = 0usize;
                while i < values.len() {
                    let value = values[i].1;
                    let mut j = i + 1;
                    while j < values.len() && values[j].1 == value {
                        j += 1;
                    }
                    for k in i..j {
                        data[values[k].0][c] = PivotValue::Number(next_rank as f64);
                    }
                    next_rank += j - i;
                    i = j;
                }
            }
        }
    }

    fn apply_rank_grouped_by_column(
        data: &mut [Vec<PivotValue>],
        leaf_rows: &[usize],
        cols: &[usize],
        col_group_ids: &[usize],
        group_count: usize,
        descending: bool,
    ) {
        if group_count == 0 || leaf_rows.is_empty() {
            return;
        }

        for &r in leaf_rows {
            let mut values_by_group: Vec<Vec<(usize, f64)>> = vec![Vec::new(); group_count];
            for (col_idx, &c) in cols.iter().enumerate() {
                let Some(group_id) = col_group_ids.get(col_idx).copied() else {
                    continue;
                };
                let Some(n) = data
                    .get(r)
                    .and_then(|row| row.get(c))
                    .and_then(|v| v.as_number())
                else {
                    continue;
                };
                if let Some(group) = values_by_group.get_mut(group_id) {
                    group.push((col_idx, n));
                }
            }

            for mut values in values_by_group {
                if values.is_empty() {
                    continue;
                }

                values.sort_by(|(_, a), (_, b)| {
                    if descending {
                        b.total_cmp(a)
                    } else {
                        a.total_cmp(b)
                    }
                });

                let mut next_rank = 1usize;
                let mut i = 0usize;
                while i < values.len() {
                    let value = values[i].1;
                    let mut j = i + 1;
                    while j < values.len() && values[j].1 == value {
                        j += 1;
                    }
                    for k in i..j {
                        let col_idx = values[k].0;
                        if let Some(&c) = cols.get(col_idx) {
                            data[r][c] = PivotValue::Number(next_rank as f64);
                        }
                    }
                    next_rank += j - i;
                    i = j;
                }
            }
        }
    }

    fn apply_rank(
        data: &mut [Vec<PivotValue>],
        leaf_rows: &[usize],
        cols: &[usize],
        descending: bool,
    ) {
        for &c in cols {
            let mut values: Vec<(usize, f64)> = leaf_rows
                .iter()
                .filter_map(|&r| data[r][c].as_number().map(|n| (r, n)))
                .collect();
            if values.is_empty() {
                continue;
            }

            values.sort_by(|(_, a), (_, b)| {
                if descending {
                    b.total_cmp(a)
                } else {
                    a.total_cmp(b)
                }
            });

            let mut rank_by_row: HashMap<usize, usize> = HashMap::new();
            let mut next_rank = 1usize;
            let mut i = 0usize;
            while i < values.len() {
                let value = values[i].1;
                let mut j = i + 1;
                while j < values.len() && values[j].1 == value {
                    j += 1;
                }
                for k in i..j {
                    rank_by_row.insert(values[k].0, next_rank);
                }
                next_rank += j - i;
                i = j;
            }

            for &r in leaf_rows {
                if let Some(rank) = rank_by_row.get(&r) {
                    data[r][c] = PivotValue::Number(*rank as f64);
                }
            }
        }
    }
}

#[derive(Debug, Clone)]
struct GroupAccumulator {
    cells: HashMap<PivotKey, Vec<Accumulator>>,
}

impl GroupAccumulator {
    fn new() -> Self {
        Self {
            cells: HashMap::new(),
        }
    }

    fn merge_row(
        &mut self,
        row_map: Option<&HashMap<PivotKey, Vec<Accumulator>>>,
        col_keys: &[PivotKey],
        value_field_count: usize,
    ) {
        for col_key in col_keys {
            let src = row_map.and_then(|m| m.get(col_key));
            if let Some(src_accs) = src {
                let dst = self.cells.entry(col_key.clone()).or_insert_with(|| {
                    (0..value_field_count).map(|_| Accumulator::new()).collect()
                });
                for i in 0..value_field_count {
                    dst[i].merge(&src_accs[i]);
                }
            }
        }
    }
}

struct FieldIndices {
    row_indices: Vec<usize>,
    col_indices: Vec<usize>,
    value_indices: Vec<usize>,
    filter_indices: Vec<(usize, Option<HashSet<PivotKeyPart>>)>,
}

impl FieldIndices {
    fn new<S: PivotRecordSource + ?Sized>(
        source: &S,
        cfg: &PivotConfig,
    ) -> Result<Self, PivotError> {
        let resolve_field_index = |field: &PivotFieldRef| -> Result<usize, PivotError> {
            // Most pivots are cache-backed; try the cache-field name / legacy string form first.
            let field_name = pivot_field_ref_name(field);
            if let Some(idx) = source.field_index(field_name.as_ref()) {
                return Ok(idx);
            }

            // Best-effort: match Data Model refs against cache field captions. Caches may store
            // measures either as the raw name (`Total`) or in DAX bracket form (`[Total]`), and may
            // store column refs with or without quoted table names.
            match field {
                PivotFieldRef::DataModelMeasure(name) => {
                    if let Some(idx) = source.field_index(name) {
                        return Ok(idx);
                    }
                }
                PivotFieldRef::DataModelColumn { table, column } => {
                    let column_escaped = escape_dax_bracket_identifier(column);
                    let unquoted = format!("{table}[{column_escaped}]");
                    if let Some(idx) = source.field_index(&unquoted) {
                        return Ok(idx);
                    }
                    let quoted = dax_quoted_column_ref(table, column);
                    if let Some(idx) = source.field_index(&quoted) {
                        return Ok(idx);
                    }

                    // Rare: quoted table name but raw (unescaped) column name.
                    let quoted_table = dax_quoted_table_name(table);
                    let quoted_unescaped = format!("'{quoted_table}'[{column}]");
                    if quoted_unescaped != quoted {
                        if let Some(idx) = source.field_index(&quoted_unescaped) {
                            return Ok(idx);
                        }
                    }
                }
                PivotFieldRef::CacheFieldName(_) => {}
            }

            // Best-effort fallback: try the `Display` rendering (used by some Data Model producers).
            let label = field.to_string();
            source
                .field_index(&label)
                .ok_or_else(|| PivotError::MissingField(field_name.to_string()))
        };

        let mut row_indices = Vec::new();
        for f in &cfg.row_fields {
            row_indices.push(resolve_field_index(&f.source_field)?);
        }

        let mut col_indices = Vec::new();
        for f in &cfg.column_fields {
            col_indices.push(resolve_field_index(&f.source_field)?);
        }

        let mut value_indices = Vec::new();
        for f in &cfg.value_fields {
            value_indices.push(resolve_field_index(&f.source_field)?);
        }

        let mut filter_indices = Vec::new();
        for f in &cfg.filter_fields {
            let idx = resolve_field_index(&f.source_field)?;
            filter_indices.push((idx, f.allowed.clone()));
        }

        Ok(Self {
            row_indices,
            col_indices,
            value_indices,
            filter_indices,
        })
    }

    fn build_key<S: PivotRecordSource + ?Sized>(
        &self,
        source: &S,
        row: usize,
        indices: &[usize],
    ) -> PivotKey {
        PivotKey(
            indices
                .iter()
                .map(|idx| source.value(row, *idx).to_key_part())
                .collect(),
        )
    }

    fn passes_filters<S: PivotRecordSource + ?Sized>(&self, source: &S, row: usize) -> bool {
        for (idx, allowed) in &self.filter_indices {
            if let Some(set) = allowed {
                let val = source.value(row, *idx).to_key_part();
                if !set.contains(&val) {
                    return false;
                }
            }
        }
        true
    }
}

#[derive(Debug, Clone)]
struct KeySortSpec {
    sort_order: SortOrder,
    manual_index: Option<HashMap<PivotKeyPart, usize>>,
}

impl KeySortSpec {
    fn for_field(field: &PivotField) -> Self {
        let manual_index = if field.sort_order == SortOrder::Manual {
            field.manual_sort.as_ref().map(|items| {
                let mut index = HashMap::with_capacity(items.len());
                for (pos, part) in items.iter().enumerate() {
                    index.entry(part.clone()).or_insert(pos);
                }
                index
            })
        } else {
            None
        };

        Self {
            sort_order: field.sort_order,
            manual_index,
        }
    }
}

fn compare_key_parts_ascending(left: &PivotKeyPart, right: &PivotKeyPart) -> Ordering {
    left.cmp(right)
}

fn compare_key_parts_for_field(
    left: &PivotKeyPart,
    right: &PivotKeyPart,
    spec: &KeySortSpec,
) -> Ordering {
    // Blank values always sort last, regardless of ascending/descending/manual.
    match (left, right) {
        (PivotKeyPart::Blank, PivotKeyPart::Blank) => return Ordering::Equal,
        (PivotKeyPart::Blank, _) => return Ordering::Greater,
        (_, PivotKeyPart::Blank) => return Ordering::Less,
        _ => {}
    }

    match spec.sort_order {
        SortOrder::Ascending => compare_key_parts_ascending(left, right),
        SortOrder::Descending => compare_key_parts_ascending(left, right).reverse(),
        SortOrder::Manual => {
            if let Some(index) = &spec.manual_index {
                match (index.get(left), index.get(right)) {
                    (Some(a_pos), Some(b_pos)) => a_pos.cmp(b_pos),
                    (Some(_), None) => Ordering::Less,
                    (None, Some(_)) => Ordering::Greater,
                    (None, None) => compare_key_parts_ascending(left, right),
                }
            } else {
                compare_key_parts_ascending(left, right)
            }
        }
    }
}

fn compare_pivot_keys(left: &PivotKey, right: &PivotKey, specs: &[KeySortSpec]) -> Ordering {
    let blank = PivotKeyPart::Blank;
    for (idx, spec) in specs.iter().enumerate() {
        let left_part = left.0.get(idx).unwrap_or(&blank);
        let right_part = right.0.get(idx).unwrap_or(&blank);
        let ord = compare_key_parts_for_field(left_part, right_part, spec);
        if ord != Ordering::Equal {
            return ord;
        }
    }
    // If all configured fields compare equal, fall back to the full typed key ordering to keep
    // output deterministic even when display strings collide (HashSet iteration order is not
    // stable across runs).
    left.cmp(right)
}

fn common_prefix_len(a: &[PivotKeyPart], b: &[PivotKeyPart]) -> usize {
    let mut i = 0;
    while i < a.len() && i < b.len() && a[i] == b[i] {
        i += 1;
    }
    i
}

#[cfg(test)]
mod tests {
    use super::*;

    use formula_columnar::{
        ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value,
    };
    use pretty_assertions::assert_eq;
    use std::sync::Arc;

    fn cache_field(name: &str) -> PivotFieldRef {
        PivotFieldRef::CacheFieldName(name.to_string())
    }

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
    }

    #[test]
    fn pivot_cache_field_index_ref_resolves_quoted_data_model_column_headers() {
        let data = vec![
            pv_row(&["'Sales Table'[Region]".into(), "Sales".into()]),
            pv_row(&["East".into(), 100.into()]),
        ];
        let cache = PivotCache::from_range(&data).unwrap();

        let field = PivotFieldRef::DataModelColumn {
            table: "Sales Table".to_string(),
            column: "Region".to_string(),
        };
        assert_eq!(cache.field_index_ref(&field), Some(0));
    }

    #[test]
    fn pivot_value_display_string_uses_general_number_formatting_and_does_not_saturate_large_ints()
    {
        let s = PivotValue::Number(1e20).display_string();
        assert_ne!(s, i64::MAX.to_string());
        assert!(s.contains('E'), "{s}");
        assert!(s.starts_with('1'), "{s}");
    }

    #[test]
    fn pivot_value_display_string_normalizes_negative_zero() {
        assert_eq!(PivotValue::Number(-0.0).display_string(), "0");
    }

    #[test]
    fn pivot_value_display_string_formats_booleans_like_excel() {
        assert_eq!(PivotValue::Bool(true).display_string(), "TRUE");
        assert_eq!(PivotValue::Bool(false).display_string(), "FALSE");
    }

    #[test]
    fn pivot_value_display_string_formats_non_finite_numbers_as_num_error() {
        // Excel doesn't have NaN/Infinity; Formula renders them as #NUM! (matching `formula-format`).
        assert_eq!(PivotValue::Number(f64::NAN).display_string(), "#NUM!");
        assert_eq!(PivotValue::Number(f64::INFINITY).display_string(), "#NUM!");
        assert_eq!(PivotValue::Number(f64::NEG_INFINITY).display_string(), "#NUM!");
    }

    #[test]
    fn pivot_key_part_display_string_uses_general_number_formatting_and_does_not_saturate_large_ints()
    {
        let s = PivotKeyPart::Number((1e20_f64).to_bits()).display_string();
        assert_ne!(s, i64::MAX.to_string());
        assert!(s.contains('E'), "{s}");
        assert!(s.starts_with('1'), "{s}");
    }

    #[test]
    fn pivot_key_part_display_string_normalizes_negative_zero() {
        assert_eq!(
            PivotKeyPart::Number((-0.0_f64).to_bits()).display_string(),
            "0"
        );
    }

    #[test]
    fn pivot_key_part_display_string_formats_booleans_like_excel() {
        assert_eq!(PivotKeyPart::Bool(true).display_string(), "TRUE");
        assert_eq!(PivotKeyPart::Bool(false).display_string(), "FALSE");
    }

    #[test]
    fn pivot_key_part_display_string_formats_non_finite_numbers_as_num_error() {
        assert_eq!(
            PivotKeyPart::Number(PivotValue::canonical_number_bits(f64::NAN)).display_string(),
            "#NUM!"
        );
        assert_eq!(
            PivotKeyPart::Number(f64::INFINITY.to_bits()).display_string(),
            "#NUM!"
        );
        assert_eq!(
            PivotKeyPart::Number(f64::NEG_INFINITY.to_bits()).display_string(),
            "#NUM!"
        );
    }

    #[test]
    fn pivot_config_serde_roundtrips_with_calculated_fields_and_items() {
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![],
            value_fields: vec![],
            filter_fields: vec![],
            calculated_fields: vec![CalculatedField {
                name: "Profit".to_string(),
                formula: "Sales - Cost".to_string(),
            }],
            calculated_items: vec![CalculatedItem {
                field: "Region".to_string(),
                name: "East+West".to_string(),
                formula: "\"East\" + \"West\"".to_string(),
            }],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals::default(),
        };

        let json = serde_json::to_value(&cfg).unwrap();
        assert!(json.get("calculatedFields").is_some());
        assert!(json.get("calculatedItems").is_some());

        let decoded: PivotConfig = serde_json::from_value(json.clone()).unwrap();
        assert_eq!(decoded, cfg);

        // Backward-compat: missing keys should default to empty vectors.
        let mut json_without = json;
        if let Some(obj) = json_without.as_object_mut() {
            obj.remove("calculatedFields");
            obj.remove("calculatedItems");
        }
        let decoded: PivotConfig = serde_json::from_value(json_without).unwrap();
        assert!(decoded.calculated_fields.is_empty());
        assert!(decoded.calculated_items.is_empty());
    }

    #[test]
    fn create_pivot_table_request_parses_legacy_pivot_field_refs() {
        let req = CreatePivotTableRequest {
            name: None,
            row_fields: vec!["Sales[Amount]".to_string()],
            column_fields: vec![],
            value_fields: vec![CreatePivotValueSpec {
                field: "[Total Sales]".to_string(),
                aggregation: AggregationType::Sum,
                name: None,
            }],
            filter_fields: vec![CreatePivotFilterSpec {
                field: "Region".to_string(),
                allowed: None,
            }],
            calculated_fields: None,
            calculated_items: None,
            layout: None,
            subtotals: None,
            grand_totals: None,
        };

        let cfg = req.into_config();

        assert_eq!(
            cfg.row_fields[0].source_field,
            PivotFieldRef::DataModelColumn {
                table: "Sales".to_string(),
                column: "Amount".to_string(),
            }
        );
        assert_eq!(
            cfg.value_fields[0].source_field,
            PivotFieldRef::DataModelMeasure("Total Sales".to_string())
        );
        assert_eq!(cfg.value_fields[0].name, "Sum of Total Sales");
        assert_eq!(cfg.filter_fields[0].source_field, cache_field("Region"));
    }

    #[test]
    fn pivot_header_renders_data_model_columns_using_display_string() {
        let data = vec![
            pv_row(&["'Sales Table'[Region]".into(), "Sales".into()]),
            pv_row(&["East".into(), 100.into()]),
            pv_row(&["West".into(), 200.into()]),
        ];
        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField {
                source_field: PivotFieldRef::DataModelColumn {
                    table: "Sales Table".to_string(),
                    column: "Region".to_string(),
                },
                sort_order: SortOrder::default(),
                manual_sort: None,
            }],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data[0][0],
            PivotValue::Text("Sales Table[Region]".to_string())
        );
    }

    #[test]
    fn pivot_header_resolves_quoted_table_with_unescaped_brackets_in_column_caption() {
        let data = vec![
            pv_row(&["'Sales Table'[A]B]".into(), "Sales".into()]),
            pv_row(&["East".into(), 100.into()]),
            pv_row(&["West".into(), 200.into()]),
        ];
        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField {
                source_field: PivotFieldRef::DataModelColumn {
                    table: "Sales Table".to_string(),
                    column: "A]B".to_string(),
                },
                sort_order: SortOrder::default(),
                manual_sort: None,
            }],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data[0][0],
            PivotValue::Text("Sales Table[A]]B]".to_string())
        );
    }

    #[test]
    fn calculates_sum_by_single_row_field_with_grand_total() {
        let data = vec![
            pv_row(&["Region".into(), "Product".into(), "Sales".into()]),
            pv_row(&["East".into(), "A".into(), 100.into()]),
            pv_row(&["East".into(), "B".into(), 150.into()]),
            pv_row(&["West".into(), "A".into(), 200.into()]),
            pv_row(&["West".into(), "B".into(), 250.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            column_fields: vec![],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec!["Region".into(), "Sum of Sales".into()],
                vec!["East".into(), 250.into()],
                vec!["West".into(), 450.into()],
                vec!["Grand Total".into(), 700.into()],
            ]
        );
    }

    #[test]
    fn calculated_item_on_row_field_creates_synthetic_row_and_updates_grand_total() {
        let data = vec![
            pv_row(&["Region".into(), "Product".into(), "Sales".into()]),
            pv_row(&["East".into(), "A".into(), 100.into()]),
            pv_row(&["East".into(), "B".into(), 150.into()]),
            pv_row(&["West".into(), "A".into(), 200.into()]),
            pv_row(&["West".into(), "B".into(), 250.into()]),
        ];
        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![CalculatedItem {
                field: "Region".to_string(),
                name: "East+West".to_string(),
                formula: "\"East\" + \"West\"".to_string(),
            }],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        // Text sorting puts "East+West" after "East" and before "West".
        assert_eq!(
            result.data,
            vec![
                vec!["Region".into(), "Sum of Sales".into()],
                vec!["East".into(), 250.into()],
                vec!["East+West".into(), 700.into()],
                vec!["West".into(), 450.into()],
                vec!["Grand Total".into(), 1400.into()],
            ]
        );
    }

    #[test]
    fn calculated_item_resolves_unicode_item_refs_case_insensitively() {
        let data = vec![
            pv_row(&["Region".into(), "Sales".into()]),
            pv_row(&["Straße".into(), 100.into()]),
            pv_row(&["West".into(), 200.into()]),
        ];
        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![CalculatedItem {
                field: "Region".to_string(),
                name: "Straße+West".to_string(),
                // `ß` uppercases to `SS`, so the existing item `"Straße"` should be addressable as
                // `"STRASSE"` in calculated item formulas.
                formula: "\"STRASSE\" + \"WEST\"".to_string(),
            }],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec!["Region".into(), "Sum of Sales".into()],
                vec!["Straße".into(), 100.into()],
                vec!["Straße+West".into(), 300.into()],
                vec!["West".into(), 200.into()],
                vec!["Grand Total".into(), 600.into()],
            ]
        );
    }

    #[test]
    fn calculated_item_field_must_be_in_layout() {
        let data = vec![
            pv_row(&["Region".into(), "Sales".into()]),
            pv_row(&["East".into(), 100.into()]),
        ];
        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![CalculatedItem {
                field: "NotInLayout".to_string(),
                name: "Any".to_string(),
                formula: "\"East\"".to_string(),
            }],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let err = PivotEngine::calculate(&cache, &cfg).unwrap_err();
        match err {
            PivotError::CalculatedItemFieldNotInLayout(field) => {
                assert_eq!(field, "NotInLayout");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn calculated_item_name_must_not_collide_with_existing_item() {
        let data = vec![
            pv_row(&["Region".into(), "Sales".into()]),
            pv_row(&["East".into(), 100.into()]),
            pv_row(&["West".into(), 200.into()]),
        ];
        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![CalculatedItem {
                field: "Region".to_string(),
                name: "East".to_string(),
                formula: "\"West\"".to_string(),
            }],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let err = PivotEngine::calculate(&cache, &cfg).unwrap_err();
        match err {
            PivotError::CalculatedItemNameConflictsWithExistingItem { field, item } => {
                assert_eq!(field, "Region");
                assert_eq!(item, "East");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn respects_filters() {
        let data = vec![
            pv_row(&["Region".into(), "Product".into(), "Sales".into()]),
            pv_row(&["East".into(), "A".into(), 100.into()]),
            pv_row(&["East".into(), "B".into(), 150.into()]),
            pv_row(&["West".into(), "A".into(), 200.into()]),
            pv_row(&["West".into(), "B".into(), 250.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let mut allowed = HashSet::new();
        allowed.insert(PivotKeyPart::Text("East".to_string()));

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![FilterField {
                source_field: cache_field("Region"),
                allowed: Some(allowed),
            }],
            column_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec!["Region".into(), "Sum of Sales".into()],
                vec!["East".into(), 250.into()],
                vec!["Grand Total".into(), 250.into()],
            ]
        );
    }

    #[test]
    fn supports_column_fields_and_column_grand_totals() {
        let data = vec![
            pv_row(&["Region".into(), "Product".into(), "Sales".into()]),
            pv_row(&["East".into(), "A".into(), 100.into()]),
            pv_row(&["East".into(), "B".into(), 150.into()]),
            pv_row(&["West".into(), "A".into(), 200.into()]),
            pv_row(&["West".into(), "B".into(), 250.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![PivotField::new("Product")],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: true,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        // Column ordering is alphabetical A, B.
        assert_eq!(
            result.data,
            vec![
                vec![
                    "Region".into(),
                    "A - Sum of Sales".into(),
                    "B - Sum of Sales".into(),
                    "Grand Total - Sum of Sales".into()
                ],
                vec!["East".into(), 100.into(), 150.into(), 250.into()],
                vec!["West".into(), 200.into(), 250.into(), 450.into()],
                vec!["Grand Total".into(), 300.into(), 400.into(), 700.into()],
            ]
        );
    }

    #[test]
    fn sorts_row_keys_descending_for_numeric_field() {
        let data = vec![
            pv_row(&["Num".into(), "Value".into()]),
            pv_row(&[1.into(), 10.into()]),
            pv_row(&[2.into(), 20.into()]),
            pv_row(&[10.into(), 30.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField {
                sort_order: SortOrder::Descending,
                ..PivotField::new("Num")
            }],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Value"),
                name: "Sum of Value".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec!["Num".into(), "Sum of Value".into()],
                vec![10.into(), 30.into()],
                vec![2.into(), 20.into()],
                vec![1.into(), 10.into()],
            ]
        );
    }

    #[test]
    fn renders_numeric_row_labels_as_numbers_in_tabular_layout() {
        let data = vec![
            pv_row(&["Num".into(), "Sales".into()]),
            pv_row(&[1.into(), 10.into()]),
            pv_row(&[1.5.into(), 20.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Num")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec!["Num".into(), "Sum of Sales".into()],
                vec![1.into(), 10.into()],
                vec![1.5.into(), 20.into()],
            ]
        );
    }

    #[test]
    fn sorts_column_keys_descending_for_text_field() {
        let data = vec![
            pv_row(&["Region".into(), "Product".into(), "Sales".into()]),
            pv_row(&["East".into(), "A".into(), 100.into()]),
            pv_row(&["East".into(), "B".into(), 150.into()]),
            pv_row(&["West".into(), "A".into(), 200.into()]),
            pv_row(&["West".into(), "B".into(), 250.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![PivotField {
                sort_order: SortOrder::Descending,
                ..PivotField::new("Product")
            }],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec![
                    "Region".into(),
                    "B - Sum of Sales".into(),
                    "A - Sum of Sales".into()
                ],
                vec!["East".into(), 150.into(), 100.into()],
                vec!["West".into(), 250.into(), 200.into()],
            ]
        );
    }

    #[test]
    fn sorts_manual_order_first_then_remaining_ascending() {
        let data = vec![
            pv_row(&["Region".into(), "Sales".into()]),
            pv_row(&["East".into(), 100.into()]),
            pv_row(&["West".into(), 200.into()]),
            pv_row(&["North".into(), 50.into()]),
            pv_row(&["South".into(), 40.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField {
                sort_order: SortOrder::Manual,
                manual_sort: Some(vec![
                    PivotKeyPart::Text("West".to_string()),
                    PivotKeyPart::Text("South".to_string()),
                ]),
                ..PivotField::new("Region")
            }],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec!["Region".into(), "Sum of Sales".into()],
                vec!["West".into(), 200.into()],
                vec!["South".into(), 40.into()],
                vec!["East".into(), 100.into()],
                vec!["North".into(), 50.into()],
            ]
        );
    }

    #[test]
    fn sorts_bool_descending_and_keeps_blanks_last() {
        let data = vec![
            pv_row(&["Flag".into(), "Sales".into()]),
            pv_row(&[false.into(), 1.into()]),
            pv_row(&[true.into(), 2.into()]),
            pv_row(&[PivotValue::Blank, 3.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField {
                sort_order: SortOrder::Descending,
                ..PivotField::new("Flag")
            }],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        // Descending: true before false. Blank is always last.
        assert_eq!(
            result.data,
            vec![
                vec!["Flag".into(), "Sum of Sales".into()],
                vec![true.into(), 2.into()],
                vec![false.into(), 1.into()],
                vec!["(blank)".into(), 3.into()],
            ]
        );
    }

    #[test]
    fn sorts_dates_descending_and_keeps_blanks_last() {
        let jan_01 = NaiveDate::from_ymd_opt(2024, 1, 1).unwrap();
        let jan_02 = NaiveDate::from_ymd_opt(2024, 1, 2).unwrap();

        let data = vec![
            pv_row(&["Date".into(), "Sales".into()]),
            pv_row(&[jan_01.into(), 10.into()]),
            pv_row(&[jan_02.into(), 20.into()]),
            pv_row(&[PivotValue::Blank, 5.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField {
                sort_order: SortOrder::Descending,
                ..PivotField::new("Date")
            }],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        // Descending: newest date first. Blank is always last.
        assert_eq!(
            result.data,
            vec![
                vec!["Date".into(), "Sum of Sales".into()],
                vec![jan_02.into(), 20.into()],
                vec![jan_01.into(), 10.into()],
                vec!["(blank)".into(), 5.into()],
            ]
        );
    }

    #[test]
    fn produces_basic_subtotals_for_multiple_row_fields() {
        let data = vec![
            pv_row(&["Region".into(), "Product".into(), "Sales".into()]),
            pv_row(&["East".into(), "A".into(), 100.into()]),
            pv_row(&["East".into(), "B".into(), 150.into()]),
            pv_row(&["West".into(), "A".into(), 200.into()]),
            pv_row(&["West".into(), "B".into(), 250.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region"), PivotField::new("Product")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::Bottom,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec!["Region".into(), "Product".into(), "Sum of Sales".into()],
                vec!["East".into(), "A".into(), 100.into()],
                vec!["East".into(), "B".into(), 150.into()],
                vec!["East Total".into(), PivotValue::Blank, 250.into()],
                vec!["West".into(), "A".into(), 200.into()],
                vec!["West".into(), "B".into(), 250.into()],
                vec!["West Total".into(), PivotValue::Blank, 450.into()],
                vec!["Grand Total".into(), PivotValue::Blank, 700.into()],
            ]
        );
    }

    #[test]
    fn displays_blank_dimension_items_as_excel_blank_in_row_and_column_labels() {
        let data = vec![
            pv_row(&["Region".into(), "Product".into(), "Sales".into()]),
            pv_row(&[PivotValue::Blank, "A".into(), 10.into()]),
            pv_row(&["East".into(), PivotValue::Blank, 20.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![PivotField::new("Product")],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec![
                    "Region".into(),
                    "A - Sum of Sales".into(),
                    "(blank) - Sum of Sales".into(),
                ],
                vec!["East".into(), PivotValue::Blank, 20.into()],
                vec!["(blank)".into(), 10.into(), PivotValue::Blank],
            ]
        );
    }

    #[test]
    fn produces_top_subtotals_for_multiple_row_fields() {
        let data = vec![
            pv_row(&["Region".into(), "Product".into(), "Sales".into()]),
            pv_row(&["East".into(), "A".into(), 100.into()]),
            pv_row(&["East".into(), "B".into(), 150.into()]),
            pv_row(&["West".into(), "A".into(), 200.into()]),
            pv_row(&["West".into(), "B".into(), 250.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region"), PivotField::new("Product")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::Top,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec!["Region".into(), "Product".into(), "Sum of Sales".into()],
                vec!["East Total".into(), PivotValue::Blank, 250.into()],
                vec!["East".into(), "A".into(), 100.into()],
                vec!["East".into(), "B".into(), 150.into()],
                vec!["West Total".into(), PivotValue::Blank, 450.into()],
                vec!["West".into(), "A".into(), 200.into()],
                vec!["West".into(), "B".into(), 250.into()],
                vec!["Grand Total".into(), PivotValue::Blank, 700.into()],
            ]
        );
    }

    #[test]
    fn places_nested_subtotal_labels_in_correct_row_field_column() {
        let data = vec![
            pv_row(&[
                "Region".into(),
                "Product".into(),
                "Month".into(),
                "Sales".into(),
            ]),
            pv_row(&["East".into(), "A".into(), "1".into(), 100.into()]),
            pv_row(&["East".into(), "A".into(), "2".into(), 150.into()]),
            pv_row(&["East".into(), "B".into(), "1".into(), 200.into()]),
            pv_row(&["West".into(), "A".into(), "1".into(), 250.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![
                PivotField::new("Region"),
                PivotField::new("Product"),
                PivotField::new("Month"),
            ],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::Bottom,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec![
                    "Region".into(),
                    "Product".into(),
                    "Month".into(),
                    "Sum of Sales".into()
                ],
                vec!["East".into(), "A".into(), "1".into(), 100.into()],
                vec!["East".into(), "A".into(), "2".into(), 150.into()],
                vec![
                    "East".into(),
                    "A Total".into(),
                    PivotValue::Blank,
                    250.into()
                ],
                vec!["East".into(), "B".into(), "1".into(), 200.into()],
                vec![
                    "East".into(),
                    "B Total".into(),
                    PivotValue::Blank,
                    200.into()
                ],
                vec![
                    "East Total".into(),
                    PivotValue::Blank,
                    PivotValue::Blank,
                    450.into()
                ],
                vec!["West".into(), "A".into(), "1".into(), 250.into()],
                vec![
                    "West".into(),
                    "A Total".into(),
                    PivotValue::Blank,
                    250.into()
                ],
                vec![
                    "West Total".into(),
                    PivotValue::Blank,
                    PivotValue::Blank,
                    250.into()
                ],
                vec![
                    "Grand Total".into(),
                    PivotValue::Blank,
                    PivotValue::Blank,
                    700.into()
                ],
            ]
        );
    }

    #[test]
    fn show_as_percent_of_grand_total() {
        let data = vec![
            pv_row(&["Region".into(), "Sales".into()]),
            pv_row(&["East".into(), 1.into()]),
            pv_row(&["West".into(), 3.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::PercentOfGrandTotal),
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec!["Region".into(), "Sum of Sales".into()],
                vec!["East".into(), 0.25.into()],
                vec!["West".into(), 0.75.into()],
                vec!["Grand Total".into(), 1.0.into()],
            ]
        );
    }

    #[test]
    fn sorts_numeric_row_keys_by_numeric_value() {
        let data = vec![
            pv_row(&["Num".into(), "Sales".into()]),
            pv_row(&[2.into(), 1.into()]),
            pv_row(&[10.into(), 1.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Num")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec!["Num".into(), "Sum of Sales".into()],
                vec![2.into(), 1.into()],
                vec![10.into(), 1.into()],
            ]
        );
    }

    #[test]
    fn renders_blank_row_items_as_excel_blank_text() {
        let data = vec![
            pv_row(&["Key".into(), "Sales".into()]),
            pv_row(&[PivotValue::Blank, 10.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Key")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
                vec![
                    vec!["Key".into(), "Sum of Sales".into()],
                    vec!["(blank)".into(), 10.into()],
            ]
        );
    }

    #[test]
    fn show_as_percent_of_row_total() {
        let data = vec![
            pv_row(&["Region".into(), "Product".into(), "Sales".into()]),
            pv_row(&["East".into(), "A".into(), 1.into()]),
            pv_row(&["East".into(), "B".into(), 1.into()]),
            pv_row(&["West".into(), "A".into(), 3.into()]),
            pv_row(&["West".into(), "B".into(), 1.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![PivotField::new("Product")],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::PercentOfRowTotal),
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: true,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec![
                    "Region".into(),
                    "A - Sum of Sales".into(),
                    "B - Sum of Sales".into(),
                    "Grand Total - Sum of Sales".into(),
                ],
                vec!["East".into(), 0.5.into(), 0.5.into(), 1.0.into()],
                vec!["West".into(), 0.75.into(), 0.25.into(), 1.0.into()],
            ]
        );
    }

    #[test]
    fn show_as_percent_of_column_total() {
        let data = vec![
            pv_row(&["Region".into(), "Product".into(), "Sales".into()]),
            pv_row(&["East".into(), "A".into(), 1.into()]),
            pv_row(&["East".into(), "B".into(), 1.into()]),
            pv_row(&["West".into(), "A".into(), 3.into()]),
            pv_row(&["West".into(), "B".into(), 1.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![PivotField::new("Product")],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::PercentOfColumnTotal),
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec![
                    "Region".into(),
                    "A - Sum of Sales".into(),
                    "B - Sum of Sales".into(),
                ],
                vec!["East".into(), 0.25.into(), 0.5.into()],
                vec!["West".into(), 0.75.into(), 0.5.into()],
                vec!["Grand Total".into(), 1.0.into(), 1.0.into()],
            ]
        );
    }

    #[test]
    fn show_as_percent_of_base_item_row_field() {
        let data = vec![
            pv_row(&["Year".into(), "Sales".into()]),
            pv_row(&["2019".into(), 2.into()]),
            pv_row(&["2020".into(), 6.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Year")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::PercentOf),
                base_field: Some(cache_field("Year")),
                base_item: Some("2019".to_string()),
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data,
            vec![
                vec!["Year".into(), "Sum of Sales".into()],
                vec!["2019".into(), 1.0.into()],
                vec!["2020".into(), 3.0.into()],
            ]
        );
    }

    #[test]
    fn show_as_percent_difference_from_base_item_row_field() {
        let data = vec![
            pv_row(&["Year".into(), "Sales".into()]),
            pv_row(&["2019".into(), 2.into()]),
            pv_row(&["2020".into(), 6.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Year")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::PercentDifferenceFrom),
                base_field: Some(cache_field("Year")),
                base_item: Some("2019".to_string()),
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data,
            vec![
                vec!["Year".into(), "Sum of Sales".into()],
                vec!["2019".into(), 0.0.into()],
                vec!["2020".into(), 2.0.into()],
            ]
        );
    }

    #[test]
    fn show_as_percent_of_base_item_column_field() {
        let data = vec![
            pv_row(&["Region".into(), "Year".into(), "Sales".into()]),
            pv_row(&["East".into(), "2019".into(), 2.into()]),
            pv_row(&["East".into(), "2020".into(), 6.into()]),
            pv_row(&["West".into(), "2019".into(), 4.into()]),
            pv_row(&["West".into(), "2020".into(), 8.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![PivotField::new("Year")],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::PercentOf),
                base_field: Some(cache_field("Year")),
                base_item: Some("2019".to_string()),
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data,
            vec![
                vec![
                    "Region".into(),
                    "2019 - Sum of Sales".into(),
                    "2020 - Sum of Sales".into(),
                ],
                vec!["East".into(), 1.0.into(), 3.0.into()],
                vec!["West".into(), 1.0.into(), 2.0.into()],
            ]
        );
    }

    #[test]
    fn show_as_percent_difference_from_base_item_column_field() {
        let data = vec![
            pv_row(&["Region".into(), "Year".into(), "Sales".into()]),
            pv_row(&["East".into(), "2019".into(), 2.into()]),
            pv_row(&["East".into(), "2020".into(), 6.into()]),
            pv_row(&["West".into(), "2019".into(), 4.into()]),
            pv_row(&["West".into(), "2020".into(), 8.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![PivotField::new("Year")],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::PercentDifferenceFrom),
                base_field: Some(cache_field("Year")),
                base_item: Some("2019".to_string()),
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec![
                    "Region".into(),
                    "2019 - Sum of Sales".into(),
                    "2020 - Sum of Sales".into(),
                ],
                vec!["East".into(), 0.0.into(), 2.0.into()],
                vec!["West".into(), 0.0.into(), 1.0.into()],
            ]
        );
    }

    #[test]
    fn show_as_percent_of_missing_base_field_blanks_values() {
        let data = vec![
            pv_row(&["Year".into(), "Sales".into()]),
            pv_row(&["2019".into(), 2.into()]),
            pv_row(&["2020".into(), 6.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Year")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::PercentOf),
                base_field: None, // invalid (required)
                base_item: Some("2019".to_string()),
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data,
            vec![
                vec!["Year".into(), "Sum of Sales".into()],
                vec!["2019".into(), PivotValue::Blank],
                vec!["2020".into(), PivotValue::Blank],
            ]
        );
    }

    #[test]
    fn show_as_percent_difference_from_unknown_base_item_blanks_values() {
        let data = vec![
            pv_row(&["Year".into(), "Sales".into()]),
            pv_row(&["2019".into(), 2.into()]),
            pv_row(&["2020".into(), 6.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Year")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::PercentDifferenceFrom),
                base_field: Some(cache_field("Year")),
                base_item: Some("1900".to_string()), // not found
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data,
            vec![
                vec!["Year".into(), "Sum of Sales".into()],
                vec!["2019".into(), PivotValue::Blank],
                vec!["2020".into(), PivotValue::Blank],
            ]
        );
    }

    #[test]
    fn show_as_percent_of_base_item_row_field_applies_to_subtotals() {
        let data = vec![
            pv_row(&["Region".into(), "Product".into(), "Sales".into()]),
            pv_row(&["East".into(), "A".into(), 100.into()]),
            pv_row(&["East".into(), "B".into(), 200.into()]),
            pv_row(&["West".into(), "A".into(), 50.into()]),
            pv_row(&["West".into(), "B".into(), 150.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region"), PivotField::new("Product")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::PercentOf),
                base_field: Some(cache_field("Region")),
                base_item: Some("East".to_string()),
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::Bottom,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data,
            vec![
                vec!["Region".into(), "Product".into(), "Sum of Sales".into()],
                vec!["East".into(), "A".into(), 1.0.into()],
                vec!["East".into(), "B".into(), 1.0.into()],
                vec!["East Total".into(), PivotValue::Blank, 1.0.into()],
                vec!["West".into(), "A".into(), 0.5.into()],
                vec!["West".into(), "B".into(), 0.75.into()],
                vec![
                    "West Total".into(),
                    PivotValue::Blank,
                    (200.0 / 300.0).into()
                ],
                vec![
                    "Grand Total".into(),
                    PivotValue::Blank,
                    (500.0 / 300.0).into()
                ],
            ]
        );
    }

    #[test]
    fn show_as_percent_of_base_item_column_field_applies_to_subtotals() {
        let data = vec![
            pv_row(&[
                "Region".into(),
                "Product".into(),
                "Year".into(),
                "Sales".into(),
            ]),
            pv_row(&["East".into(), "A".into(), "2019".into(), 10.into()]),
            pv_row(&["East".into(), "A".into(), "2020".into(), 20.into()]),
            pv_row(&["East".into(), "B".into(), "2019".into(), 30.into()]),
            pv_row(&["East".into(), "B".into(), "2020".into(), 60.into()]),
            pv_row(&["West".into(), "A".into(), "2019".into(), 5.into()]),
            pv_row(&["West".into(), "A".into(), "2020".into(), 15.into()]),
            pv_row(&["West".into(), "B".into(), "2019".into(), 25.into()]),
            pv_row(&["West".into(), "B".into(), "2020".into(), 50.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region"), PivotField::new("Product")],
            column_fields: vec![PivotField::new("Year")],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::PercentOf),
                base_field: Some(cache_field("Year")),
                base_item: Some("2019".to_string()),
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::Bottom,
            grand_totals: GrandTotals {
                rows: true,
                columns: true,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data,
            vec![
                vec![
                    "Region".into(),
                    "Product".into(),
                    "2019 - Sum of Sales".into(),
                    "2020 - Sum of Sales".into(),
                    "Grand Total - Sum of Sales".into(),
                ],
                vec![
                    "East".into(),
                    "A".into(),
                    1.0.into(),
                    2.0.into(),
                    3.0.into()
                ],
                vec![
                    "East".into(),
                    "B".into(),
                    1.0.into(),
                    2.0.into(),
                    3.0.into()
                ],
                vec![
                    "East Total".into(),
                    PivotValue::Blank,
                    1.0.into(),
                    2.0.into(),
                    3.0.into(),
                ],
                vec![
                    "West".into(),
                    "A".into(),
                    1.0.into(),
                    3.0.into(),
                    4.0.into()
                ],
                vec![
                    "West".into(),
                    "B".into(),
                    1.0.into(),
                    2.0.into(),
                    3.0.into()
                ],
                vec![
                    "West Total".into(),
                    PivotValue::Blank,
                    1.0.into(),
                    (65.0 / 30.0).into(),
                    (95.0 / 30.0).into(),
                ],
                vec![
                    "Grand Total".into(),
                    PivotValue::Blank,
                    1.0.into(),
                    (145.0 / 70.0).into(),
                    (215.0 / 70.0).into(),
                ],
            ]
        );
    }

    #[test]
    fn show_as_running_total() {
        let data = vec![
            pv_row(&["Item".into(), "Sales".into()]),
            pv_row(&["A".into(), 1.into()]),
            pv_row(&["B".into(), 2.into()]),
            pv_row(&["C".into(), 3.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Item")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::RunningTotal),
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec!["Item".into(), "Sum of Sales".into()],
                vec!["A".into(), 1.into()],
                vec!["B".into(), 3.into()],
                vec!["C".into(), 6.into()],
            ]
        );
    }

    #[test]
    fn show_as_running_total_respects_row_base_field() {
        let data = vec![
            pv_row(&["Region".into(), "Year".into(), "Sales".into()]),
            pv_row(&["East".into(), "2019".into(), 2.into()]),
            pv_row(&["East".into(), "2020".into(), 6.into()]),
            pv_row(&["West".into(), "2019".into(), 4.into()]),
            pv_row(&["West".into(), "2020".into(), 8.into()]),
        ];
        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region"), PivotField::new("Year")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::RunningTotal),
                base_field: Some(cache_field("Year")),
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data,
            vec![
                vec!["Region".into(), "Year".into(), "Sum of Sales".into()],
                vec!["East".into(), "2019".into(), 2.into()],
                vec!["East".into(), "2020".into(), 8.into()],
                vec!["West".into(), "2019".into(), 4.into()],
                vec!["West".into(), "2020".into(), 12.into()],
            ]
        );
    }

    #[test]
    fn show_as_running_total_respects_column_base_field() {
        let data = vec![
            pv_row(&["Region".into(), "Year".into(), "Sales".into()]),
            pv_row(&["East".into(), "2019".into(), 2.into()]),
            pv_row(&["East".into(), "2020".into(), 6.into()]),
            pv_row(&["West".into(), "2019".into(), 4.into()]),
            pv_row(&["West".into(), "2020".into(), 8.into()]),
        ];
        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![PivotField::new("Year")],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::RunningTotal),
                base_field: Some(cache_field("Year")),
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data,
            vec![
                vec![
                    "Region".into(),
                    "2019 - Sum of Sales".into(),
                    "2020 - Sum of Sales".into(),
                ],
                vec!["East".into(), 2.into(), 8.into()],
                vec!["West".into(), 4.into(), 12.into()],
            ]
        );
    }

    #[test]
    fn show_as_rank_ascending() {
        let data = vec![
            pv_row(&["Item".into(), "Sales".into()]),
            pv_row(&["A".into(), 30.into()]),
            pv_row(&["B".into(), 10.into()]),
            pv_row(&["C".into(), 20.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Item")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::RankAscending),
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec!["Item".into(), "Sum of Sales".into()],
                vec!["A".into(), 3.into()],
                vec!["B".into(), 1.into()],
                vec!["C".into(), 2.into()],
            ]
        );
    }

    #[test]
    fn show_as_rank_respects_row_base_field() {
        let data = vec![
            pv_row(&["Region".into(), "Year".into(), "Sales".into()]),
            pv_row(&["East".into(), "2019".into(), 2.into()]),
            pv_row(&["East".into(), "2020".into(), 6.into()]),
            pv_row(&["West".into(), "2019".into(), 4.into()]),
            pv_row(&["West".into(), "2020".into(), 8.into()]),
        ];
        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region"), PivotField::new("Year")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::RankDescending),
                base_field: Some(cache_field("Year")),
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data,
            vec![
                vec!["Region".into(), "Year".into(), "Sum of Sales".into()],
                vec!["East".into(), "2019".into(), 2.into()],
                vec!["East".into(), "2020".into(), 1.into()],
                vec!["West".into(), "2019".into(), 2.into()],
                vec!["West".into(), "2020".into(), 1.into()],
            ]
        );
    }

    #[test]
    fn show_as_rank_respects_column_base_field() {
        let data = vec![
            pv_row(&["Region".into(), "Year".into(), "Sales".into()]),
            pv_row(&["East".into(), "2019".into(), 2.into()]),
            pv_row(&["East".into(), "2020".into(), 6.into()]),
            pv_row(&["West".into(), "2019".into(), 4.into()]),
            pv_row(&["West".into(), "2020".into(), 8.into()]),
        ];
        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![PivotField::new("Year")],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::RankDescending),
                base_field: Some(cache_field("Year")),
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data,
            vec![
                vec![
                    "Region".into(),
                    "2019 - Sum of Sales".into(),
                    "2020 - Sum of Sales".into(),
                ],
                vec!["East".into(), 2.into(), 1.into()],
                vec!["West".into(), 2.into(), 1.into()],
            ]
        );
    }

    #[test]
    fn show_as_rank_descending() {
        let data = vec![
            pv_row(&["Item".into(), "Sales".into()]),
            pv_row(&["A".into(), 30.into()]),
            pv_row(&["B".into(), 10.into()]),
            pv_row(&["C".into(), 20.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Item")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Sales"),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::RankDescending),
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();

        assert_eq!(
            result.data,
            vec![
                vec!["Item".into(), "Sum of Sales".into()],
                vec!["A".into(), 1.into()],
                vec!["B".into(), 3.into()],
                vec!["C".into(), 2.into()],
            ]
        );
    }

    #[test]
    fn unique_values_are_type_aware_and_collision_free() {
        let data = vec![
            pv_row(&["Key".into()]),
            pv_row(&[1.into()]),
            pv_row(&["1".into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let values = cache.unique_values.get("Key").unwrap();
        assert_eq!(values.len(), 2);
        assert!(matches!(values[0], PivotValue::Number(n) if n == 1.0));
        assert!(matches!(&values[1], PivotValue::Text(s) if s == "1"));
    }

    #[test]
    fn pivot_order_is_deterministic_when_display_strings_collide() {
        let data = vec![
            pv_row(&["Key".into(), "Amount".into()]),
            pv_row(&[PivotValue::Number(1.0), 10.into()]),
            pv_row(&[PivotValue::Text("1".to_string()), 20.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Key")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: cache_field("Amount"),
                name: "Sum of Amount".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let expected = vec![
            vec!["Key".into(), "Sum of Amount".into()],
            vec![1.into(), 10.into()],
            vec!["1".into(), 20.into()],
        ];

        for _ in 0..32 {
            let result = PivotEngine::calculate(&cache, &cfg).unwrap();
            assert_eq!(result.data, expected);
        }
    }

    #[test]
    fn pivot_cache_normalizes_blank_and_duplicate_headers() {
        let data = vec![
            pv_row(&[
                PivotValue::Text(" ".to_string()),
                PivotValue::Text(" Sales ".to_string()),
                PivotValue::Text("sales".to_string()),
                PivotValue::Blank,
                PivotValue::Text("Region".to_string()),
                PivotValue::Text("REGION".to_string()),
            ]),
            pv_row(&[
                PivotValue::Text("East".to_string()),
                10.into(),
                20.into(),
                30.into(),
                PivotValue::Text("X".to_string()),
                PivotValue::Text("Y".to_string()),
            ]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        let field_names: Vec<String> = cache.fields.iter().map(|f| f.name.clone()).collect();
        assert_eq!(
            field_names,
            vec![
                "Column1".to_string(),
                "Sales".to_string(),
                "sales (2)".to_string(),
                "Column2".to_string(),
                "Region".to_string(),
                "REGION (2)".to_string(),
            ]
        );

        let folded: HashSet<String> = field_names
            .iter()
            .map(|s| fold_text_case_insensitive(s))
            .collect();
        assert_eq!(folded.len(), field_names.len());

        assert_eq!(cache.unique_values.len(), field_names.len());
        for name in &field_names {
            assert!(cache.unique_values.contains_key(name));
        }
    }

    #[test]
    fn pivot_cache_normalizes_unicode_duplicate_headers_case_insensitively() {
        // Use a German sharp S (ß) to ensure we handle Unicode-aware case-insensitive header
        // normalization (ß -> SS).
        let data = vec![
            pv_row(&["Straße".into(), "STRASSE".into()]),
            pv_row(&[1.into(), 2.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();
        let field_names: Vec<String> = cache.fields.iter().map(|f| f.name.clone()).collect();
        assert_eq!(
            field_names,
            vec!["Straße".to_string(), "STRASSE (2)".to_string()]
        );

        let folded: HashSet<String> = field_names
            .iter()
            .map(|s| fold_text_case_insensitive(s))
            .collect();
        assert_eq!(folded.len(), field_names.len());
    }

    #[test]
    fn pivot_engine_can_reference_normalized_field_names() {
        let data = vec![
            pv_row(&[PivotValue::Blank, "Sales".into(), "Sales".into()]),
            pv_row(&["East".into(), 100.into(), 200.into()]),
            pv_row(&["West".into(), 300.into(), 400.into()]),
        ];

        let cache = PivotCache::from_range(&data).unwrap();

        // The cache should expose non-empty, unique names.
        assert_eq!(
            cache
                .fields
                .iter()
                .map(|f| f.name.as_str())
                .collect::<Vec<_>>(),
            vec!["Column1", "Sales", "Sales (2)"]
        );

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Column1")],
            column_fields: vec![],
            value_fields: vec![
                ValueField {
                    source_field: cache_field("Sales"),
                    name: "Sum of Sales".to_string(),
                    aggregation: AggregationType::Sum,
                    number_format: None,
                    show_as: None,
                    base_field: None,
                    base_item: None,
                },
                ValueField {
                    source_field: cache_field("Sales (2)"),
                    name: "Sum of Sales 2".to_string(),
                    aggregation: AggregationType::Sum,
                    number_format: None,
                    show_as: None,
                    base_field: None,
                    base_item: None,
                },
            ],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let result = PivotEngine::calculate(&cache, &cfg).unwrap();
        assert_eq!(
            result.data,
            vec![
                vec![
                    "Column1".into(),
                    "Sum of Sales".into(),
                    "Sum of Sales 2".into()
                ],
                vec!["East".into(), 100.into(), 200.into()],
                vec!["West".into(), 300.into(), 400.into()],
                vec!["Grand Total".into(), 400.into(), 600.into()],
            ]
        );
    }
    #[test]
    fn calculates_from_columnar_table_source() {
        let schema = vec![
            ColumnSchema {
                name: "Region".to_string(),
                column_type: ColumnType::String,
            },
            ColumnSchema {
                name: "Product".to_string(),
                column_type: ColumnType::String,
            },
            ColumnSchema {
                name: "Sales".to_string(),
                column_type: ColumnType::Number,
            },
        ];
        let options = TableOptions {
            page_size_rows: 1024,
            cache: PageCacheConfig { max_entries: 8 },
        };

        let mut builder = ColumnarTableBuilder::new(schema, options);
        let east = Arc::<str>::from("East");
        let west = Arc::<str>::from("West");
        let a = Arc::<str>::from("A");
        let b = Arc::<str>::from("B");
        builder.append_row(&[
            Value::String(east.clone()),
            Value::String(a.clone()),
            Value::Number(100.0),
        ]);
        builder.append_row(&[
            Value::String(east.clone()),
            Value::String(b.clone()),
            Value::Number(150.0),
        ]);
        builder.append_row(&[
            Value::String(west.clone()),
            Value::String(a.clone()),
            Value::Number(200.0),
        ]);
        builder.append_row(&[
            Value::String(west.clone()),
            Value::String(b.clone()),
            Value::Number(250.0),
        ]);

        let table = builder.finalize();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![PivotField::new("Product")],
             value_fields: vec![ValueField {
                 source_field: cache_field("Sales"),
                 name: "Sum of Sales".to_string(),
                 aggregation: AggregationType::Sum,
                 number_format: None,
                 show_as: None,
                 base_field: None,
                 base_item: None,
             }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: true,
            },
        };

        let result = PivotEngine::calculate_streaming(&table, &cfg).unwrap();
        assert_eq!(
            result.data,
            vec![
                vec![
                    "Region".into(),
                    "A - Sum of Sales".into(),
                    "B - Sum of Sales".into(),
                    "Grand Total - Sum of Sales".into()
                ],
                vec!["East".into(), 100.into(), 150.into(), 250.into()],
                vec!["West".into(), 200.into(), 250.into(), 450.into()],
                vec!["Grand Total".into(), 300.into(), 400.into(), 700.into()],
            ]
        );
    }

    /// This is a manual benchmark/test intended for investigating performance/memory
    /// characteristics on large datasets.
    ///
    /// Run with:
    /// `cargo test -p formula-engine pivot::tests::pivot_streaming_large_dataset_bench -- --ignored --nocapture`
    #[test]
    #[ignore]
    fn pivot_streaming_large_dataset_bench() {
        use std::time::Instant;

        let rows: usize = std::env::var("FORMULA_PIVOT_BENCH_ROWS")
            .ok()
            .and_then(|v| v.parse().ok())
            .unwrap_or(1_000_000);

        let schema = vec![
            ColumnSchema {
                name: "Cat".to_string(),
                column_type: ColumnType::Number,
            },
            ColumnSchema {
                name: "Amount".to_string(),
                column_type: ColumnType::Number,
            },
        ];
        let options = TableOptions {
            page_size_rows: 65_536,
            cache: PageCacheConfig { max_entries: 16 },
        };

        let mut builder = ColumnarTableBuilder::new(schema, options);
        for i in 0..rows {
            builder.append_row(&[
                Value::Number((i % 1000) as f64),
                Value::Number((i % 100) as f64),
            ]);
        }
        let table = builder.finalize();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Cat")],
            column_fields: vec![],
             value_fields: vec![ValueField {
                 source_field: cache_field("Amount"),
                 name: "Sum of Amount".to_string(),
                 aggregation: AggregationType::Sum,
                 number_format: None,
                 show_as: None,
                 base_field: None,
                 base_item: None,
             }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let start = Instant::now();
        let result = PivotEngine::calculate_streaming(&table, &cfg).unwrap();
        let elapsed = start.elapsed();
        println!(
            "columnar streaming: rows={rows} output_rows={} table_compressed_bytes={} elapsed_ms={:.2}",
            result.data.len(),
            table.compressed_size_bytes(),
            elapsed.as_secs_f64() * 1000.0
        );

        // Optional baseline: build the existing row-backed cache and compute the same pivot.
        // This can be memory-heavy; use a smaller `FORMULA_PIVOT_BENCH_ROWS` if it OOMs.
        let baseline = std::env::var("FORMULA_PIVOT_BENCH_ROW_BACKED")
            .ok()
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(true);
        if !baseline {
            return;
        }

        let start = Instant::now();
        let mut range: Vec<Vec<PivotValue>> = Vec::with_capacity(rows + 1);
        range.push(vec!["Cat".into(), "Amount".into()]);
        for i in 0..rows {
            range.push(vec![((i % 1000) as i64).into(), ((i % 100) as i64).into()]);
        }
        let cache = PivotCache::from_range(&range).unwrap();
        drop(range);
        let build_elapsed = start.elapsed();

        let estimated_bytes = cache.records.capacity() * std::mem::size_of::<Vec<PivotValue>>()
            + cache
                .records
                .iter()
                .map(|r| r.capacity() * std::mem::size_of::<PivotValue>())
                .sum::<usize>();

        let start = Instant::now();
        let result2 = PivotEngine::calculate(&cache, &cfg).unwrap();
        let calc_elapsed = start.elapsed();

        assert_eq!(result2.data.len(), result.data.len());
        println!(
            "row-backed cache: rows={rows} estimated_record_bytes={} build_ms={:.2} calc_ms={:.2}",
            estimated_bytes,
            build_elapsed.as_secs_f64() * 1000.0,
            calc_elapsed.as_secs_f64() * 1000.0
        );
    }
}
