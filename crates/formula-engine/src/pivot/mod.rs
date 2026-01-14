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
use std::cmp::Ordering;
use std::collections::{BTreeMap, HashMap, HashSet};

pub use formula_model::pivots::{
    AggregationType, CalculatedField, CalculatedItem, FilterField, GrandTotals, Layout, PivotConfig,
    PivotField, PivotKeyPart, PivotValue, ShowAsType, SortOrder, SubtotalPosition, ValueField,
};
use serde::{Deserialize, Serialize};
use thiserror::Error;

use std::sync::atomic::{AtomicU64, Ordering as AtomicOrdering};

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
    #[error("invalid calculated field formula for {field}: {message}")]
    InvalidCalculatedFieldFormula { field: String, message: String },
    #[error("invalid calculated item formula for {field}::{item}: {message}")]
    InvalidCalculatedItemFormula {
        field: String,
        item: String,
        message: String,
    },
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
            if used_folded.contains(&name.to_ascii_lowercase()) {
                let mut suffix = 2usize;
                loop {
                    name = format!("{base} ({suffix})");
                    let folded = name.to_ascii_lowercase();
                    if !used_folded.contains(&folded) {
                        break;
                    }
                    suffix += 1;
                }
            }

            used_folded.insert(name.to_ascii_lowercase());
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
            fields.push(CacheField {
                name,
                index: idx,
            });
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
            row_fields: self.row_fields.into_iter().map(PivotField::new).collect(),
            column_fields: self
                .column_fields
                .into_iter()
                .map(PivotField::new)
                .collect(),
            value_fields: self
                .value_fields
                .into_iter()
                .map(|vf| ValueField {
                    name: vf
                        .name
                        .unwrap_or_else(|| format!("{:?} of {}", vf.aggregation, vf.field)),
                    source_field: vf.field,
                    aggregation: vf.aggregation,
                    number_format: None,
                    show_as: None,
                    base_field: None,
                    base_item: None,
                })
                .collect(),
            filter_fields: self
                .filter_fields
                .into_iter()
                .map(|f| FilterField {
                    source_field: f.field,
                    allowed: f
                        .allowed
                        .map(|vals| vals.into_iter().map(|v| v.to_key_part()).collect()),
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

pub struct PivotEngine;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PivotRowKind {
    Header,
    Leaf,
    Subtotal,
    GrandTotal,
}

impl PivotEngine {
    pub fn calculate(cache: &PivotCache, cfg: &PivotConfig) -> Result<PivotResult, PivotError> {
        if cfg.value_fields.is_empty() {
            return Err(PivotError::NoValueFields);
        }

        let indices = FieldIndices::new(cache, cfg)?;

        let mut cube: HashMap<PivotKey, HashMap<PivotKey, Vec<Accumulator>>> = HashMap::new();
        let mut row_keys: HashSet<PivotKey> = HashSet::new();
        let mut col_keys: HashSet<PivotKey> = HashSet::new();

        for record in &cache.records {
            if !indices.passes_filters(record, cfg) {
                continue;
            }

            let row_key = indices.build_key(record, &indices.row_indices);
            let col_key = indices.build_key(record, &indices.col_indices);

            row_keys.insert(row_key.clone());
            col_keys.insert(col_key.clone());

            let row_entry = cube.entry(row_key).or_default();
            let cell = row_entry.entry(col_key).or_insert_with(|| {
                (0..cfg.value_fields.len())
                    .map(|_| Accumulator::new())
                    .collect()
            });

            for (vf_idx, _vf) in cfg.value_fields.iter().enumerate() {
                let val = record
                    .get(indices.value_indices[vf_idx])
                    .unwrap_or(&PivotValue::Blank);
                cell[vf_idx].update(val);
            }
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
                for row_key in &row_keys {
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
                            row_kinds.push(PivotRowKind::Subtotal);
                        }
                    }

                    let row_map = cube.get(row_key);
                    data.push(Self::render_row(
                        row_key, row_map, &col_keys, cfg, /*label*/ None,
                    ));
                    row_kinds.push(PivotRowKind::Leaf);

                    if let Some(acc) = grand_acc.as_mut() {
                        acc.merge_row(row_map, &col_keys, cfg.value_fields.len());
                    }

                    prev_row_key = Some(row_key.clone());
                }
            }
            SubtotalPosition::Bottom if subtotal_levels > 0 => {
                let mut group_accs: Vec<Option<GroupAccumulator>> = vec![None; subtotal_levels];

                let mut prev_row_key: Option<PivotKey> = None;
                for row_key in &row_keys {
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
                    row_kinds.push(PivotRowKind::Leaf);

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
                for row_key in &row_keys {
                    let row_map = cube.get(row_key);
                    data.push(Self::render_row(
                        row_key, row_map, &col_keys, cfg, /*label*/ None,
                    ));
                    row_kinds.push(PivotRowKind::Leaf);
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

    fn build_header_row(col_keys: &[PivotKey], cfg: &PivotConfig) -> Vec<PivotValue> {
        let mut row = Vec::new();

        match cfg.layout {
            Layout::Compact => {
                row.push(PivotValue::Text("Row Labels".to_string()));
            }
            Layout::Outline | Layout::Tabular => {
                for f in &cfg.row_fields {
                    row.push(PivotValue::Text(f.source_field.clone()));
                }
            }
        }

        // Flatten column keys × value fields.
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
                    format!("{:?} of {}", vf.aggregation, vf.source_field)
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
                    format!("{:?} of {}", vf.aggregation, vf.source_field)
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
                // Compact: join row keys into one cell.
                let s = row_key
                    .display_strings()
                    .into_iter()
                    .filter(|s| !s.is_empty())
                    .collect::<Vec<_>>()
                    .join(" / ");
                row.push(label.unwrap_or_else(|| PivotValue::Text(s)));
            }
            Layout::Outline | Layout::Tabular => {
                for (idx, part) in row_key.0.iter().enumerate() {
                    if idx == 0 {
                        if let Some(l) = label.as_ref() {
                            row.push(l.clone());
                            continue;
                        }
                    }
                    row.push(PivotValue::Text(part.display_string()));
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
                            let prefix = prefix_parts
                                .get(idx)
                                .map(|p| p.display_string())
                                .unwrap_or_default();
                            row.push(PivotValue::Text(prefix));
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
                row_kinds.push(PivotRowKind::Subtotal);
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

        let leaf_rows: Vec<usize> = row_kinds
            .iter()
            .enumerate()
            .filter_map(|(idx, kind)| (*kind == PivotRowKind::Leaf).then_some(idx))
            .collect();

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
                    Self::apply_running_total(data, &leaf_rows, &cols);
                }
                ShowAsType::RankAscending => {
                    Self::apply_rank(data, &leaf_rows, &cols, /*descending*/ false);
                }
                ShowAsType::RankDescending => {
                    Self::apply_rank(data, &leaf_rows, &cols, /*descending*/ true);
                }
                // Not implemented yet.
                ShowAsType::PercentOf | ShowAsType::PercentDifferenceFrom => {}
                ShowAsType::Normal => {}
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
    fn new(cache: &PivotCache, cfg: &PivotConfig) -> Result<Self, PivotError> {
        let mut row_indices = Vec::new();
        for f in &cfg.row_fields {
            row_indices.push(
                cache
                    .field_index(&f.source_field)
                    .ok_or_else(|| PivotError::MissingField(f.source_field.clone()))?,
            );
        }
        let mut col_indices = Vec::new();
        for f in &cfg.column_fields {
            col_indices.push(
                cache
                    .field_index(&f.source_field)
                    .ok_or_else(|| PivotError::MissingField(f.source_field.clone()))?,
            );
        }
        let mut value_indices = Vec::new();
        for f in &cfg.value_fields {
            value_indices.push(
                cache
                    .field_index(&f.source_field)
                    .ok_or_else(|| PivotError::MissingField(f.source_field.clone()))?,
            );
        }
        let mut filter_indices = Vec::new();
        for f in &cfg.filter_fields {
            let idx = cache
                .field_index(&f.source_field)
                .ok_or_else(|| PivotError::MissingField(f.source_field.clone()))?;
            filter_indices.push((idx, f.allowed.clone()));
        }
        Ok(Self {
            row_indices,
            col_indices,
            value_indices,
            filter_indices,
        })
    }

    fn build_key(&self, record: &[PivotValue], indices: &[usize]) -> PivotKey {
        PivotKey(
            indices
                .iter()
                .map(|idx| record.get(*idx).unwrap_or(&PivotValue::Blank).to_key_part())
                .collect(),
        )
    }

    fn passes_filters(&self, record: &[PivotValue], _cfg: &PivotConfig) -> bool {
        for (idx, allowed) in &self.filter_indices {
            if let Some(set) = allowed {
                let val = record.get(*idx).unwrap_or(&PivotValue::Blank).to_key_part();
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

    use pretty_assertions::assert_eq;

    fn pv_row(values: &[PivotValue]) -> Vec<PivotValue> {
        values.to_vec()
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
                source_field: "Sales".to_string(),
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
                source_field: "Sales".to_string(),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![FilterField {
                source_field: "Region".to_string(),
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
                source_field: "Sales".to_string(),
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
                source_field: "Value".to_string(),
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
                vec!["10".into(), 30.into()],
                vec!["2".into(), 20.into()],
                vec!["1".into(), 10.into()],
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
                source_field: "Sales".to_string(),
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
                source_field: "Sales".to_string(),
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
                source_field: "Sales".to_string(),
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
                vec!["true".into(), 2.into()],
                vec!["false".into(), 1.into()],
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
                source_field: "Sales".to_string(),
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
                vec!["2024-01-02".into(), 20.into()],
                vec!["2024-01-01".into(), 10.into()],
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
                source_field: "Sales".to_string(),
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
                source_field: "Sales".to_string(),
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
                source_field: "Sales".to_string(),
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
                source_field: "Sales".to_string(),
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
                source_field: "Sales".to_string(),
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
                source_field: "Sales".to_string(),
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
                vec!["2".into(), 1.into()],
                vec!["10".into(), 1.into()],
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
                source_field: "Sales".to_string(),
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
                source_field: "Sales".to_string(),
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
                source_field: "Sales".to_string(),
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
                source_field: "Sales".to_string(),
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
                source_field: "Sales".to_string(),
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
                source_field: "Amount".to_string(),
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
            vec!["1".into(), 10.into()],
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

        let folded: HashSet<String> = field_names.iter().map(|s| s.to_ascii_lowercase()).collect();
        assert_eq!(folded.len(), field_names.len());

        assert_eq!(cache.unique_values.len(), field_names.len());
        for name in &field_names {
            assert!(cache.unique_values.contains_key(name));
        }
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
            cache.fields.iter().map(|f| f.name.as_str()).collect::<Vec<_>>(),
            vec!["Column1", "Sales", "Sales (2)"]
        );

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Column1")],
            column_fields: vec![],
            value_fields: vec![
                ValueField {
                    source_field: "Sales".to_string(),
                    name: "Sum of Sales".to_string(),
                    aggregation: AggregationType::Sum,
                    number_format: None,
                    show_as: None,
                    base_field: None,
                    base_item: None,
                },
                ValueField {
                    source_field: "Sales (2)".to_string(),
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
}
