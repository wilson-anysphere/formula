use chrono::NaiveDate;
use formula_format::{FormatOptions, Value as FmtValue};
use ordered_float::OrderedFloat;
use serde::{Deserialize, Serialize};
use std::cmp::Ordering;
use std::collections::HashMap;
use uuid::Uuid;

mod model;
mod schema;
pub mod slicers;
pub mod workbook;

pub use model::{
    DefinedNameIdentifier, PivotCacheId, PivotDestination, PivotSource, PivotTableModel,
};
pub use schema::{
    parse_dax_column_ref, parse_dax_measure_ref, CalculatedField, CalculatedItem, FilterField,
    GrandTotals, Layout, PivotConfig, PivotFieldRef, SubtotalPosition,
};
pub use workbook::{PivotCacheModel, PivotChartModel, SlicerModel, TimelineModel};

pub type PivotTableId = Uuid;
pub type PivotChartId = Uuid;

/// Sort order applied to pivot fields.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SortOrder {
    #[default]
    Ascending,
    Descending,
    Manual,
}

/// Value representation used for manual pivot-field ordering.
///
/// This is intentionally lightweight and serde-friendly since it may cross IPC
/// boundaries.
#[derive(Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(tag = "type", content = "value", rename_all = "camelCase")]
pub enum PivotKeyPart {
    Blank,
    Number(u64),
    Date(NaiveDate),
    Text(String),
    Bool(bool),
}

fn cmp_text_case_insensitive(a: &str, b: &str) -> Ordering {
    if a.is_ascii() && b.is_ascii() {
        return cmp_ascii_case_insensitive(a, b);
    }

    // Compare using Unicode-aware uppercasing so semantics match Excel-like case-insensitive
    // ordering for non-ASCII text (e.g. ß -> SS).
    let mut a_iter = a.chars().flat_map(|c| c.to_uppercase());
    let mut b_iter = b.chars().flat_map(|c| c.to_uppercase());
    loop {
        match (a_iter.next(), b_iter.next()) {
            (Some(ac), Some(bc)) => match ac.cmp(&bc) {
                Ordering::Equal => continue,
                ord => return ord,
            },
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (None, None) => return Ordering::Equal,
        }
    }
}

fn cmp_ascii_case_insensitive(a: &str, b: &str) -> Ordering {
    let mut a_iter = a.as_bytes().iter();
    let mut b_iter = b.as_bytes().iter();
    loop {
        match (a_iter.next(), b_iter.next()) {
            (Some(&ac), Some(&bc)) => {
                let ac = ac.to_ascii_uppercase();
                let bc = bc.to_ascii_uppercase();
                match ac.cmp(&bc) {
                    Ordering::Equal => continue,
                    ord => return ord,
                }
            }
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (None, None) => return Ordering::Equal,
        }
    }
}
impl PivotKeyPart {
    fn kind_rank(&self) -> u8 {
        match self {
            PivotKeyPart::Number(_) => 0,
            PivotKeyPart::Date(_) => 1,
            PivotKeyPart::Text(_) => 2,
            PivotKeyPart::Bool(_) => 3,
            PivotKeyPart::Blank => 4,
        }
    }

    /// Human-friendly (Excel-like) string representation of a pivot item value.
    pub fn display_string(&self) -> String {
        match self {
            PivotKeyPart::Blank => "(blank)".to_string(),
            PivotKeyPart::Number(bits) => {
                PivotValue::Number(f64::from_bits(*bits)).display_string()
            }
            PivotKeyPart::Date(d) => d.to_string(),
            PivotKeyPart::Text(s) => s.clone(),
            PivotKeyPart::Bool(b) => {
                if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
        }
    }

    /// Converts this key part into the [`PivotValue`] that should be emitted when rendering pivot
    /// item labels into a worksheet grid.
    ///
    /// Notably, Excel renders blank pivot items as the literal "(blank)" text label rather than an
    /// empty cell.
    pub fn to_pivot_value(&self) -> PivotValue {
        match self {
            PivotKeyPart::Blank => PivotValue::Text("(blank)".to_string()),
            PivotKeyPart::Number(bits) => PivotValue::Number(f64::from_bits(*bits)),
            PivotKeyPart::Date(d) => PivotValue::Date(*d),
            PivotKeyPart::Text(s) => PivotValue::Text(s.clone()),
            PivotKeyPart::Bool(b) => PivotValue::Bool(*b),
        }
    }
}

impl PartialOrd for PivotKeyPart {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for PivotKeyPart {
    fn cmp(&self, other: &Self) -> Ordering {
        let rank_cmp = self.kind_rank().cmp(&other.kind_rank());
        if rank_cmp != Ordering::Equal {
            // Excel uses a fixed cross-type ordering (numbers/dates, then text, then booleans,
            // blanks last) regardless of ascending/descending selection. Pivots currently always
            // sort ascending within that global type ordering.
            return rank_cmp;
        }

        match (self, other) {
            (PivotKeyPart::Blank, PivotKeyPart::Blank) => Ordering::Equal,
            (PivotKeyPart::Number(a), PivotKeyPart::Number(b)) => {
                let a = f64::from_bits(*a);
                let b = f64::from_bits(*b);
                a.total_cmp(&b)
            }
            (PivotKeyPart::Date(a), PivotKeyPart::Date(b)) => a.cmp(b),
            (PivotKeyPart::Text(a), PivotKeyPart::Text(b)) => {
                // Pivot tables in Excel sort text case-insensitively by default. Use a casefolded
                // comparison as the primary key, with a deterministic case-sensitive tiebreak so
                // the overall ordering remains total.
                let ord = cmp_text_case_insensitive(a, b);
                if ord != Ordering::Equal {
                    ord
                } else {
                    a.cmp(b)
                }
            }
            (PivotKeyPart::Bool(a), PivotKeyPart::Bool(b)) => a.cmp(b),
            _ => Ordering::Equal,
        }
    }
}

/// Scalar value representation used for pivot caches and pivot output payloads.
///
/// This is the canonical serde format used across engine / XLSX / IPC:
/// a tagged enum in the shape `{ "type": "...", "value": ... }`.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value", rename_all = "camelCase")]
pub enum PivotValue {
    Blank,
    Number(f64),
    /// A calendar date coming from source data / pivot items.
    ///
    /// When rendering into worksheet cell values, this should typically be converted to Excel's
    /// *date serial number* (a number) and paired with a date number format in the styling layer.
    /// Keeping it typed in the pivot model allows downstream formulas and pivot-specific functions
    /// (e.g. GETPIVOTDATA) to match Excel semantics.
    Date(NaiveDate),
    Text(String),
    Bool(bool),
}

impl PivotValue {
    /// Returns a canonical bit pattern for numeric pivot items.
    ///
    /// This matches Excel behavior where `0.0` and `-0.0` are treated as the same
    /// item, and all NaN payloads are treated as the same item.
    pub fn canonical_number_bits(n: f64) -> u64 {
        if n == 0.0 {
            // Treat -0.0 and 0.0 as identical (Excel renders them identically).
            return 0.0_f64.to_bits();
        }
        if n.is_nan() {
            // Canonicalize all NaNs so we don't emit multiple distinct "NaN" keys.
            return f64::NAN.to_bits();
        }
        n.to_bits()
    }

    /// Converts this value into a typed key part suitable for grouping and sorting.
    pub fn to_key_part(&self) -> PivotKeyPart {
        match self {
            PivotValue::Blank => PivotKeyPart::Blank,
            PivotValue::Number(n) => PivotKeyPart::Number(Self::canonical_number_bits(*n)),
            PivotValue::Date(d) => PivotKeyPart::Date(*d),
            PivotValue::Text(s) => PivotKeyPart::Text(s.clone()),
            PivotValue::Bool(b) => PivotKeyPart::Bool(*b),
        }
    }

    /// Returns a display-oriented string for this pivot value (not a stable serialization).
    pub fn display_string(&self) -> String {
        match self {
            PivotValue::Blank => String::new(),
            PivotValue::Number(n) => {
                formula_format::format_value(FmtValue::Number(*n), None, &FormatOptions::default())
                    .text
            }
            PivotValue::Date(d) => d.to_string(),
            PivotValue::Text(s) => s.clone(),
            PivotValue::Bool(b) => {
                if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
        }
    }

    pub fn as_number(&self) -> Option<f64> {
        match self {
            PivotValue::Number(n) => Some(*n),
            _ => None,
        }
    }

    pub fn is_blank(&self) -> bool {
        matches!(self, PivotValue::Blank)
    }
}

impl From<&str> for PivotValue {
    fn from(value: &str) -> Self {
        PivotValue::Text(value.to_string())
    }
}

impl From<String> for PivotValue {
    fn from(value: String) -> Self {
        PivotValue::Text(value)
    }
}

impl From<f64> for PivotValue {
    fn from(value: f64) -> Self {
        PivotValue::Number(value)
    }
}

impl From<i64> for PivotValue {
    fn from(value: i64) -> Self {
        PivotValue::Number(value as f64)
    }
}

impl From<NaiveDate> for PivotValue {
    fn from(value: NaiveDate) -> Self {
        PivotValue::Date(value)
    }
}

impl From<bool> for PivotValue {
    fn from(value: bool) -> Self {
        PivotValue::Bool(value)
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotField {
    pub source_field: PivotFieldRef,
    #[serde(default)]
    pub sort_order: SortOrder,
    #[serde(default)]
    pub manual_sort: Option<Vec<PivotKeyPart>>,
}

impl PivotField {
    pub fn new(source_field: impl Into<PivotFieldRef>) -> Self {
        Self {
            source_field: source_field.into(),
            sort_order: SortOrder::default(),
            manual_sort: None,
        }
    }
}
#[derive(Clone, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum ScalarValue {
    Text(String),
    Number(OrderedFloat<f64>),
    Date(NaiveDate),
    Bool(bool),
    Blank,
}

impl ScalarValue {
    pub fn as_f64(&self) -> Option<f64> {
        match self {
            ScalarValue::Number(v) => Some(v.0),
            _ => None,
        }
    }
}

/// Aggregation modes for value fields in a pivot table.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum AggregationType {
    Sum,
    Count,
    Average,
    Max,
    Min,
    Product,
    CountNumbers,
    StdDev,
    StdDevP,
    Var,
    VarP,
}
impl AggregationType {
    /// Returns whether this aggregation is currently supported for Data Model pivots.
    pub fn is_supported_for_data_model(&self) -> bool {
        matches!(
            self,
            AggregationType::Sum
                | AggregationType::Count
                | AggregationType::Average
                | AggregationType::Max
                | AggregationType::Min
                | AggregationType::CountNumbers
        )
    }
}
/// Excel-style "Show Values As" transformations.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ShowAsType {
    Normal,
    PercentOfGrandTotal,
    PercentOfRowTotal,
    PercentOfColumnTotal,
    PercentOf,
    PercentDifferenceFrom,
    RunningTotal,
    RankAscending,
    RankDescending,
}

/// Configuration for a pivot table value field.
///
/// This struct is part of the canonical pivot model (IPC/serialization friendly).
/// `show_as: None` (or missing in serialized data) is treated as "normal".
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ValueField {
    pub source_field: PivotFieldRef,
    pub name: String,
    pub aggregation: AggregationType,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_as: Option<ShowAsType>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub base_field: Option<PivotFieldRef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub base_item: Option<String>,
}
impl From<&str> for ScalarValue {
    fn from(value: &str) -> Self {
        ScalarValue::Text(value.to_string())
    }
}

impl From<String> for ScalarValue {
    fn from(value: String) -> Self {
        ScalarValue::Text(value)
    }
}

impl From<f64> for ScalarValue {
    fn from(value: f64) -> Self {
        ScalarValue::Number(OrderedFloat(value))
    }
}

impl From<NaiveDate> for ScalarValue {
    fn from(value: NaiveDate) -> Self {
        ScalarValue::Date(value)
    }
}

#[derive(Clone, Debug)]
pub struct DataTable {
    headers: Vec<String>,
    rows: Vec<Vec<ScalarValue>>,
    header_index: HashMap<String, usize>,
}

impl DataTable {
    pub fn new(headers: Vec<String>, rows: Vec<Vec<ScalarValue>>) -> Result<Self, String> {
        let mut header_index: HashMap<String, usize> = HashMap::new();
        if header_index.try_reserve(headers.len()).is_err() {
            debug_assert!(
                false,
                "allocation failed (pivot data table header index, cols={})",
                headers.len()
            );
            return Err(String::new());
        }
        for (idx, header) in headers.iter().enumerate() {
            if header_index.insert(header.clone(), idx).is_some() {
                return Err(format!("duplicate column header: {header}"));
            }
        }

        for (row_idx, row) in rows.iter().enumerate() {
            if row.len() != headers.len() {
                return Err(format!(
                    "row {row_idx} has {} cells but table has {} columns",
                    row.len(),
                    headers.len()
                ));
            }
        }

        Ok(Self {
            headers,
            rows,
            header_index,
        })
    }

    pub fn headers(&self) -> &[String] {
        &self.headers
    }

    pub fn rows(&self) -> &[Vec<ScalarValue>] {
        &self.rows
    }

    pub fn column_index(&self, header: &str) -> Option<usize> {
        self.header_index.get(header).copied()
    }

    pub fn cell(&self, row: usize, column: &str) -> Option<&ScalarValue> {
        let col_idx = self.column_index(column)?;
        self.rows.get(row)?.get(col_idx)
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct PivotOutput {
    pub headers: Vec<String>,
    pub rows: Vec<Vec<ScalarValue>>,
}

#[derive(Clone, Debug)]
pub struct PivotTable {
    pub id: PivotTableId,
    pub name: String,
    source: DataTable,
    row_fields: Vec<String>,
    value_field: String,
    output: PivotOutput,
}

impl PivotTable {
    pub fn new(
        name: impl Into<String>,
        source: DataTable,
        row_fields: Vec<String>,
        value_field: impl Into<String>,
    ) -> Result<Self, String> {
        let name = name.into();
        let value_field = value_field.into();
        if row_fields.is_empty() {
            return Err("pivot table must have at least one row field".to_string());
        }
        for field in row_fields.iter().chain(std::iter::once(&value_field)) {
            if source.column_index(field).is_none() {
                return Err(format!(
                    "pivot field {field} does not exist in source table"
                ));
            }
        }

        Ok(Self {
            id: crate::new_uuid(),
            name,
            source,
            row_fields,
            value_field,
            output: PivotOutput {
                headers: Vec::new(),
                rows: Vec::new(),
            },
        })
    }

    pub fn refresh(&mut self, filters: &[slicers::RowFilter]) -> Result<(), String> {
        let mut row_indices: Vec<usize> = Vec::new();
        if row_indices.try_reserve_exact(self.row_fields.len()).is_err() {
            debug_assert!(
                false,
                "allocation failed (pivot row field indices, fields={})",
                self.row_fields.len()
            );
            return Err(String::new());
        }
        for field in &self.row_fields {
            row_indices.push(
                self.source
                    .column_index(field)
                    .ok_or_else(|| format!("missing row field {field}"))?,
            );
        }
        let value_idx = self
            .source
            .column_index(&self.value_field)
            .ok_or_else(|| format!("missing value field {}", self.value_field))?;

        let mut aggregates: HashMap<Vec<ScalarValue>, f64> = HashMap::new();
        'rows: for row in self.source.rows() {
            for filter in filters {
                if !filter.matches(&self.source, row)? {
                    continue 'rows;
                }
            }

            let mut key: Vec<ScalarValue> = Vec::new();
            if key.try_reserve_exact(row_indices.len()).is_err() {
                debug_assert!(
                    false,
                    "allocation failed (pivot row key, fields={})",
                    row_indices.len()
                );
                return Err(String::new());
            }
            for idx in &row_indices {
                key.push(row[*idx].clone());
            }
            let value = row[value_idx]
                .as_f64()
                .ok_or_else(|| format!("pivot value field {} must be numeric", self.value_field))?;
            *aggregates.entry(key).or_insert(0.0) += value;
        }

        let mut rows: Vec<(Vec<ScalarValue>, f64)> = Vec::new();
        if rows.try_reserve_exact(aggregates.len()).is_err() {
            debug_assert!(
                false,
                "allocation failed (pivot aggregate rows, rows={})",
                aggregates.len()
            );
            return Err(String::new());
        }
        for (key, value) in aggregates {
            rows.push((key, value));
        }
        rows.sort_by(|(left_key, _), (right_key, _)| left_key.cmp(right_key));

        let header_count = self.row_fields.len().saturating_add(1);
        let mut headers: Vec<String> = Vec::new();
        if headers.try_reserve_exact(header_count).is_err() {
            debug_assert!(
                false,
                "allocation failed (pivot output headers, fields={})",
                header_count
            );
            return Err(String::new());
        }
        for field in &self.row_fields {
            headers.push(field.clone());
        }
        headers.push(self.value_field.clone());

        let mut output_rows: Vec<Vec<ScalarValue>> = Vec::new();
        if output_rows.try_reserve_exact(rows.len()).is_err() {
            debug_assert!(
                false,
                "allocation failed (pivot output rows, rows={})",
                rows.len()
            );
            return Err(String::new());
        }
        for (key, value) in rows {
            let mut row = key;
            row.push(ScalarValue::Number(OrderedFloat(value)));
            output_rows.push(row);
        }

        self.output = PivotOutput {
            headers,
            rows: output_rows,
        };
        Ok(())
    }

    pub fn output(&self) -> &PivotOutput {
        &self.output
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct PivotChartData {
    pub categories: Vec<Vec<ScalarValue>>,
    pub values: Vec<f64>,
}

#[derive(Clone, Debug)]
pub struct PivotChart {
    pub id: PivotChartId,
    pub name: String,
    pub pivot_table_id: PivotTableId,
    data: PivotChartData,
}

impl PivotChart {
    pub fn new(name: impl Into<String>, pivot_table_id: PivotTableId) -> Self {
        Self {
            id: crate::new_uuid(),
            name: name.into(),
            pivot_table_id,
            data: PivotChartData {
                categories: Vec::new(),
                values: Vec::new(),
            },
        }
    }

    pub fn refresh_from_pivot(&mut self, pivot: &PivotTable) -> Result<(), String> {
        if pivot.id != self.pivot_table_id {
            return Err("pivot chart is bound to a different pivot table".to_string());
        }

        let output = pivot.output();
        if output.headers.is_empty() {
            self.data = PivotChartData {
                categories: Vec::new(),
                values: Vec::new(),
            };
            return Ok(());
        }

        let category_width = output.headers.len().saturating_sub(1);
        let mut categories: Vec<Vec<ScalarValue>> = Vec::new();
        let mut values: Vec<f64> = Vec::new();
        if categories.try_reserve_exact(output.rows.len()).is_err()
            || values.try_reserve_exact(output.rows.len()).is_err()
        {
            debug_assert!(
                false,
                "allocation failed (pivot chart data, rows={})",
                output.rows.len()
            );
            return Err(String::new());
        }
        for row in &output.rows {
            if row.len() != output.headers.len() {
                return Err("pivot output row width mismatch".to_string());
            }
            categories.push(row[..category_width].to_vec());
            let value = row[category_width]
                .as_f64()
                .ok_or_else(|| "pivot chart value must be numeric".to_string())?;
            values.push(value);
        }

        self.data = PivotChartData { categories, values };
        Ok(())
    }

    pub fn data(&self) -> &PivotChartData {
        &self.data
    }
}

#[derive(Default)]
pub struct PivotManager {
    pivots: HashMap<PivotTableId, PivotTable>,
    slicers: HashMap<slicers::SlicerId, slicers::Slicer>,
    timelines: HashMap<slicers::TimelineId, slicers::Timeline>,
    charts: HashMap<PivotChartId, PivotChart>,
}

impl PivotManager {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn add_pivot_table(&mut self, pivot: PivotTable) -> PivotTableId {
        let id = pivot.id;
        self.pivots.insert(id, pivot);
        id
    }

    pub fn create_pivot_table(
        &mut self,
        name: impl Into<String>,
        source: DataTable,
        row_fields: Vec<String>,
        value_field: impl Into<String>,
    ) -> Result<PivotTableId, String> {
        let mut pivot = PivotTable::new(name, source, row_fields, value_field)?;
        pivot.refresh(&[])?;
        let id = pivot.id;
        self.pivots.insert(id, pivot);
        Ok(id)
    }

    pub fn create_pivot_chart(
        &mut self,
        pivot_table_id: PivotTableId,
        name: impl Into<String>,
    ) -> Result<PivotChartId, String> {
        let pivot = self
            .pivots
            .get(&pivot_table_id)
            .ok_or_else(|| "unknown pivot table".to_string())?;
        let mut chart = PivotChart::new(name, pivot_table_id);
        chart.refresh_from_pivot(pivot)?;
        let id = chart.id;
        self.charts.insert(id, chart);
        Ok(id)
    }

    pub fn add_slicer_to_pivot(
        &mut self,
        pivot_table_id: PivotTableId,
        slicer_name: impl Into<String>,
        field: impl Into<String>,
    ) -> Result<slicers::SlicerId, String> {
        if !self.pivots.contains_key(&pivot_table_id) {
            return Err("unknown pivot table".to_string());
        }
        let mut slicer = slicers::Slicer::new(slicer_name, field);
        slicer.connect(pivot_table_id);
        let id = slicer.id;
        self.slicers.insert(id, slicer);
        self.refresh_pivot_and_dependents(pivot_table_id)?;
        Ok(id)
    }

    pub fn add_timeline_to_pivot(
        &mut self,
        pivot_table_id: PivotTableId,
        timeline_name: impl Into<String>,
        field: impl Into<String>,
    ) -> Result<slicers::TimelineId, String> {
        if !self.pivots.contains_key(&pivot_table_id) {
            return Err("unknown pivot table".to_string());
        }
        let mut timeline = slicers::Timeline::new(timeline_name, field);
        timeline.connect(pivot_table_id);
        let id = timeline.id;
        self.timelines.insert(id, timeline);
        self.refresh_pivot_and_dependents(pivot_table_id)?;
        Ok(id)
    }

    pub fn set_slicer_selection(
        &mut self,
        slicer_id: slicers::SlicerId,
        selection: slicers::SlicerSelection,
    ) -> Result<(), String> {
        let pivot_ids = {
            let slicer = self
                .slicers
                .get_mut(&slicer_id)
                .ok_or_else(|| "unknown slicer".to_string())?;
            slicer.selection = selection;
            let mut out: Vec<PivotTableId> = Vec::new();
            if out.try_reserve_exact(slicer.connected_pivots.len()).is_err() {
                debug_assert!(
                    false,
                    "allocation failed (slicer pivot list, pivots={})",
                    slicer.connected_pivots.len()
                );
                return Err(String::new());
            }
            out.extend(slicer.connected_pivots.iter().copied());
            out
        };

        for pivot_id in pivot_ids {
            self.refresh_pivot_and_dependents(pivot_id)?;
        }

        Ok(())
    }

    pub fn set_timeline_selection(
        &mut self,
        timeline_id: slicers::TimelineId,
        selection: slicers::TimelineSelection,
    ) -> Result<(), String> {
        let pivot_ids = {
            let timeline = self
                .timelines
                .get_mut(&timeline_id)
                .ok_or_else(|| "unknown timeline".to_string())?;
            timeline.selection = selection;
            let mut out: Vec<PivotTableId> = Vec::new();
            if out.try_reserve_exact(timeline.connected_pivots.len()).is_err() {
                debug_assert!(
                    false,
                    "allocation failed (timeline pivot list, pivots={})",
                    timeline.connected_pivots.len()
                );
                return Err(String::new());
            }
            out.extend(timeline.connected_pivots.iter().copied());
            out
        };

        for pivot_id in pivot_ids {
            self.refresh_pivot_and_dependents(pivot_id)?;
        }

        Ok(())
    }

    pub fn pivot_output(&self, pivot_table_id: PivotTableId) -> Option<&PivotOutput> {
        self.pivots.get(&pivot_table_id).map(|pivot| pivot.output())
    }

    pub fn chart_data(&self, chart_id: PivotChartId) -> Option<&PivotChartData> {
        self.charts.get(&chart_id).map(|chart| chart.data())
    }

    fn refresh_pivot_and_dependents(&mut self, pivot_table_id: PivotTableId) -> Result<(), String> {
        let filters = self.filters_for_pivot(pivot_table_id);
        let pivot = self
            .pivots
            .get_mut(&pivot_table_id)
            .ok_or_else(|| "unknown pivot table".to_string())?;
        pivot.refresh(&filters)?;

        let pivot_snapshot = pivot.clone();
        for chart in self.charts.values_mut() {
            if chart.pivot_table_id == pivot_table_id {
                chart.refresh_from_pivot(&pivot_snapshot)?;
            }
        }

        Ok(())
    }

    fn filters_for_pivot(&self, pivot_table_id: PivotTableId) -> Vec<slicers::RowFilter> {
        let mut filters = Vec::new();
        for slicer in self.slicers.values() {
            if slicer.connected_pivots.contains(&pivot_table_id) {
                filters.push(slicer.as_filter());
            }
        }
        for timeline in self.timelines.values() {
            if timeline.connected_pivots.contains(&pivot_table_id) {
                filters.push(timeline.as_filter());
            }
        }
        filters
    }
}

#[cfg(test)]
mod tests;
