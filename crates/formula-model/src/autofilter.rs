use crate::Range;
use chrono::{NaiveDate, NaiveDateTime};
use serde::{Deserialize, Serialize};

/// Sort condition within an AutoFilter / SortState payload.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct SortCondition {
    pub range: Range,
    pub descending: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct SortState {
    pub conditions: Vec<SortCondition>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "snake_case")]
pub enum FilterJoin {
    /// Any criterion may match (logical OR).
    #[default]
    Any,
    /// All criteria must match (logical AND).
    All,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum FilterValue {
    Text(String),
    Number(f64),
    Bool(bool),
    DateTime(NaiveDateTime),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum TextMatchKind {
    Contains,
    BeginsWith,
    EndsWith,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TextMatch {
    pub kind: TextMatchKind,
    pub pattern: String,
    #[serde(default)]
    pub case_sensitive: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum NumberComparison {
    GreaterThan(f64),
    GreaterThanOrEqual(f64),
    LessThan(f64),
    LessThanOrEqual(f64),
    Between { min: f64, max: f64 },
    NotEqual(f64),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum DateComparison {
    After(NaiveDateTime),
    Before(NaiveDateTime),
    Between {
        start: NaiveDateTime,
        end: NaiveDateTime,
    },
    OnDate(NaiveDate),
    Today,
    Yesterday,
    Tomorrow,
}

/// Preserve an unsupported custom filter operator/value pair so it can be
/// round-tripped without loss.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct OpaqueCustomFilter {
    pub operator: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
}

/// Preserve an unsupported dynamic filter payload so it can be round-tripped
/// without loss.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct OpaqueDynamicFilter {
    #[serde(rename = "type")]
    pub filter_type: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_value: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum FilterCriterion {
    Equals(FilterValue),
    TextMatch(TextMatch),
    Number(NumberComparison),
    Date(DateComparison),
    Blanks,
    NonBlanks,
    /// An unknown `customFilter operator="..."` payload.
    OpaqueCustom(OpaqueCustomFilter),
    /// An unknown `<dynamicFilter type="..."/>` payload.
    OpaqueDynamic(OpaqueDynamicFilter),
}

/// A filter definition for a column within an AutoFilter range.
///
/// `col_id` is a 0-based offset from the AutoFilter range start column, matching
/// Excel's `filterColumn/@colId`.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FilterColumn {
    /// 0-based column offset within the AutoFilter range.
    pub col_id: u32,
    /// How multiple criteria are combined.
    #[serde(default)]
    pub join: FilterJoin,
    /// Filter criteria for this column.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub criteria: Vec<FilterCriterion>,
    /// Legacy value-list representation used by earlier schema versions.
    ///
    /// This corresponds to the `<filters><filter val="..."/></filters>` form.
    /// New code should generally prefer [`FilterColumn::criteria`].
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub values: Vec<String>,
    /// Opaque XML payload for unsupported criteria within the filterColumn.
    ///
    /// Each entry should be a full XML element (e.g. `<top10 .../>`), suitable
    /// for re-insertion into the `<filterColumn>` body during XLSX round-trip.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub raw_xml: Vec<String>,
}

/// Worksheet-level AutoFilter state.
///
/// This payload corresponds to the worksheet `<autoFilter>` element, and is
/// also reused by table definitions (`table.xml` contains an `<autoFilter>`
/// element with identical semantics).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SheetAutoFilter {
    pub range: Range,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub filter_columns: Vec<FilterColumn>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort_state: Option<SortState>,
    /// Opaque XML payload for unsupported children within `<autoFilter>`.
    ///
    /// Each entry should be a full XML element suitable for re-insertion into
    /// the `<autoFilter>` body.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub raw_xml: Vec<String>,
}
