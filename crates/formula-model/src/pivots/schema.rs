use serde::{Deserialize, Serialize};
use std::collections::HashSet;

use super::{PivotField, PivotKeyPart, ValueField};

/// An Excel-style PivotTable *calculated field*.
///
/// In Excel, a calculated field is a named formula that behaves like an extra source column:
/// it is evaluated for each record in the pivot cache and can then be used as a field in the
/// pivot configuration (most commonly in the "Values" area).
///
/// The Formula pivot engine persists the raw formula text and treats the calculated field as
/// an additional cache field named [`CalculatedField::name`] when building or refreshing the
/// pivot cache.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CalculatedField {
    pub name: String,
    pub formula: String,
}

/// An Excel-style PivotTable *calculated item*.
///
/// In Excel, a calculated item creates a synthetic member (an "item") inside a specific pivot
/// field. The item is defined by a name and a formula that typically references other items
/// within the same field (for example, creating `Q1` from `Jan + Feb + Mar` inside a `Month`
/// field).
///
/// The Formula pivot engine interprets a calculated item as a post-aggregation transform:
/// after the pivot is grouped by [`CalculatedItem::field`], the engine evaluates
/// [`CalculatedItem::formula`] against the existing items for that field and inserts a new item
/// named [`CalculatedItem::name`] into the pivot results.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CalculatedItem {
    pub field: String,
    pub name: String,
    pub formula: String,
}

/// Filter configuration for a pivot field.
///
/// When `allowed` is `None`, the field is unfiltered (all values allowed). When
/// set, it contains the allowed values (represented using the canonical
/// [`PivotKeyPart`] typed key).
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FilterField {
    pub source_field: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub allowed: Option<HashSet<PivotKeyPart>>,
}

/// PivotTable report layout mode (Excel: Compact/Outline/Tabular).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum Layout {
    Compact,
    Outline,
    Tabular,
}

impl Default for Layout {
    fn default() -> Self {
        Self::Tabular
    }
}

/// Where subtotals should be rendered for row field groupings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SubtotalPosition {
    None,
    Top,
    Bottom,
}

impl Default for SubtotalPosition {
    fn default() -> Self {
        Self::None
    }
}

/// Whether to render grand totals for rows and/or columns.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GrandTotals {
    pub rows: bool,
    pub columns: bool,
}

impl Default for GrandTotals {
    fn default() -> Self {
        Self {
            rows: true,
            columns: true,
        }
    }
}

/// Canonical pivot configuration stored in a [`super::PivotTableModel`] and used
/// by the pivot engine.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct PivotConfig {
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub row_fields: Vec<PivotField>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub column_fields: Vec<PivotField>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub value_fields: Vec<ValueField>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub filter_fields: Vec<FilterField>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub calculated_fields: Vec<CalculatedField>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub calculated_items: Vec<CalculatedItem>,
    #[serde(default)]
    pub layout: Layout,
    #[serde(default)]
    pub subtotals: SubtotalPosition,
    #[serde(default)]
    pub grand_totals: GrandTotals,
}
