use serde::{Deserialize, Serialize};
use uuid::Uuid;

use crate::table::TableIdentifier;
use crate::{CellRef, Range, WorksheetId};

use std::collections::HashSet;

use super::{CalculatedField, CalculatedItem, PivotField, PivotKeyPart, PivotTableId, ValueField};

pub type PivotCacheId = Uuid;

/// Canonical (IPC/persistence-friendly) pivot table definition stored in a [`crate::Workbook`].
///
/// This is intentionally distinct from the legacy in-memory [`super::PivotTable`] runtime type.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotTableModel {
    pub id: PivotTableId,
    pub name: String,
    pub source: PivotSource,
    pub destination: PivotDestination,
    pub config: PivotConfig,
    /// Placeholder for future pivot-cache storage.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cache_id: Option<PivotCacheId>,
}

/// Source data for a pivot table.
///
/// Shapes are aligned with `docs/07-power-features.md` (Range/Table/DataModel).
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum PivotSource {
    Range { sheet_id: WorksheetId, range: Range },
    Table { table: TableIdentifier },
    DataModel { table: String },
}

/// Where pivot output is rendered.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum PivotDestination {
    /// Anchor the pivot at a specific top-left cell.
    Cell {
        sheet_id: WorksheetId,
        cell: CellRef,
    },
    /// Anchor the pivot to a range (typically the existing pivot output range).
    Range { sheet_id: WorksheetId, range: Range },
}

/// Excel-style layout mode for pivot output.
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

/// Where subtotals appear for a grouped field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SubtotalPosition {
    Top,
    Bottom,
    None,
}

impl Default for SubtotalPosition {
    fn default() -> Self {
        Self::None
    }
}

/// Controls whether grand totals are produced for rows and/or columns.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GrandTotals {
    pub rows: bool,
    pub columns: bool,
}

impl Default for GrandTotals {
    fn default() -> Self {
        // Match Excel defaults.
        Self {
            rows: true,
            columns: true,
        }
    }
}

/// Configuration for a pivot table filter field.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FilterField {
    pub source_field: String,
    /// Allowed values. `None` means allow all.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub allowed: Option<HashSet<PivotKeyPart>>,
}

/// Canonical pivot table configuration (field layout + display options).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotConfig {
    pub row_fields: Vec<PivotField>,
    pub column_fields: Vec<PivotField>,
    pub value_fields: Vec<ValueField>,
    pub filter_fields: Vec<FilterField>,
    #[serde(default)]
    pub calculated_fields: Vec<CalculatedField>,
    #[serde(default)]
    pub calculated_items: Vec<CalculatedItem>,
    pub layout: Layout,
    pub subtotals: SubtotalPosition,
    pub grand_totals: GrandTotals,
}

impl Default for PivotConfig {
    fn default() -> Self {
        Self {
            row_fields: Vec::new(),
            column_fields: Vec::new(),
            value_fields: Vec::new(),
            filter_fields: Vec::new(),
            calculated_fields: Vec::new(),
            calculated_items: Vec::new(),
            layout: Layout::default(),
            subtotals: SubtotalPosition::default(),
            grand_totals: GrandTotals::default(),
        }
    }
}
