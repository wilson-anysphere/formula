use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use uuid::Uuid;

use crate::table::TableIdentifier;
use crate::{CellRef, DefinedNameId, Range, WorksheetId};

use super::{
    CalculatedField, CalculatedItem, PivotField, PivotKeyPart, PivotTableId, ValueField,
};

pub type PivotCacheId = Uuid;

/// Pivot layout style (Excel-like).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum Layout {
    Compact,
    Outline,
    Tabular,
}

impl Default for Layout {
    fn default() -> Self {
        Layout::Tabular
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SubtotalPosition {
    Top,
    Bottom,
    None,
}

impl Default for SubtotalPosition {
    fn default() -> Self {
        SubtotalPosition::None
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
    #[serde(default)]
    pub filter_fields: Vec<FilterField>,
    #[serde(default)]
    pub calculated_fields: Vec<CalculatedField>,
    #[serde(default)]
    pub calculated_items: Vec<CalculatedItem>,
    #[serde(default)]
    pub layout: Layout,
    #[serde(default)]
    pub subtotals: SubtotalPosition,
    #[serde(default)]
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

/// Identifier for a workbook defined name (named range) when used as a pivot source.
///
/// We prefer stable ids, but allow string references for backward-compat / imported metadata
/// that has not been resolved.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(untagged)]
pub enum DefinedNameIdentifier {
    Name(String),
    Id(DefinedNameId),
}

impl From<DefinedNameId> for DefinedNameIdentifier {
    fn from(value: DefinedNameId) -> Self {
        Self::Id(value)
    }
}

impl From<String> for DefinedNameIdentifier {
    fn from(value: String) -> Self {
        Self::Name(value)
    }
}

impl From<&str> for DefinedNameIdentifier {
    fn from(value: &str) -> Self {
        Self::Name(value.to_string())
    }
}

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
    Range {
        sheet_id: WorksheetId,
        range: Range,
    },
    /// Legacy / unresolved range reference keyed by sheet name.
    ///
    /// Prefer [`PivotSource::Range`] when possible.
    RangeName {
        sheet_name: String,
        range: Range,
    },
    /// Named range / defined name.
    NamedRange {
        name: DefinedNameIdentifier,
    },
    Table {
        table: TableIdentifier,
    },
    DataModel {
        table: String,
    },
}

impl PivotSource {
    /// Rewrite any string-based sheet reference from `old_name` to `new_name`.
    pub fn rewrite_sheet_name(&mut self, old_name: &str, new_name: &str) -> bool {
        match self {
            PivotSource::RangeName { sheet_name, .. } => {
                if crate::formula_rewrite::sheet_name_eq_case_insensitive(sheet_name, old_name) {
                    *sheet_name = new_name.to_string();
                    true
                } else {
                    false
                }
            }
            _ => false,
        }
    }

    /// Rewrite any string-based table reference from `old_name` to `new_name`.
    pub fn rewrite_table_name(&mut self, old_name: &str, new_name: &str) -> bool {
        match self {
            PivotSource::Table {
                table: TableIdentifier::Name(name),
            } => {
                if name.eq_ignore_ascii_case(old_name) {
                    *name = new_name.to_string();
                    true
                } else {
                    false
                }
            }
            _ => false,
        }
    }

    /// Rewrite any string-based defined-name reference from `old_name` to `new_name`.
    pub fn rewrite_defined_name(&mut self, old_name: &str, new_name: &str) -> bool {
        match self {
            PivotSource::NamedRange {
                name: DefinedNameIdentifier::Name(name),
            } => {
                if name.eq_ignore_ascii_case(old_name) {
                    *name = new_name.to_string();
                    true
                } else {
                    false
                }
            }
            _ => false,
        }
    }
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
    /// Legacy / unresolved destination keyed by sheet name.
    ///
    /// Prefer [`PivotDestination::Cell`] when possible.
    CellName { sheet_name: String, cell: CellRef },
    /// Anchor the pivot to a range (typically the existing pivot output range).
    Range { sheet_id: WorksheetId, range: Range },
    /// Legacy / unresolved destination keyed by sheet name.
    ///
    /// Prefer [`PivotDestination::Range`] when possible.
    RangeName { sheet_name: String, range: Range },
}

impl PivotDestination {
    /// Rewrite any string-based sheet reference from `old_name` to `new_name`.
    pub fn rewrite_sheet_name(&mut self, old_name: &str, new_name: &str) -> bool {
        match self {
            PivotDestination::CellName { sheet_name, .. }
            | PivotDestination::RangeName { sheet_name, .. } => {
                if crate::formula_rewrite::sheet_name_eq_case_insensitive(sheet_name, old_name) {
                    *sheet_name = new_name.to_string();
                    true
                } else {
                    false
                }
            }
            _ => false,
        }
    }
}
