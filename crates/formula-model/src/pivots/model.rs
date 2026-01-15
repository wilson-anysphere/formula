use serde::{Deserialize, Serialize};
use uuid::Uuid;

use crate::table::TableIdentifier;
use crate::value::text_eq_case_insensitive;
use crate::{CellRef, DefinedNameId, Range, WorksheetId};

use super::{PivotConfig, PivotTableId};

pub type PivotCacheId = Uuid;
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
                if crate::sheet_name::sheet_name_eq_case_insensitive(sheet_name, old_name) {
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
                if text_eq_case_insensitive(name, old_name) {
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
                if crate::sheet_name::sheet_name_eq_case_insensitive(sheet_name, old_name) {
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
