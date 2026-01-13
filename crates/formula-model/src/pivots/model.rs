use serde::{Deserialize, Serialize};
use uuid::Uuid;

use crate::table::TableIdentifier;
use crate::{CellRef, Range, WorksheetId};

use super::{PivotConfig, PivotTableId};

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
