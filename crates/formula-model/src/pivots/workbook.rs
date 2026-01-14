use serde::{Deserialize, Serialize};

use crate::WorksheetId;

use super::{PivotCacheId, PivotChartId, PivotSource, PivotTableId};

fn is_false(v: &bool) -> bool {
    !*v
}

/// Workbook-level pivot cache metadata.
///
/// The model stores caches separately from pivot tables so multiple pivots can share the same
/// cached record set (Excel-style). The cache payload itself is not modeled yet; `needs_refresh`
/// acts as a lightweight invalidation flag for higher layers.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotCacheModel {
    pub id: PivotCacheId,
    pub source: PivotSource,
    /// When set, the cache should be rebuilt from `source` on the next refresh.
    #[serde(default, skip_serializing_if = "is_false")]
    pub needs_refresh: bool,
}

/// Workbook-level pivot chart definition bound to a pivot table.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotChartModel {
    pub id: PivotChartId,
    pub name: String,
    pub pivot_table_id: PivotTableId,
    /// Optional placement hint for sheet-hosted pivot charts.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_id: Option<WorksheetId>,
}

/// Workbook-level slicer definition.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlicerModel {
    pub id: super::slicers::SlicerId,
    pub name: String,
    pub connected_pivots: Vec<PivotTableId>,
    pub sheet_id: WorksheetId,
}

/// Workbook-level timeline definition.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TimelineModel {
    pub id: super::slicers::TimelineId,
    pub name: String,
    pub connected_pivots: Vec<PivotTableId>,
    pub sheet_id: WorksheetId,
}

