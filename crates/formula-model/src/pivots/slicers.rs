use chrono::NaiveDate;
use std::collections::{HashSet, VecDeque};
use uuid::Uuid;

use super::{DataTable, PivotTableId, ScalarValue};

pub type SlicerId = Uuid;
pub type TimelineId = Uuid;

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum SlicerSelection {
    #[default]
    All,
    Items(HashSet<ScalarValue>),
}

impl SlicerSelection {
    pub fn is_all(&self) -> bool {
        matches!(self, SlicerSelection::All)
    }

    pub fn matches(&self, value: &ScalarValue) -> bool {
        match self {
            SlicerSelection::All => true,
            SlicerSelection::Items(items) => items.contains(value),
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TimelineSelection {
    pub start: Option<NaiveDate>,
    pub end: Option<NaiveDate>,
}

impl TimelineSelection {
    pub fn matches(&self, value: &ScalarValue) -> bool {
        let ScalarValue::Date(date) = value else {
            return false;
        };

        if let Some(start) = self.start {
            if *date < start {
                return false;
            }
        }
        if let Some(end) = self.end {
            if *date > end {
                return false;
            }
        }
        true
    }
}

#[derive(Clone, Debug)]
pub struct Slicer {
    pub id: SlicerId,
    pub name: String,
    pub field: String,
    pub selection: SlicerSelection,
    pub connected_pivots: HashSet<PivotTableId>,
}

impl Slicer {
    pub fn new(name: impl Into<String>, field: impl Into<String>) -> Self {
        Self {
            id: crate::new_uuid(),
            name: name.into(),
            field: field.into(),
            selection: SlicerSelection::All,
            connected_pivots: HashSet::new(),
        }
    }

    pub fn connect(&mut self, pivot_table_id: PivotTableId) {
        self.connected_pivots.insert(pivot_table_id);
    }

    pub fn disconnect(&mut self, pivot_table_id: PivotTableId) {
        self.connected_pivots.remove(&pivot_table_id);
    }

    pub fn as_filter(&self) -> RowFilter {
        RowFilter::Slicer {
            field: self.field.clone(),
            selection: self.selection.clone(),
        }
    }
}

#[derive(Clone, Debug)]
pub struct Timeline {
    pub id: TimelineId,
    pub name: String,
    pub field: String,
    pub selection: TimelineSelection,
    pub connected_pivots: HashSet<PivotTableId>,
}

impl Timeline {
    pub fn new(name: impl Into<String>, field: impl Into<String>) -> Self {
        Self {
            id: crate::new_uuid(),
            name: name.into(),
            field: field.into(),
            selection: TimelineSelection::default(),
            connected_pivots: HashSet::new(),
        }
    }

    pub fn connect(&mut self, pivot_table_id: PivotTableId) {
        self.connected_pivots.insert(pivot_table_id);
    }

    pub fn disconnect(&mut self, pivot_table_id: PivotTableId) {
        self.connected_pivots.remove(&pivot_table_id);
    }

    pub fn as_filter(&self) -> RowFilter {
        RowFilter::Timeline {
            field: self.field.clone(),
            selection: self.selection.clone(),
        }
    }
}

#[derive(Clone, Debug)]
pub enum RowFilter {
    Slicer {
        field: String,
        selection: SlicerSelection,
    },
    Timeline {
        field: String,
        selection: TimelineSelection,
    },
}

impl RowFilter {
    pub fn matches(&self, table: &DataTable, row: &[ScalarValue]) -> Result<bool, String> {
        match self {
            RowFilter::Slicer { field, selection } => {
                let idx = table
                    .column_index(field)
                    .ok_or_else(|| format!("filter refers to unknown field {field}"))?;
                Ok(selection.matches(&row[idx]))
            }
            RowFilter::Timeline { field, selection } => {
                let idx = table
                    .column_index(field)
                    .ok_or_else(|| format!("filter refers to unknown field {field}"))?;
                Ok(selection.matches(&row[idx]))
            }
        }
    }
}

/// Computes the list of items that should be rendered in a slicer UI for a given field.
///
/// The model layer deduplicates values and preserves the order of first appearance, which
/// is typically what users expect when creating slicers from raw data.
pub fn distinct_items(table: &DataTable, field: &str) -> Result<Vec<ScalarValue>, String> {
    let idx = table
        .column_index(field)
        .ok_or_else(|| format!("unknown slicer field {field}"))?;

    let mut seen = HashSet::new();
    let mut ordered = VecDeque::new();
    for row in table.rows() {
        let value = row[idx].clone();
        if seen.insert(value.clone()) {
            ordered.push_back(value);
        }
    }

    let mut out: Vec<ScalarValue> = Vec::new();
    if out.try_reserve_exact(ordered.len()).is_err() {
        debug_assert!(
            false,
            "allocation failed (distinct slicer items, count={})",
            ordered.len()
        );
        return Err(String::new());
    }
    out.extend(ordered.into_iter());
    Ok(out)
}
