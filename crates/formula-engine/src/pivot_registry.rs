use crate::eval::CellAddr;
use crate::pivot::PivotTable;
use std::collections::HashMap;
use std::sync::Arc;
use thiserror::Error;

#[derive(Debug, Error)]
pub enum PivotRegistryError {
    #[error("pivot field {0:?} is missing from the pivot cache")]
    MissingField(String),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotAxis {
    Row,
    Column,
    Filter,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PivotFieldPosition {
    pub axis: PivotAxis,
    pub index: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PivotDestination {
    pub start: CellAddr,
    pub end: CellAddr,
}

impl PivotDestination {
    pub fn contains(&self, addr: CellAddr) -> bool {
        addr.row >= self.start.row
            && addr.row <= self.end.row
            && addr.col >= self.start.col
            && addr.col <= self.end.col
    }
}

/// Metadata for a pivot table rendered into a worksheet destination range.
#[derive(Debug, Clone)]
pub struct PivotRegistryEntry {
    pub pivot_id: String,
    pub sheet_id: usize,
    pub destination: PivotDestination,
    pub pivot: Arc<PivotTable>,
    /// Case-folded field display name -> axis position (row/column/filter + index).
    pub field_positions: HashMap<String, PivotFieldPosition>,
    /// Case-folded field display name -> source-cache column index.
    pub field_indices: HashMap<String, usize>,
    /// Case-folded value-field caption -> value-field index in `pivot.config.value_fields`.
    pub value_field_indices: HashMap<String, usize>,
    /// Case-folded cache field name -> canonical cache field name (used to access `cache.unique_values`).
    pub cache_field_names: HashMap<String, String>,
    /// Per value-field index, the source-cache column index for that value field.
    pub value_field_source_indices: Vec<usize>,
}

impl PivotRegistryEntry {
    pub fn new(
        sheet_id: usize,
        destination: PivotDestination,
        pivot: PivotTable,
    ) -> Result<Self, PivotRegistryError> {
        let pivot_id = pivot.id.clone();
        let pivot = Arc::new(pivot);

        let mut cache_field_indices: HashMap<String, usize> = HashMap::new();
        let mut cache_field_names: HashMap<String, String> = HashMap::new();
        for f in &pivot.cache.fields {
            cache_field_indices.insert(crate::value::casefold(&f.name), f.index);
            cache_field_names.insert(crate::value::casefold(&f.name), f.name.clone());
        }

        let mut field_positions: HashMap<String, PivotFieldPosition> = HashMap::new();
        let mut field_indices: HashMap<String, usize> = HashMap::new();

        for (idx, f) in pivot.config.row_fields.iter().enumerate() {
            let name = f
                .source_field
                .as_cache_field_name()
                .ok_or_else(|| PivotRegistryError::MissingField(f.source_field.to_string()))?;
            let key = crate::value::casefold(name);
            let cache_idx = cache_field_indices
                .get(&key)
                .copied()
                .ok_or_else(|| PivotRegistryError::MissingField(f.source_field.to_string()))?;
            field_positions.insert(
                key.clone(),
                PivotFieldPosition {
                    axis: PivotAxis::Row,
                    index: idx,
                },
            );
            field_indices.insert(key, cache_idx);
        }

        for (idx, f) in pivot.config.column_fields.iter().enumerate() {
            let name = f
                .source_field
                .as_cache_field_name()
                .ok_or_else(|| PivotRegistryError::MissingField(f.source_field.to_string()))?;
            let key = crate::value::casefold(name);
            let cache_idx = cache_field_indices
                .get(&key)
                .copied()
                .ok_or_else(|| PivotRegistryError::MissingField(f.source_field.to_string()))?;
            field_positions.insert(
                key.clone(),
                PivotFieldPosition {
                    axis: PivotAxis::Column,
                    index: idx,
                },
            );
            field_indices.insert(key, cache_idx);
        }

        for (idx, f) in pivot.config.filter_fields.iter().enumerate() {
            let name = f
                .source_field
                .as_cache_field_name()
                .ok_or_else(|| PivotRegistryError::MissingField(f.source_field.to_string()))?;
            let key = crate::value::casefold(name);
            let cache_idx = cache_field_indices
                .get(&key)
                .copied()
                .ok_or_else(|| PivotRegistryError::MissingField(f.source_field.to_string()))?;
            field_positions.insert(
                key.clone(),
                PivotFieldPosition {
                    axis: PivotAxis::Filter,
                    index: idx,
                },
            );
            field_indices.insert(key, cache_idx);
        }

        let mut value_field_indices: HashMap<String, usize> = HashMap::new();
        let mut value_field_source_indices: Vec<usize> =
            Vec::with_capacity(pivot.config.value_fields.len());
        for (idx, vf) in pivot.config.value_fields.iter().enumerate() {
            value_field_indices.insert(crate::value::casefold(&vf.name), idx);
            let name = vf
                .source_field
                .as_cache_field_name()
                .ok_or_else(|| PivotRegistryError::MissingField(vf.source_field.to_string()))?;
            let key = crate::value::casefold(name);
            let cache_idx = cache_field_indices
                .get(&key)
                .copied()
                .ok_or_else(|| PivotRegistryError::MissingField(vf.source_field.to_string()))?;
            value_field_source_indices.push(cache_idx);
        }

        Ok(Self {
            pivot_id,
            sheet_id,
            destination,
            pivot,
            field_positions,
            field_indices,
            value_field_indices,
            cache_field_names,
            value_field_source_indices,
        })
    }
}

/// Registry of pivot tables, keyed by destination worksheet ranges.
#[derive(Debug, Default, Clone)]
pub struct PivotRegistry {
    entries: Vec<PivotRegistryEntry>,
}

impl PivotRegistry {
    pub fn entries(&self) -> &[PivotRegistryEntry] {
        &self.entries
    }

    pub fn register(&mut self, entry: PivotRegistryEntry) {
        self.entries.retain(|e| {
            !(e.sheet_id == entry.sheet_id
                && e.destination.start == entry.destination.start
                && e.destination.end == entry.destination.end)
        });
        self.entries.push(entry);
    }

    pub fn find_by_cell(&self, sheet_id: usize, addr: CellAddr) -> Option<&PivotRegistryEntry> {
        // Prefer the most recently registered pivot if ranges overlap.
        self.entries
            .iter()
            .rev()
            .find(|e| e.sheet_id == sheet_id && e.destination.contains(addr))
    }

    pub fn clear(&mut self) {
        self.entries.clear();
    }
}
