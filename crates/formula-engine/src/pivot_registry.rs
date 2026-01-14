use crate::eval::CellAddr;
use crate::editing::rewrite::{
    rewrite_formula_for_range_map,
    rewrite_formula_for_structural_edit,
    RangeMapEdit,
    StructuralEdit,
};
use crate::pivot::PivotTable;
use formula_model::{CellRef, Range};
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
            let key = crate::value::casefold(f.source_field.canonical_name().as_ref());
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
            let key = crate::value::casefold(f.source_field.canonical_name().as_ref());
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
            let key = crate::value::casefold(f.source_field.canonical_name().as_ref());
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
            let key = crate::value::casefold(vf.source_field.canonical_name().as_ref());
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
        // A pivot can be refreshed with a different destination footprint (e.g. the output grows or
        // shrinks). Ensure we don't keep stale metadata around for the old destination range.
        //
        // `pivot_id` is treated as a stable identifier for a logical pivot across refreshes.
        self.entries.retain(|e| {
            !(e.pivot_id == entry.pivot_id
                || (e.sheet_id == entry.sheet_id
                    && e.destination.start == entry.destination.start
                    && e.destination.end == entry.destination.end))
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

    /// Rewrite any registered pivot destinations affected by a row/column insertion/deletion.
    ///
    /// Pivot registrations are keyed by sheet id + destination range. When structural edits shift
    /// worksheet cells, pivot destinations must shift as well so `GETPIVOTDATA` doesn't resolve
    /// against stale ranges.
    pub fn apply_structural_edit(
        &mut self,
        edit: &StructuralEdit,
        sheet_names: &HashMap<usize, String>,
    ) {
        self.entries.retain_mut(|entry| {
            let Some(ctx_sheet) = sheet_names.get(&entry.sheet_id) else {
                // Sheet no longer exists; drop stale entry.
                return false;
            };
            let Some(new_dest) = rewrite_pivot_destination_for_structural_edit(
                entry.destination,
                ctx_sheet,
                edit,
            ) else {
                return false;
            };
            entry.destination = new_dest;
            true
        });
    }

    /// Rewrite any registered pivot destinations affected by a range move / insert-cells /
    /// delete-cells edit.
    pub fn apply_range_map_edit(
        &mut self,
        edit: &RangeMapEdit,
        sheet_names: &HashMap<usize, String>,
    ) {
        self.entries.retain_mut(|entry| {
            let Some(ctx_sheet) = sheet_names.get(&entry.sheet_id) else {
                return false;
            };
            let Some(new_dest) =
                rewrite_pivot_destination_for_range_map_edit(entry.destination, ctx_sheet, edit)
            else {
                return false;
            };
            entry.destination = new_dest;
            true
        });
    }

    pub fn clear(&mut self) {
        self.entries.clear();
    }
}

fn rewrite_pivot_destination_for_structural_edit(
    destination: PivotDestination,
    ctx_sheet: &str,
    edit: &StructuralEdit,
) -> Option<PivotDestination> {
    let range = Range::new(
        CellRef::new(destination.start.row, destination.start.col),
        CellRef::new(destination.end.row, destination.end.col),
    );
    let formula = format!("={range}");
    let (out, _) = rewrite_formula_for_structural_edit(
        &formula,
        ctx_sheet,
        crate::CellAddr::new(0, 0),
        edit,
    );
    let new_range = parse_a1_range_from_formula(&out)?;
    Some(PivotDestination {
        start: CellAddr {
            row: new_range.start.row,
            col: new_range.start.col,
        },
        end: CellAddr {
            row: new_range.end.row,
            col: new_range.end.col,
        },
    })
}

fn rewrite_pivot_destination_for_range_map_edit(
    destination: PivotDestination,
    ctx_sheet: &str,
    edit: &RangeMapEdit,
) -> Option<PivotDestination> {
    let range = Range::new(
        CellRef::new(destination.start.row, destination.start.col),
        CellRef::new(destination.end.row, destination.end.col),
    );
    let formula = format!("={range}");
    let (out, _) = rewrite_formula_for_range_map(
        &formula,
        ctx_sheet,
        crate::CellAddr::new(0, 0),
        edit,
    );
    let new_range = parse_a1_range_from_formula(&out)?;
    Some(PivotDestination {
        start: CellAddr {
            row: new_range.start.row,
            col: new_range.start.col,
        },
        end: CellAddr {
            row: new_range.end.row,
            col: new_range.end.col,
        },
    })
}

fn parse_a1_range_from_formula(formula: &str) -> Option<Range> {
    // Mirror pivot definition rewrite helpers: accept only a simple A1 range/cell or #REF!.
    let trimmed = formula.trim_start();
    let expr = trimmed.strip_prefix('=').unwrap_or(trimmed).trim();
    if expr.eq_ignore_ascii_case("#REF!") {
        return None;
    }
    // Reject anything that's not a plain A1 range/cell.
    if expr.contains(',') || expr.contains(' ') || expr.contains('(') || expr.contains(')') {
        return None;
    }
    Range::from_a1(expr).ok()
}
