use crate::editing::rewrite::{
    rewrite_formula_for_range_map, rewrite_formula_for_structural_edit, RangeMapEdit,
    StructuralEdit,
};
use crate::eval::CellAddr;
use crate::pivot::PivotTable;
use formula_model::pivots::{parse_dax_column_ref, parse_dax_measure_ref, PivotFieldRef};
use formula_model::{CellRef, Range};
use std::borrow::Cow;
use std::collections::HashMap;
use std::sync::Arc;
use thiserror::Error;

#[derive(Debug, Error)]
pub enum PivotRegistryError {
    #[error("pivot field {0:?} is missing from the pivot cache")]
    MissingField(String),
}

pub(crate) fn normalize_pivot_cache_field_name(name: &str) -> Cow<'_, str> {
    if let Some(measure) = parse_dax_measure_ref(name) {
        return Cow::Owned(
            PivotFieldRef::DataModelMeasure(measure)
                .canonical_name()
                .into_owned(),
        );
    }
    if let Some((table, column)) = parse_dax_column_ref(name) {
        return Cow::Owned(
            PivotFieldRef::DataModelColumn { table, column }
                .canonical_name()
                .into_owned(),
        );
    }
    Cow::Borrowed(name)
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
    /// Case-folded field key (`PivotFieldRef::canonical_name`) -> pivot-cache field name.
    ///
    /// This is used to access cache metadata keyed by the stored field names (e.g. `unique_values`).
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
            // Store lookups for both the raw cache header and a normalized "Data Model" label.
            //
            // This lets `GETPIVOTDATA` resolve fields by either:
            // - the literal cache header (e.g. `"'Dim Product'[Category]"`)
            // - the unquoted DAX-like form (e.g. `"Dim Product[Category]"`)
            //
            // Note: `normalize_pivot_cache_field_name` is best-effort and may cause collisions if a
            // cache contains multiple headers that normalize to the same string. In that case we
            // preserve the first inserted mapping; callers can still match by exact raw header.
            let raw_key = crate::value::casefold(&f.name);
            cache_field_indices
                .entry(raw_key.clone())
                .or_insert(f.index);
            cache_field_names.entry(raw_key).or_insert(f.name.clone());

            let normalized = normalize_pivot_cache_field_name(&f.name);
            let normalized_key = crate::value::casefold(normalized.as_ref());
            cache_field_indices
                .entry(normalized_key.clone())
                .or_insert(f.index);
            cache_field_names
                .entry(normalized_key)
                .or_insert(f.name.clone());
        }

        let resolve_cache_field =
            |field: &PivotFieldRef, key: &str| -> Result<(usize, String), PivotRegistryError> {
                // Fast path: direct lookup by the canonical key.
                if let Some(idx) = cache_field_indices.get(key).copied() {
                    let name = pivot
                        .cache
                        .fields
                        .get(idx)
                        .map(|f| f.name.clone())
                        .unwrap_or_else(|| field.to_string());
                    return Ok((idx, name));
                }

                // Fallback: resolve the field ref against the cache field captions. This handles
                // DAX-quoted column refs (`'Dim Product'[Category]`) and bracketed measure refs
                // (`[Total Sales]`) when the cache stores a different textual form.
                let Some(idx) = pivot.cache.field_index_ref(field) else {
                    return Err(PivotRegistryError::MissingField(field.to_string()));
                };
                let name = pivot
                    .cache
                    .fields
                    .get(idx)
                    .map(|f| f.name.clone())
                    .unwrap_or_else(|| field.to_string());
                Ok((idx, name))
            };

        let mut field_positions: HashMap<String, PivotFieldPosition> = HashMap::new();
        let mut field_indices: HashMap<String, usize> = HashMap::new();

        for (idx, f) in pivot.config.row_fields.iter().enumerate() {
            let canonical = f.source_field.canonical_name();
            let canonical_key = crate::value::casefold(canonical.as_ref());
            let normalized = normalize_pivot_cache_field_name(canonical.as_ref());
            let normalized_key = crate::value::casefold(normalized.as_ref());
            let display_key = crate::value::casefold(&f.source_field.to_string());

            let (cache_idx, cache_name) = resolve_cache_field(&f.source_field, &canonical_key)?;
            let pos = PivotFieldPosition {
                axis: PivotAxis::Row,
                index: idx,
            };

            cache_field_names.insert(canonical_key.clone(), cache_name.clone());
            field_positions.insert(canonical_key.clone(), pos);
            field_indices.insert(canonical_key.clone(), cache_idx);

            if normalized_key != canonical_key {
                cache_field_names.insert(normalized_key.clone(), cache_name.clone());
                field_positions.insert(normalized_key.clone(), pos);
                field_indices.insert(normalized_key.clone(), cache_idx);
            }
            if display_key != canonical_key && display_key != normalized_key {
                cache_field_names.insert(display_key.clone(), cache_name.clone());
                field_positions.insert(display_key.clone(), pos);
                field_indices.insert(display_key, cache_idx);
            }
        }

        for (idx, f) in pivot.config.column_fields.iter().enumerate() {
            let canonical = f.source_field.canonical_name();
            let canonical_key = crate::value::casefold(canonical.as_ref());
            let normalized = normalize_pivot_cache_field_name(canonical.as_ref());
            let normalized_key = crate::value::casefold(normalized.as_ref());
            let display_key = crate::value::casefold(&f.source_field.to_string());

            let (cache_idx, cache_name) = resolve_cache_field(&f.source_field, &canonical_key)?;
            let pos = PivotFieldPosition {
                axis: PivotAxis::Column,
                index: idx,
            };

            cache_field_names.insert(canonical_key.clone(), cache_name.clone());
            field_positions.insert(canonical_key.clone(), pos);
            field_indices.insert(canonical_key.clone(), cache_idx);

            if normalized_key != canonical_key {
                cache_field_names.insert(normalized_key.clone(), cache_name.clone());
                field_positions.insert(normalized_key.clone(), pos);
                field_indices.insert(normalized_key.clone(), cache_idx);
            }
            if display_key != canonical_key && display_key != normalized_key {
                cache_field_names.insert(display_key.clone(), cache_name.clone());
                field_positions.insert(display_key.clone(), pos);
                field_indices.insert(display_key, cache_idx);
            }
        }

        for (idx, f) in pivot.config.filter_fields.iter().enumerate() {
            let canonical = f.source_field.canonical_name();
            let canonical_key = crate::value::casefold(canonical.as_ref());
            let normalized = normalize_pivot_cache_field_name(canonical.as_ref());
            let normalized_key = crate::value::casefold(normalized.as_ref());
            let display_key = crate::value::casefold(&f.source_field.to_string());

            let (cache_idx, cache_name) = resolve_cache_field(&f.source_field, &canonical_key)?;
            let pos = PivotFieldPosition {
                axis: PivotAxis::Filter,
                index: idx,
            };

            cache_field_names.insert(canonical_key.clone(), cache_name.clone());
            field_positions.insert(canonical_key.clone(), pos);
            field_indices.insert(canonical_key.clone(), cache_idx);

            if normalized_key != canonical_key {
                cache_field_names.insert(normalized_key.clone(), cache_name.clone());
                field_positions.insert(normalized_key.clone(), pos);
                field_indices.insert(normalized_key.clone(), cache_idx);
            }
            if display_key != canonical_key && display_key != normalized_key {
                cache_field_names.insert(display_key.clone(), cache_name.clone());
                field_positions.insert(display_key.clone(), pos);
                field_indices.insert(display_key, cache_idx);
            }
        }

        let mut value_field_indices: HashMap<String, usize> = HashMap::new();
        let mut value_field_source_indices: Vec<usize> =
            Vec::with_capacity(pivot.config.value_fields.len());
        for (idx, vf) in pivot.config.value_fields.iter().enumerate() {
            value_field_indices.insert(crate::value::casefold(&vf.name), idx);
            let field_name = vf.source_field.canonical_name();
            let field_name = normalize_pivot_cache_field_name(field_name.as_ref());
            let key = crate::value::casefold(field_name.as_ref());
            let (cache_idx, _cache_name) = resolve_cache_field(&vf.source_field, &key)?;
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

    /// Remove any registry entries that belong to `sheet_id`.
    ///
    /// Sheet ids are stable for the lifetime of a workbook. When a worksheet is deleted, the
    /// engine keeps the deleted id reserved but marks the sheet as missing; registry entries must
    /// be pruned explicitly to avoid leaking stale metadata and accidentally resolving deleted
    /// destinations in `GETPIVOTDATA`.
    pub fn prune_sheet(&mut self, sheet_id: usize) {
        self.entries.retain(|e| e.sheet_id != sheet_id);
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

    /// Remove any registry entries with the given `pivot_id`.
    pub fn unregister(&mut self, pivot_id: &str) {
        self.entries.retain(|e| e.pivot_id != pivot_id);
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
            let Some(new_dest) =
                rewrite_pivot_destination_for_structural_edit(entry.destination, ctx_sheet, edit)
            else {
                return false;
            };
            entry.destination = new_dest;
            true
        });
    }

    /// Rewrite any registered pivot destinations affected by a row/column insertion/deletion on a
    /// specific worksheet, identified by internal `sheet_id`.
    ///
    /// This is preferred over [`PivotRegistry::apply_structural_edit`] when callers have already
    /// resolved the edited sheet name (which may be either the worksheet's stable key or its
    /// display name) to a stable sheet id.
    ///
    /// Pivot registry entries store only `sheet_id`, so the edited sheet should be matched by id
    /// to avoid missing shifts when the host addresses a worksheet by a different alias.
    pub fn apply_structural_edit_with_sheet_id(
        &mut self,
        edited_sheet_id: usize,
        edit: &StructuralEdit,
        sheet_names: &HashMap<usize, String>,
    ) {
        let edit_sheet = match edit {
            StructuralEdit::InsertRows { sheet, .. }
            | StructuralEdit::DeleteRows { sheet, .. }
            | StructuralEdit::InsertCols { sheet, .. }
            | StructuralEdit::DeleteCols { sheet, .. } => sheet.as_str(),
        };

        self.entries.retain_mut(|entry| {
            if sheet_names.get(&entry.sheet_id).is_none() {
                // Sheet no longer exists; drop stale entry.
                return false;
            }
            if entry.sheet_id != edited_sheet_id {
                return true;
            }
            let Some(new_dest) =
                rewrite_pivot_destination_for_structural_edit(entry.destination, edit_sheet, edit)
            else {
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

    /// Rewrite any registered pivot destinations affected by a range move / insert-cells /
    /// delete-cells edit on a specific worksheet, identified by internal `sheet_id`.
    ///
    /// This is preferred over [`PivotRegistry::apply_range_map_edit`] when callers have already
    /// resolved the edited sheet name (which may be either the worksheet's stable key or its
    /// display name) to a stable sheet id.
    pub fn apply_range_map_edit_with_sheet_id(
        &mut self,
        edited_sheet_id: usize,
        edit: &RangeMapEdit,
        sheet_names: &HashMap<usize, String>,
    ) {
        let edit_sheet = edit.sheet.as_str();
        self.entries.retain_mut(|entry| {
            if sheet_names.get(&entry.sheet_id).is_none() {
                return false;
            }
            if entry.sheet_id != edited_sheet_id {
                return true;
            }
            let Some(new_dest) =
                rewrite_pivot_destination_for_range_map_edit(entry.destination, edit_sheet, edit)
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

#[cfg(test)]
mod normalize_pivot_cache_field_name_escape_tests {
    use super::normalize_pivot_cache_field_name;

    #[test]
    fn normalize_pivot_cache_field_name_preserves_bracket_escapes() {
        // Column refs escape `]` as `]]` within `[...]`.
        assert_eq!(
            normalize_pivot_cache_field_name("'Dim Product'[A]]B]").as_ref(),
            "Dim Product[A]]B]"
        );

        // Measures escape `]` the same way.
        assert_eq!(
            normalize_pivot_cache_field_name("[A]]B]").as_ref(),
            "[A]]B]"
        );
    }

    #[test]
    fn normalize_pivot_cache_field_name_escapes_dax_brackets() {
        // Bracket escapes inside identifiers should be canonicalized to DAX `]]` form.
        assert_eq!(
            normalize_pivot_cache_field_name("[Total]USD]").as_ref(),
            "[Total]]USD]"
        );
        assert_eq!(
            normalize_pivot_cache_field_name("Orders[Amount]USD]").as_ref(),
            "Orders[Amount]]USD]"
        );

        // Valid DAX-escaped refs should round-trip unchanged.
        assert_eq!(
            normalize_pivot_cache_field_name("[Total]]USD]").as_ref(),
            "[Total]]USD]"
        );
        assert_eq!(
            normalize_pivot_cache_field_name("Orders[Amount]]USD]").as_ref(),
            "Orders[Amount]]USD]"
        );

        // Quoted table names that themselves contain `[` should normalize without splitting on the
        // inner bracket.
        assert_eq!(
            normalize_pivot_cache_field_name("'My[Table]'[Col]").as_ref(),
            "'My[Table]'[Col]"
        );
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
    let (out, _) =
        rewrite_formula_for_structural_edit(&formula, ctx_sheet, crate::CellAddr::new(0, 0), edit);
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
    let (out, _) =
        rewrite_formula_for_range_map(&formula, ctx_sheet, crate::CellAddr::new(0, 0), edit);
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
