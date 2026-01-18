use core::fmt;
use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};

use crate::{CellRef, Range};

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MergeError {
    /// The requested merged region overlaps an existing merged region.
    Overlap(Range),
}

impl fmt::Display for MergeError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            MergeError::Overlap(r) => write!(f, "merged region overlaps existing merge: {r}"),
        }
    }
}

impl std::error::Error for MergeError {}

#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct MergedRegion {
    pub range: Range,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct RowSpan {
    col_start: u32,
    col_end: u32,
    region_idx: usize,
}

/// Per-worksheet merged-cell table with a compact row-index for fast lookup.
///
/// Excel models merged cells as a single cell anchored at the top-left corner of
/// the region. All addresses inside the region resolve to that anchor cell.
#[derive(Clone, Debug, Serialize, Default)]
pub struct MergedRegions {
    pub regions: Vec<MergedRegion>,
    #[serde(skip)]
    row_index: BTreeMap<u32, Vec<RowSpan>>,
}

impl MergedRegions {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.regions.is_empty()
    }

    #[must_use]
    pub fn region_count(&self) -> usize {
        self.regions.len()
    }

    #[must_use]
    pub fn iter(&self) -> impl Iterator<Item = &MergedRegion> {
        self.regions.iter()
    }

    #[must_use]
    pub fn containing_region(&self, cell: CellRef) -> Option<&MergedRegion> {
        let spans = self.row_index.get(&cell.row)?;
        for span in spans {
            if cell.col >= span.col_start && cell.col <= span.col_end {
                return self.regions.get(span.region_idx);
            }
        }
        None
    }

    #[must_use]
    pub fn containing_range(&self, cell: CellRef) -> Option<Range> {
        self.containing_region(cell).map(|r| r.range)
    }

    #[must_use]
    pub fn is_anchor(&self, cell: CellRef) -> bool {
        self.containing_range(cell)
            .is_some_and(|range| range.start == cell)
    }

    #[must_use]
    pub fn anchor_for(&self, cell: CellRef) -> Option<CellRef> {
        self.containing_range(cell).map(|range| range.start)
    }

    /// Resolve a cell address to the value-bearing anchor cell (top-left of a merge).
    #[must_use]
    pub fn resolve_cell(&self, cell: CellRef) -> CellRef {
        self.anchor_for(cell).unwrap_or(cell)
    }

    /// Add a merged region.
    ///
    /// `range` must not overlap any existing merged region. Single-cell ranges are ignored.
    pub fn add(&mut self, range: Range) -> Result<(), MergeError> {
        if range.is_single_cell() {
            return Ok(());
        }
        if self
            .regions
            .iter()
            .any(|r| ranges_intersect(r.range, range))
        {
            return Err(MergeError::Overlap(range));
        }

        self.regions.push(MergedRegion { range });
        self.rebuild_index();
        Ok(())
    }

    /// Remove any merged regions that intersect `range`.
    pub fn unmerge_range(&mut self, range: Range) -> usize {
        let before = self.regions.len();
        self.regions.retain(|r| !ranges_intersect(r.range, range));
        let removed = before - self.regions.len();
        if removed > 0 {
            self.rebuild_index();
        }
        removed
    }

    pub fn insert_rows(&mut self, at_row: u32, count: u32) {
        if count == 0 {
            return;
        }
        for region in &mut self.regions {
            let mut r = region.range;
            if at_row <= r.start.row {
                r.start.row += count;
                r.end.row += count;
            } else if at_row <= r.end.row {
                r.end.row += count;
            }
            region.range = r;
        }
        self.rebuild_index();
    }

    pub fn insert_cols(&mut self, at_col: u32, count: u32) {
        if count == 0 {
            return;
        }
        for region in &mut self.regions {
            let mut r = region.range;
            if at_col <= r.start.col {
                r.start.col += count;
                r.end.col += count;
            } else if at_col <= r.end.col {
                r.end.col += count;
            }
            region.range = r;
        }
        self.rebuild_index();
    }

    pub fn delete_rows(&mut self, at_row: u32, count: u32) {
        if count == 0 {
            return;
        }
        let del_end = at_row.saturating_add(count - 1);

        let mut new_regions: Vec<MergedRegion> = Vec::new();
        if new_regions.try_reserve_exact(self.regions.len()).is_err() {
            debug_assert!(
                false,
                "allocation failed (merged regions delete rows, regions={})",
                self.regions.len()
            );
            return;
        }
        for region in &self.regions {
            let r = region.range;
            if r.end.row < at_row {
                new_regions.push(*region);
                continue;
            }

            if r.start.row > del_end {
                let shifted = Range::new(
                    CellRef::new(r.start.row - count, r.start.col),
                    CellRef::new(r.end.row - count, r.end.col),
                );
                if !shifted.is_single_cell() {
                    new_regions.push(MergedRegion { range: shifted });
                }
                continue;
            }

            let new_start_row = if r.start.row >= at_row {
                at_row
            } else {
                r.start.row
            };
            let new_end_row = if r.end.row > del_end {
                r.end.row - count
            } else {
                at_row.saturating_sub(1)
            };

            if new_start_row > new_end_row {
                continue;
            }

            let adjusted = Range::new(
                CellRef::new(new_start_row, r.start.col),
                CellRef::new(new_end_row, r.end.col),
            );
            if !adjusted.is_single_cell() {
                new_regions.push(MergedRegion { range: adjusted });
            }
        }

        self.regions = new_regions;
        self.rebuild_index();
    }

    pub fn delete_cols(&mut self, at_col: u32, count: u32) {
        if count == 0 {
            return;
        }
        let del_end = at_col.saturating_add(count - 1);

        let mut new_regions: Vec<MergedRegion> = Vec::new();
        if new_regions.try_reserve_exact(self.regions.len()).is_err() {
            debug_assert!(
                false,
                "allocation failed (merged regions delete cols, regions={})",
                self.regions.len()
            );
            return;
        }
        for region in &self.regions {
            let r = region.range;
            if r.end.col < at_col {
                new_regions.push(*region);
                continue;
            }

            if r.start.col > del_end {
                let shifted = Range::new(
                    CellRef::new(r.start.row, r.start.col - count),
                    CellRef::new(r.end.row, r.end.col - count),
                );
                if !shifted.is_single_cell() {
                    new_regions.push(MergedRegion { range: shifted });
                }
                continue;
            }

            let new_start_col = if r.start.col >= at_col {
                at_col
            } else {
                r.start.col
            };
            let new_end_col = if r.end.col > del_end {
                r.end.col - count
            } else {
                at_col.saturating_sub(1)
            };

            if new_start_col > new_end_col {
                continue;
            }

            let adjusted = Range::new(
                CellRef::new(r.start.row, new_start_col),
                CellRef::new(r.end.row, new_end_col),
            );
            if !adjusted.is_single_cell() {
                new_regions.push(MergedRegion { range: adjusted });
            }
        }

        self.regions = new_regions;
        self.rebuild_index();
    }

    fn rebuild_index(&mut self) {
        self.row_index.clear();
        for (region_idx, region) in self.regions.iter().enumerate() {
            let r = region.range;
            for row in r.start.row..=r.end.row {
                self.row_index.entry(row).or_default().push(RowSpan {
                    col_start: r.start.col,
                    col_end: r.end.col,
                    region_idx,
                });
            }
        }
    }
}

impl<'de> Deserialize<'de> for MergedRegions {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        #[derive(Deserialize)]
        struct Helper {
            #[serde(default)]
            regions: Vec<MergedRegion>,
        }

        let helper = Helper::deserialize(deserializer)?;
        let mut regions = MergedRegions {
            regions: helper.regions,
            row_index: BTreeMap::new(),
        };
        regions.rebuild_index();
        Ok(regions)
    }
}

const fn ranges_intersect(a: Range, b: Range) -> bool {
    a.start.row <= b.end.row
        && a.end.row >= b.start.row
        && a.start.col <= b.end.col
        && a.end.col >= b.start.col
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolve_anchor() {
        let mut merges = MergedRegions::new();
        merges
            .add(Range::new(CellRef::new(0, 0), CellRef::new(1, 1)))
            .unwrap();

        assert!(merges.is_anchor(CellRef::new(0, 0)));
        assert!(!merges.is_anchor(CellRef::new(0, 1)));
        assert_eq!(merges.resolve_cell(CellRef::new(1, 1)), CellRef::new(0, 0));
    }

    #[test]
    fn insert_rows_expands_inside() {
        let mut merges = MergedRegions::new();
        merges
            .add(Range::new(CellRef::new(0, 0), CellRef::new(1, 1)))
            .unwrap();

        merges.insert_rows(1, 2);
        let range = merges
            .containing_range(CellRef::new(0, 0))
            .expect("missing range");
        assert_eq!(range, Range::new(CellRef::new(0, 0), CellRef::new(3, 1)));
    }

    #[test]
    fn delete_rows_shrinks() {
        let mut merges = MergedRegions::new();
        merges
            .add(Range::new(CellRef::new(0, 0), CellRef::new(3, 1)))
            .unwrap();

        merges.delete_rows(1, 2); // delete rows 1-2
        let range = merges
            .containing_range(CellRef::new(0, 0))
            .expect("missing range");
        assert_eq!(range, Range::new(CellRef::new(0, 0), CellRef::new(1, 1)));
    }
}
