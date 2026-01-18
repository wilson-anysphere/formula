use serde::{Deserialize, Serialize};
use thiserror::Error;

use std::collections::HashMap;

use chrono::Datelike;

use crate::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use crate::editing::rewrite::{
    rewrite_formula_for_range_map, rewrite_formula_for_structural_edit, GridRange, RangeMapEdit,
    StructuralEdit,
};
use crate::editing::EditOp;
use crate::CellAddr;
use formula_model::{sheet_name_eq_case_insensitive, CellRef, Range, Style, EXCEL_MAX_COLS};

use super::source::coerce_pivot_value_with_number_format;
use super::{
    Layout, PivotApplyOptions, PivotCache, PivotConfig, PivotEngine, PivotError, PivotResult,
    PivotTable, PivotValue, ShowAsType,
};

/// Stable identifier for a pivot table stored in the engine.
///
/// NOTE: This is currently engine-local (allocated monotonically). Higher-level
/// model layers may want to use UUIDs; this is an internal MVP.
pub type PivotTableId = u64;

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotDestination {
    pub sheet: String,
    pub cell: CellRef,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum PivotSource {
    /// A worksheet range (including header row).
    Range { sheet: String, range: Option<Range> },
    /// A stable Excel table id. The table's current range is tracked by the table object itself.
    Table { table_id: u32 },
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotTableDefinition {
    /// Pivot id (engine-local).
    pub id: PivotTableId,
    pub name: String,
    pub source: PivotSource,
    pub destination: PivotDestination,
    pub config: PivotConfig,
    /// Whether to apply number formats from value fields when rendering pivot output.
    ///
    /// In XLSX, this corresponds to `pivotTableDefinition@applyNumberFormats` (defaulting to true).
    #[serde(default = "default_true")]
    pub apply_number_formats: bool,
    /// Last output footprint written into the destination sheet.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_output_range: Option<Range>,
    /// If true, the pivot output may no longer match the sheet contents (e.g. an overlapping edit).
    #[serde(default)]
    pub needs_refresh: bool,
}

fn default_true() -> bool {
    true
}

impl PivotTableDefinition {
    pub fn apply_edit_op(&mut self, op: &EditOp) {
        let mut resolver = |_name: &str| None;
        self.apply_edit_op_with_sheet_resolver(op, &mut resolver);
    }

    /// Apply a structural edit operation to this pivot definition, using a caller-provided sheet
    /// resolver to match sheet-key/display-name aliases.
    ///
    /// `resolve_sheet_id` should return a stable sheet id for a given sheet name (either stable key
    /// or user-visible display name). When available, pivot metadata will treat two sheet names as
    /// equivalent if they resolve to the same sheet id, even when the raw strings differ.
    pub fn apply_edit_op_with_sheet_resolver(
        &mut self,
        op: &EditOp,
        resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    ) {
        match op {
            EditOp::InsertRows { sheet, row, count } => {
                let edit = StructuralEdit::InsertRows {
                    sheet: sheet.clone(),
                    row: *row,
                    count: *count,
                };
                self.apply_structural_edit(&edit, resolve_sheet_id);
            }
            EditOp::DeleteRows { sheet, row, count } => {
                let edit = StructuralEdit::DeleteRows {
                    sheet: sheet.clone(),
                    row: *row,
                    count: *count,
                };
                self.apply_structural_edit(&edit, resolve_sheet_id);
            }
            EditOp::InsertCols { sheet, col, count } => {
                let edit = StructuralEdit::InsertCols {
                    sheet: sheet.clone(),
                    col: *col,
                    count: *count,
                };
                self.apply_structural_edit(&edit, resolve_sheet_id);
            }
            EditOp::DeleteCols { sheet, col, count } => {
                let edit = StructuralEdit::DeleteCols {
                    sheet: sheet.clone(),
                    col: *col,
                    count: *count,
                };
                self.apply_structural_edit(&edit, resolve_sheet_id);
            }
            EditOp::InsertCellsShiftRight { sheet, range } => {
                let width = range.width();
                let edit = RangeMapEdit {
                    sheet: sheet.clone(),
                    moved_region: GridRange::new(
                        range.start.row,
                        range.start.col,
                        range.end.row,
                        u32::MAX,
                    ),
                    delta_row: 0,
                    delta_col: width as i32,
                    deleted_region: None,
                };
                self.apply_range_map_edit(&edit, resolve_sheet_id);
            }
            EditOp::InsertCellsShiftDown { sheet, range } => {
                let height = range.height();
                let edit = RangeMapEdit {
                    sheet: sheet.clone(),
                    moved_region: GridRange::new(
                        range.start.row,
                        range.start.col,
                        u32::MAX,
                        range.end.col,
                    ),
                    delta_row: height as i32,
                    delta_col: 0,
                    deleted_region: None,
                };
                self.apply_range_map_edit(&edit, resolve_sheet_id);
            }
            EditOp::DeleteCellsShiftLeft { sheet, range } => {
                let width = range.width();
                let start_col = range.end.col.saturating_add(1);
                let edit = RangeMapEdit {
                    sheet: sheet.clone(),
                    moved_region: GridRange::new(
                        range.start.row,
                        start_col,
                        range.end.row,
                        u32::MAX,
                    ),
                    delta_row: 0,
                    delta_col: -(width as i32),
                    deleted_region: Some(GridRange::new(
                        range.start.row,
                        range.start.col,
                        range.end.row,
                        range.end.col,
                    )),
                };
                self.apply_range_map_edit(&edit, resolve_sheet_id);
            }
            EditOp::DeleteCellsShiftUp { sheet, range } => {
                let height = range.height();
                let start_row = range.end.row.saturating_add(1);
                let edit = RangeMapEdit {
                    sheet: sheet.clone(),
                    moved_region: GridRange::new(
                        start_row,
                        range.start.col,
                        u32::MAX,
                        range.end.col,
                    ),
                    delta_row: -(height as i32),
                    delta_col: 0,
                    deleted_region: Some(GridRange::new(
                        range.start.row,
                        range.start.col,
                        range.end.row,
                        range.end.col,
                    )),
                };
                self.apply_range_map_edit(&edit, resolve_sheet_id);
            }
            EditOp::MoveRange {
                sheet,
                src,
                dst_top_left,
            } => {
                let edit = RangeMapEdit {
                    sheet: sheet.clone(),
                    moved_region: GridRange::new(
                        src.start.row,
                        src.start.col,
                        src.end.row,
                        src.end.col,
                    ),
                    delta_row: dst_top_left.row as i32 - src.start.row as i32,
                    delta_col: dst_top_left.col as i32 - src.start.col as i32,
                    deleted_region: None,
                };
                self.apply_range_map_edit(&edit, resolve_sheet_id);
            }
            // CopyRange does not move existing cells, so pivot definitions do not shift.
            EditOp::CopyRange {
                sheet,
                src,
                dst_top_left,
            } => {
                let dst = Range::new(
                    *dst_top_left,
                    CellRef::new(
                        dst_top_left.row + src.height().saturating_sub(1),
                        dst_top_left.col + src.width().saturating_sub(1),
                    ),
                );
                self.invalidate_if_overlaps(sheet, &dst, resolve_sheet_id);
            }
            EditOp::Fill { sheet, src: _, dst } => {
                self.invalidate_if_overlaps(sheet, dst, resolve_sheet_id);
            }
        }
    }

    fn apply_structural_edit(
        &mut self,
        edit: &StructuralEdit,
        resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    ) {
        let edit_sheet = match edit {
            StructuralEdit::InsertRows { sheet, .. }
            | StructuralEdit::DeleteRows { sheet, .. }
            | StructuralEdit::InsertCols { sheet, .. }
            | StructuralEdit::DeleteCols { sheet, .. } => sheet.as_str(),
        };

        // Destination top-left cell behaves like a cell reference.
        if sheet_matches(&self.destination.sheet, edit_sheet, resolve_sheet_id) {
            if let Some(cell) =
                rewrite_cell_ref_for_structural_edit(self.destination.cell, edit_sheet, edit)
            {
                self.destination.cell = cell;
            } else {
                // If the destination anchor is deleted, clamp it to the edit start.
                match edit {
                    StructuralEdit::DeleteRows { row, .. } => self.destination.cell.row = *row,
                    StructuralEdit::DeleteCols { col, .. } => self.destination.cell.col = *col,
                    _ => {}
                }
                self.needs_refresh = true;
            }
        }

        // Update source range reference.
        if let PivotSource::Range { sheet, range } = &mut self.source {
            if sheet_matches(sheet, edit_sheet, resolve_sheet_id) {
                if let Some(r) = *range {
                    *range = rewrite_range_for_structural_edit(r, edit_sheet, edit);
                    if range.is_none() {
                        self.needs_refresh = true;
                    }
                }
            }
        }

        // Update (or invalidate) last output footprint.
        if let Some(prev) = self.last_output_range {
            if sheet_matches(&self.destination.sheet, edit_sheet, resolve_sheet_id) {
                self.last_output_range = rewrite_range_for_structural_edit(prev, edit_sheet, edit);
                if self.last_output_range.is_none() {
                    self.needs_refresh = true;
                }
            }
        }

        // If the structural edit intersects the pivot output region, treat it as needing refresh.
        self.invalidate_if_structural_edit_overlaps_output(edit, resolve_sheet_id);
    }

    fn apply_range_map_edit(
        &mut self,
        edit: &RangeMapEdit,
        resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    ) {
        let edit_sheet = edit.sheet.as_str();
        let prev_output = self.last_output_range;

        if sheet_matches(&self.destination.sheet, edit_sheet, resolve_sheet_id) {
            if let Some(cell) =
                rewrite_cell_ref_for_range_map_edit(self.destination.cell, edit_sheet, edit)
            {
                self.destination.cell = cell;
            } else {
                // Deleted/moved out of bounds; best-effort clamp to origin of deleted region.
                if let Some(deleted) = edit.deleted_region {
                    self.destination.cell.row = deleted.start_row;
                    self.destination.cell.col = deleted.start_col;
                }
                self.needs_refresh = true;
            }
        }

        if let PivotSource::Range { sheet, range } = &mut self.source {
            if sheet_matches(sheet, edit_sheet, resolve_sheet_id) {
                if let Some(r) = *range {
                    *range = rewrite_range_for_range_map_edit(r, edit_sheet, edit);
                    if range.is_none() {
                        self.needs_refresh = true;
                    }
                }
            }
        }

        if let Some(prev) = self.last_output_range {
            if sheet_matches(&self.destination.sheet, edit_sheet, resolve_sheet_id) {
                self.last_output_range = rewrite_range_for_range_map_edit(prev, edit_sheet, edit);
                if self.last_output_range.is_none() {
                    self.needs_refresh = true;
                }
            }
        }

        self.invalidate_if_range_map_edit_overlaps_output(prev_output, edit, resolve_sheet_id);
    }

    fn invalidate_if_overlaps(
        &mut self,
        sheet: &str,
        region: &Range,
        resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    ) {
        if !sheet_matches(&self.destination.sheet, sheet, resolve_sheet_id) {
            return;
        }
        let Some(output) = self.last_output_range else {
            return;
        };
        if output.intersects(region) {
            self.needs_refresh = true;
        }
    }

    fn invalidate_if_structural_edit_overlaps_output(
        &mut self,
        edit: &StructuralEdit,
        resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    ) {
        let Some(output) = self.last_output_range else {
            return;
        };
        let sheet = match edit {
            StructuralEdit::InsertRows { sheet, .. }
            | StructuralEdit::DeleteRows { sheet, .. }
            | StructuralEdit::InsertCols { sheet, .. }
            | StructuralEdit::DeleteCols { sheet, .. } => sheet.as_str(),
        };
        if !sheet_matches(&self.destination.sheet, sheet, resolve_sheet_id) {
            return;
        }

        match edit {
            StructuralEdit::InsertRows { row, .. } => {
                // Inserting rows *at* the pivot's first row simply shifts the pivot down (Excel-like).
                // Only insertions *inside* the existing output footprint require a re-render.
                if *row > output.start.row && *row <= output.end.row {
                    self.needs_refresh = true;
                }
            }
            StructuralEdit::DeleteRows { row, count, .. } => {
                let del_end = row.saturating_add(count.saturating_sub(1));
                let deleted = Range::new(CellRef::new(*row, 0), CellRef::new(del_end, u32::MAX));
                if output.intersects(&deleted) {
                    self.needs_refresh = true;
                }
            }
            StructuralEdit::InsertCols { col, .. } => {
                // Same semantics as row insertion: inserting at the left edge shifts the pivot,
                // inserting inside the existing output footprint invalidates the rendered output.
                if *col > output.start.col && *col <= output.end.col {
                    self.needs_refresh = true;
                }
            }
            StructuralEdit::DeleteCols { col, count, .. } => {
                let del_end = col.saturating_add(count.saturating_sub(1));
                let deleted = Range::new(CellRef::new(0, *col), CellRef::new(u32::MAX, del_end));
                if output.intersects(&deleted) {
                    self.needs_refresh = true;
                }
            }
        }
    }

    fn invalidate_if_range_map_edit_overlaps_output(
        &mut self,
        prev_output: Option<Range>,
        edit: &RangeMapEdit,
        resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
    ) {
        let Some(output) = prev_output else {
            return;
        };
        if !sheet_matches(&self.destination.sheet, &edit.sheet, resolve_sheet_id) {
            return;
        }

        // Deletions inside the output always invalidate it.
        if let Some(deleted) = edit.deleted_region {
            let deleted_range = Range::new(
                CellRef::new(deleted.start_row, deleted.start_col),
                CellRef::new(deleted.end_row, deleted.end_col),
            );
            if output.intersects(&deleted_range) {
                self.needs_refresh = true;
                return;
            }
        }

        // Cells moved *into* the output range invalidate it (e.g. MoveRange destination overlaps output).
        let Some(dst) =
            shift_grid_range_saturating(edit.moved_region, edit.delta_row, edit.delta_col)
        else {
            return;
        };
        let dst_range = Range::new(
            CellRef::new(dst.start_row, dst.start_col),
            CellRef::new(dst.end_row, dst.end_col),
        );
        if output.intersects(&dst_range) {
            self.needs_refresh = true;
        }
    }
}

fn sheet_matches(
    left: &str,
    right: &str,
    resolve_sheet_id: &mut impl FnMut(&str) -> Option<usize>,
) -> bool {
    match (resolve_sheet_id(left), resolve_sheet_id(right)) {
        (Some(a), Some(b)) => a == b,
        _ => sheet_name_eq_case_insensitive(left, right),
    }
}

fn shift_grid_range_saturating(
    range: GridRange,
    delta_row: i32,
    delta_col: i32,
) -> Option<GridRange> {
    let sr = range.start_row as i64 + delta_row as i64;
    let sc = range.start_col as i64 + delta_col as i64;
    let er = range.end_row as i64 + delta_row as i64;
    let ec = range.end_col as i64 + delta_col as i64;
    if sr < 0 || sc < 0 || er < 0 || ec < 0 {
        return None;
    }
    let clamp_u32 = |v: i64| -> u32 {
        if v >= u32::MAX as i64 {
            u32::MAX
        } else {
            v as u32
        }
    };
    Some(GridRange::new(
        clamp_u32(sr),
        clamp_u32(sc),
        clamp_u32(er),
        clamp_u32(ec),
    ))
}

fn rewrite_cell_ref_for_structural_edit(
    cell: CellRef,
    sheet: &str,
    edit: &StructuralEdit,
) -> Option<CellRef> {
    let mut formula = String::new();
    formula.push('=');
    formula_model::push_a1_cell_ref(cell.row, cell.col, false, false, &mut formula);
    let (out, _) = rewrite_formula_for_structural_edit(&formula, sheet, CellAddr::new(0, 0), edit);
    parse_a1_cell_from_formula(&out)
}

fn rewrite_range_for_structural_edit(
    range: Range,
    sheet: &str,
    edit: &StructuralEdit,
) -> Option<Range> {
    let formula = format!("={range}");
    let (out, _) = rewrite_formula_for_structural_edit(&formula, sheet, CellAddr::new(0, 0), edit);
    parse_a1_range_from_formula(&out)
}

fn rewrite_cell_ref_for_range_map_edit(
    cell: CellRef,
    sheet: &str,
    edit: &RangeMapEdit,
) -> Option<CellRef> {
    let mut formula = String::new();
    formula.push('=');
    formula_model::push_a1_cell_ref(cell.row, cell.col, false, false, &mut formula);
    let (out, _) = rewrite_formula_for_range_map(&formula, sheet, CellAddr::new(0, 0), edit);
    parse_a1_cell_from_formula(&out)
}

fn rewrite_range_for_range_map_edit(
    range: Range,
    sheet: &str,
    edit: &RangeMapEdit,
) -> Option<Range> {
    let formula = format!("={range}");
    let (out, _) = rewrite_formula_for_range_map(&formula, sheet, CellAddr::new(0, 0), edit);
    parse_a1_range_from_formula(&out)
}

fn parse_a1_cell_from_formula(formula: &str) -> Option<CellRef> {
    let expr = parse_formula_expr(formula)?;
    if expr.eq_ignore_ascii_case("#REF!") {
        return None;
    }
    if expr.contains(',') || expr.contains(' ') || expr.contains('(') || expr.contains(')') {
        return None;
    }
    CellRef::from_a1(expr).ok()
}

fn parse_a1_range_from_formula(formula: &str) -> Option<Range> {
    let expr = parse_formula_expr(formula)?;
    if expr.eq_ignore_ascii_case("#REF!") {
        return None;
    }
    if expr.contains(',') || expr.contains(' ') || expr.contains('(') || expr.contains(')') {
        return None;
    }
    Range::from_a1(expr).ok()
}

fn parse_formula_expr(formula: &str) -> Option<&str> {
    let trimmed = formula.trim_start();
    let expr = trimmed.strip_prefix('=').unwrap_or(trimmed);
    Some(expr.trim())
}

#[derive(Debug, Error)]
pub enum PivotRefreshError {
    #[error("unknown pivot id: {0}")]
    UnknownPivot(PivotTableId),
    #[error("pivot source is invalid")]
    InvalidSource,
    #[error("table not found: {0}")]
    TableNotFound(u32),
    #[error(transparent)]
    Pivot(#[from] PivotError),
    #[error("output range exceeds sheet bounds")]
    OutputOutOfBounds,
}

#[derive(Clone, Debug, PartialEq)]
pub struct PivotRefreshOutput {
    pub output_range: Option<Range>,
    pub result: PivotResult,
}

pub(crate) trait PivotRefreshContext {
    fn read_cell(&mut self, sheet: &str, addr: &str) -> crate::value::Value;
    fn read_cell_number_format(&self, sheet: &str, addr: &str) -> Option<String>;
    fn date_system(&self) -> ExcelDateSystem;
    fn intern_style(&mut self, style: Style) -> u32;
    fn set_cell_style_id(
        &mut self,
        sheet: &str,
        addr: &str,
        style_id: u32,
    ) -> Result<(), crate::EngineError>;
    /// Bulk-apply per-cell style ids.
    ///
    /// Default implementation falls back to per-cell [`PivotRefreshContext::set_cell_style_id`]
    /// calls.
    fn set_cell_style_ids(
        &mut self,
        sheet: &str,
        writes: &[(CellRef, u32)],
    ) -> Result<(), crate::EngineError> {
        let mut addr = String::new();
        for (cell, style_id) in writes {
            addr.clear();
            formula_model::push_a1_cell_ref(cell.row, cell.col, false, false, &mut addr);
            self.set_cell_style_id(sheet, addr.as_str(), *style_id)?;
        }
        Ok(())
    }
    fn write_cell(
        &mut self,
        sheet: &str,
        addr: &str,
        value: crate::value::Value,
    ) -> Result<(), crate::EngineError>;
    fn clear_cell(&mut self, sheet: &str, addr: &str) -> Result<(), crate::EngineError>;
    fn resolve_table(&mut self, table_id: u32) -> Option<(String, Range)>;

    /// Materialize a [`PivotCache`] from the current workbook state.
    ///
    /// Engines may override this to reuse an optimized cache builder (for example one that can
    /// resolve number formats for date inference).
    fn pivot_cache_from_range(
        &mut self,
        sheet: &str,
        range: Range,
    ) -> Result<PivotCache, PivotError> {
        let mut data: Vec<Vec<PivotValue>> = Vec::new();
        let mut addr = String::new();
        for row in range.start.row..=range.end.row {
            let mut out_row = Vec::new();
            for col in range.start.col..=range.end.col {
                addr.clear();
                formula_model::push_a1_cell_ref(row, col, false, false, &mut addr);
                let value = self.read_cell(sheet, addr.as_str());
                let pivot_value = engine_value_to_pivot_value(value);
                let number_format = self.read_cell_number_format(sheet, addr.as_str());
                out_row.push(coerce_pivot_value_with_number_format(
                    pivot_value,
                    number_format.as_deref(),
                    self.date_system(),
                ));
            }
            data.push(out_row);
        }
        PivotCache::from_range(&data)
    }

    /// Bulk-clear a rectangular range of cells.
    ///
    /// Default implementation falls back to per-cell [`PivotRefreshContext::clear_cell`] calls.
    fn clear_range(&mut self, sheet: &str, range: Range) -> Result<(), crate::EngineError> {
        let mut addr = String::new();
        for cell in range.iter() {
            addr.clear();
            formula_model::push_a1_cell_ref(cell.row, cell.col, false, false, &mut addr);
            self.clear_cell(sheet, addr.as_str())?;
        }
        Ok(())
    }

    /// Bulk-apply a rectangular range of values.
    ///
    /// Default implementation falls back to per-cell [`PivotRefreshContext::write_cell`] calls.
    fn set_range_values(
        &mut self,
        sheet: &str,
        range: Range,
        values: &[Vec<crate::value::Value>],
    ) -> Result<(), crate::EngineError> {
        let expected_rows = range.height() as usize;
        let expected_cols = range.width() as usize;

        if values.len() != expected_rows {
            let actual_cols = values.get(0).map(|row| row.len()).unwrap_or(0);
            return Err(crate::EngineError::RangeValuesDimensionMismatch {
                expected_rows,
                expected_cols,
                actual_rows: values.len(),
                actual_cols,
            });
        }
        for row in values {
            if row.len() != expected_cols {
                return Err(crate::EngineError::RangeValuesDimensionMismatch {
                    expected_rows,
                    expected_cols,
                    actual_rows: values.len(),
                    actual_cols: row.len(),
                });
            }
        }

        let mut addr = String::new();
        for (r_off, row_values) in values.iter().enumerate() {
            let row = range.start.row + r_off as u32;
            for (c_off, value) in row_values.iter().enumerate() {
                let col = range.start.col + c_off as u32;
                addr.clear();
                formula_model::push_a1_cell_ref(row, col, false, false, &mut addr);
                self.write_cell(sheet, addr.as_str(), value.clone())?;
            }
        }

        Ok(())
    }

    /// Register a pivot table's metadata for `GETPIVOTDATA` (best-effort).
    fn register_pivot_table(
        &mut self,
        _sheet: &str,
        _destination: Range,
        _pivot: PivotTable,
    ) -> Result<(), crate::pivot_registry::PivotRegistryError> {
        Ok(())
    }

    /// Unregister a pivot table's metadata for `GETPIVOTDATA` (best-effort).
    fn unregister_pivot_table(&mut self, _pivot_id: &str) {}
}

fn pivot_registry_id(id: PivotTableId) -> String {
    format!("engine-pivot-{id}")
}

pub(crate) fn refresh_pivot(
    ctx: &mut impl PivotRefreshContext,
    def: &mut PivotTableDefinition,
) -> Result<PivotRefreshOutput, PivotRefreshError> {
    let registry_pivot_id = pivot_registry_id(def.id);
    let (source_sheet, source_range) = match &def.source {
        PivotSource::Range { sheet, range } => {
            let Some(range) = range.as_ref().copied() else {
                return Err(PivotRefreshError::InvalidSource);
            };
            (sheet.clone(), range)
        }
        PivotSource::Table { table_id } => ctx
            .resolve_table(*table_id)
            .ok_or(PivotRefreshError::TableNotFound(*table_id))?,
    };

    let cache = ctx.pivot_cache_from_range(&source_sheet, source_range)?;
    // Note: computed pivot caches may be reused in the future; for MVP we rebuild each time.
    let result = PivotEngine::calculate(&cache, &def.config)?;

    let rows =
        u32::try_from(result.data.len()).map_err(|_| PivotRefreshError::OutputOutOfBounds)?;
    let cols = u32::try_from(result.data.first().map(|r| r.len()).unwrap_or(0))
        .map_err(|_| PivotRefreshError::OutputOutOfBounds)?;

    if rows == 0 || cols == 0 {
        // Treat empty results as clearing output.
        if let Some(prev) = def.last_output_range {
            ctx.clear_range(&def.destination.sheet, prev).ok();
        }
        ctx.unregister_pivot_table(&registry_pivot_id);
        def.last_output_range = None;
        def.needs_refresh = false;
        return Ok(PivotRefreshOutput {
            output_range: None,
            result,
        });
    }

    // Compute and validate the new output footprint *before* clearing any existing output so that
    // out-of-bounds failures don't wipe the previously rendered pivot.
    let end_row = def
        .destination
        .cell
        .row
        .checked_add(rows.saturating_sub(1))
        .ok_or(PivotRefreshError::OutputOutOfBounds)?;
    let end_col = def
        .destination
        .cell
        .col
        .checked_add(cols.saturating_sub(1))
        .ok_or(PivotRefreshError::OutputOutOfBounds)?;
    if end_row >= i32::MAX as u32 || end_col >= EXCEL_MAX_COLS {
        return Err(PivotRefreshError::OutputOutOfBounds);
    }
    let output_range = Range::new(def.destination.cell, CellRef::new(end_row, end_col));

    let prev_output_range = def.last_output_range;

    // When the pivot output grows (or moves), cells in the newly-covered area may already contain
    // values or formatting. Clear those cells up-front so applying the pivot output does not
    // preserve that stale state.
    if let Some(prev) = prev_output_range {
        for newly_covered in stale_ranges(output_range, prev) {
            ctx.clear_range(&def.destination.sheet, newly_covered).ok();
        }
    }

    let options = PivotApplyOptions {
        apply_number_formats: def.apply_number_formats,
        ..PivotApplyOptions::default()
    };
    let value_field_count = def.config.value_fields.len();
    let row_label_width = match def.config.layout {
        Layout::Compact => 1,
        Layout::Outline | Layout::Tabular => def.config.row_fields.len(),
    };

    // Apply styles first so "precision as displayed" rounding (when enabled) sees the final number
    // formats when we write values below.
    let mut style_cache: HashMap<String, u32> = HashMap::new();
    let date_system = ctx.date_system();
    let mut values: Vec<Vec<crate::value::Value>> = Vec::new();
    if values.try_reserve_exact(rows as usize).is_err() {
        return Err(PivotError::AllocationFailure("pivot refresh values").into());
    }
    let mut style_writes: Vec<(CellRef, u32)> = Vec::new();

    for r in 0..rows as usize {
        let mut row_out: Vec<crate::value::Value> = Vec::new();
        if row_out.try_reserve_exact(cols as usize).is_err() {
            return Err(PivotError::AllocationFailure("pivot refresh row values").into());
        }
        let src_row = result.data.get(r);
        for c in 0..cols as usize {
            let pv = src_row
                .and_then(|row| row.get(c))
                .cloned()
                .unwrap_or(PivotValue::Blank);

            let number_format: Option<&str> = if matches!(pv, PivotValue::Date(_)) {
                Some(options.default_date_number_format.as_str())
            } else if options.apply_number_formats
                && r > 0
                && value_field_count > 0
                && c >= row_label_width
            {
                let vf_idx = (c - row_label_width) % value_field_count;
                let vf = &def.config.value_fields[vf_idx];
                if let Some(fmt) = vf.number_format.as_deref() {
                    Some(fmt)
                } else if is_percent_show_as(vf.show_as) {
                    Some(options.default_percent_number_format.as_str())
                } else {
                    None
                }
            } else {
                None
            };

            let dest_cell = CellRef::new(
                def.destination.cell.row + r as u32,
                def.destination.cell.col + c as u32,
            );

            let mut desired_style_id = None;
            if let Some(fmt) = number_format {
                let style_id = match style_cache.get(fmt) {
                    Some(id) => *id,
                    None => {
                        let style_id = ctx.intern_style(Style {
                            number_format: Some(fmt.to_string()),
                            ..Style::default()
                        });
                        style_cache.insert(fmt.to_string(), style_id);
                        style_id
                    }
                };
                desired_style_id = Some(style_id);
            } else if prev_output_range.is_some_and(|prev| prev.contains(dest_cell))
                && (matches!(pv, PivotValue::Number(_))
                    || (r > 0 && value_field_count > 0 && c >= row_label_width))
            {
                // When a refresh removes a previously-applied number format (e.g. the user toggles
                // `apply_number_formats` off, or a date-typed source column becomes a plain
                // number), reset previously-rendered numeric/value-area cells to the default style
                // so we don't leave stale number formats behind.
                desired_style_id = Some(0);
            }

            if let Some(style_id) = desired_style_id {
                style_writes.push((dest_cell, style_id));
            }

            row_out.push(pivot_value_to_engine_value(pv, date_system));
        }
        values.push(row_out);
    }

    if !style_writes.is_empty() {
        ctx.set_cell_style_ids(&def.destination.sheet, &style_writes)
            .map_err(|_| PivotRefreshError::OutputOutOfBounds)?;
    }

    ctx.set_range_values(&def.destination.sheet, output_range, &values)
        .map_err(|_| PivotRefreshError::OutputOutOfBounds)?;

    // Clear any stale cells from the previous output footprint that now fall outside the updated
    // output range.
    //
    // This is done after successfully writing the new output so refresh failures (e.g. due to
    // out-of-bounds anchors) do not wipe the prior rendered pivot.
    if let Some(prev) = prev_output_range {
        for stale in stale_ranges(prev, output_range) {
            ctx.clear_range(&def.destination.sheet, stale).ok();
        }
    }

    def.last_output_range = Some(output_range);
    def.needs_refresh = false;

    // Register pivot metadata for `GETPIVOTDATA` (best-effort).
    //
    // Note: We register after successfully writing output so that if refresh fails part-way through
    // (e.g. due to an unexpected write error), we don't replace the previous registry entry.
    ctx.register_pivot_table(
        &def.destination.sheet,
        output_range,
        PivotTable {
            id: registry_pivot_id,
            name: def.name.clone(),
            config: def.config.clone(),
            cache,
        },
    )
    .ok();

    Ok(PivotRefreshOutput {
        output_range: Some(output_range),
        result,
    })
}

fn engine_value_to_pivot_value(value: crate::value::Value) -> PivotValue {
    use crate::value::Value;
    match value {
        Value::Blank => PivotValue::Blank,
        Value::Number(n) => PivotValue::Number(n),
        Value::Text(s) => PivotValue::Text(s),
        Value::Bool(b) => PivotValue::Bool(b),
        Value::Entity(e) => PivotValue::Text(e.display),
        Value::Record(r) => PivotValue::Text(r.display),
        // Errors are treated as text for pivot purposes (so aggregations won't treat them as numbers).
        Value::Error(e) => PivotValue::Text(e.as_code().to_string()),
        Value::Reference(_) => PivotValue::Blank,
        Value::ReferenceUnion(_) => PivotValue::Blank,
        Value::Array(arr) => {
            // If a dynamic array somehow lands in a cell value, match Excel's visible behavior
            // (top-left value shown in the origin cell).
            let top_left = arr.top_left();
            engine_value_to_pivot_value(top_left)
        }
        Value::Lambda(_) => PivotValue::Blank,
        Value::Spill { .. } => PivotValue::Blank,
    }
}

fn pivot_value_to_engine_value(
    value: PivotValue,
    date_system: ExcelDateSystem,
) -> crate::value::Value {
    match value {
        PivotValue::Blank => crate::value::Value::Blank,
        PivotValue::Number(n) => crate::value::Value::Number(n),
        PivotValue::Text(s) => crate::value::Value::Text(s),
        PivotValue::Bool(b) => crate::value::Value::Bool(b),
        PivotValue::Date(d) => {
            let excel_date = ExcelDate::new(d.year(), d.month() as u8, d.day() as u8);
            match ymd_to_serial(excel_date, date_system) {
                Ok(serial) => crate::value::Value::Number(serial as f64),
                Err(_) => crate::value::Value::Blank,
            }
        }
    }
}

fn is_percent_show_as(show_as: Option<ShowAsType>) -> bool {
    matches!(
        show_as.unwrap_or(ShowAsType::Normal),
        ShowAsType::PercentOfGrandTotal
            | ShowAsType::PercentOfRowTotal
            | ShowAsType::PercentOfColumnTotal
            | ShowAsType::PercentOf
            | ShowAsType::PercentDifferenceFrom
    )
}

fn stale_ranges(prev: Range, next: Range) -> Vec<Range> {
    let Some(inter) = prev.intersection(&next) else {
        return vec![prev];
    };

    let mut out = Vec::new();

    // Bands above/below the intersection span the full previous width.
    if prev.start.row < inter.start.row {
        out.push(Range::new(
            prev.start,
            CellRef::new(inter.start.row.saturating_sub(1), prev.end.col),
        ));
    }
    if inter.end.row < prev.end.row {
        out.push(Range::new(
            CellRef::new(inter.end.row.saturating_add(1), prev.start.col),
            prev.end,
        ));
    }

    // Bands left/right span the overlapping rows.
    if prev.start.col < inter.start.col {
        out.push(Range::new(
            CellRef::new(inter.start.row, prev.start.col),
            CellRef::new(inter.end.row, inter.start.col.saturating_sub(1)),
        ));
    }
    if inter.end.col < prev.end.col {
        out.push(Range::new(
            CellRef::new(inter.start.row, inter.end.col.saturating_add(1)),
            CellRef::new(inter.end.row, prev.end.col),
        ));
    }

    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pivot_definition_sheet_matching_is_unicode_case_insensitive() {
        // German sharp s uppercases to "SS" in Unicode.
        let mut def = PivotTableDefinition {
            id: 1,
            name: "Pivot".to_string(),
            source: PivotSource::Range {
                sheet: "ß".to_string(),
                range: None,
            },
            destination: PivotDestination {
                sheet: "ß".to_string(),
                cell: CellRef::new(5, 0), // A6 (0-indexed row)
            },
            config: PivotConfig::default(),
            apply_number_formats: true,
            last_output_range: None,
            needs_refresh: false,
        };

        // Insert a row on the same sheet, referenced using a casefold-equivalent name.
        def.apply_edit_op(&EditOp::InsertRows {
            sheet: "SS".to_string(),
            row: 0,
            count: 1,
        });

        // Destination anchor should shift down by one row.
        assert_eq!(def.destination.cell, CellRef::new(6, 0));
    }
}
