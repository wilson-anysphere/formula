use serde::{Deserialize, Serialize};
use thiserror::Error;

use std::collections::HashMap;

use chrono::Datelike;

use crate::editing::rewrite::{
    rewrite_formula_for_range_map, rewrite_formula_for_structural_edit, GridRange, RangeMapEdit,
    StructuralEdit,
};
use crate::editing::EditOp;
use crate::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use crate::CellAddr;
use formula_model::{CellRef, Range, Style};

use super::source::coerce_pivot_value_with_number_format;
use super::{PivotApplyOptions, PivotCache, PivotConfig, PivotEngine, PivotError, PivotResult, PivotValue};

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
    /// Last output footprint written into the destination sheet.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_output_range: Option<Range>,
    /// If true, the pivot output may no longer match the sheet contents (e.g. an overlapping edit).
    #[serde(default)]
    pub needs_refresh: bool,
}

impl PivotTableDefinition {
    pub fn apply_edit_op(&mut self, op: &EditOp) {
        match op {
            EditOp::InsertRows { sheet, row, count } => {
                let edit = StructuralEdit::InsertRows {
                    sheet: sheet.clone(),
                    row: *row,
                    count: *count,
                };
                self.apply_structural_edit(&edit);
            }
            EditOp::DeleteRows { sheet, row, count } => {
                let edit = StructuralEdit::DeleteRows {
                    sheet: sheet.clone(),
                    row: *row,
                    count: *count,
                };
                self.apply_structural_edit(&edit);
            }
            EditOp::InsertCols { sheet, col, count } => {
                let edit = StructuralEdit::InsertCols {
                    sheet: sheet.clone(),
                    col: *col,
                    count: *count,
                };
                self.apply_structural_edit(&edit);
            }
            EditOp::DeleteCols { sheet, col, count } => {
                let edit = StructuralEdit::DeleteCols {
                    sheet: sheet.clone(),
                    col: *col,
                    count: *count,
                };
                self.apply_structural_edit(&edit);
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
                self.apply_range_map_edit(&edit);
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
                self.apply_range_map_edit(&edit);
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
                self.apply_range_map_edit(&edit);
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
                self.apply_range_map_edit(&edit);
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
                self.apply_range_map_edit(&edit);
            }
            // CopyRange does not move existing cells, so pivot definitions do not shift.
            EditOp::CopyRange { sheet, src, dst_top_left } => {
                let dst = Range::new(
                    *dst_top_left,
                    CellRef::new(
                        dst_top_left.row + src.height().saturating_sub(1),
                        dst_top_left.col + src.width().saturating_sub(1),
                    ),
                );
                self.invalidate_if_overlaps(sheet, &dst);
            }
            EditOp::Fill { sheet, src: _, dst } => {
                self.invalidate_if_overlaps(sheet, dst);
            }
        }
    }

    fn apply_structural_edit(&mut self, edit: &StructuralEdit) {
        let edit_sheet = match edit {
            StructuralEdit::InsertRows { sheet, .. }
            | StructuralEdit::DeleteRows { sheet, .. }
            | StructuralEdit::InsertCols { sheet, .. }
            | StructuralEdit::DeleteCols { sheet, .. } => sheet.as_str(),
        };

        // Destination top-left cell behaves like a cell reference.
        if self
            .destination
            .sheet
            .eq_ignore_ascii_case(edit_sheet)
        {
            if let Some(cell) =
                rewrite_cell_ref_for_structural_edit(self.destination.cell, &self.destination.sheet, edit)
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
            if sheet.eq_ignore_ascii_case(edit_sheet) {
                if let Some(r) = *range {
                    *range = rewrite_range_for_structural_edit(r, sheet, edit);
                    if range.is_none() {
                        self.needs_refresh = true;
                    }
                }
            }
        }

        // Update (or invalidate) last output footprint.
        if let Some(prev) = self.last_output_range {
            if self.destination.sheet.eq_ignore_ascii_case(edit_sheet) {
                self.last_output_range = rewrite_range_for_structural_edit(prev, &self.destination.sheet, edit);
                if self.last_output_range.is_none() {
                    self.needs_refresh = true;
                }
            }
        }

        // If the structural edit intersects the pivot output region, treat it as needing refresh.
        self.invalidate_if_structural_edit_overlaps_output(edit);
    }

    fn apply_range_map_edit(&mut self, edit: &RangeMapEdit) {
        let edit_sheet = edit.sheet.as_str();
        let prev_output = self.last_output_range;

        if self.destination.sheet.eq_ignore_ascii_case(edit_sheet) {
            if let Some(cell) =
                rewrite_cell_ref_for_range_map_edit(self.destination.cell, &self.destination.sheet, edit)
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
            if sheet.eq_ignore_ascii_case(edit_sheet) {
                if let Some(r) = *range {
                    *range = rewrite_range_for_range_map_edit(r, sheet, edit);
                    if range.is_none() {
                        self.needs_refresh = true;
                    }
                }
            }
        }

        if let Some(prev) = self.last_output_range {
            if self.destination.sheet.eq_ignore_ascii_case(edit_sheet) {
                self.last_output_range =
                    rewrite_range_for_range_map_edit(prev, &self.destination.sheet, edit);
                if self.last_output_range.is_none() {
                    self.needs_refresh = true;
                }
            }
        }

        self.invalidate_if_range_map_edit_overlaps_output(prev_output, edit);
    }

    fn invalidate_if_overlaps(&mut self, sheet: &str, region: &Range) {
        if !self.destination.sheet.eq_ignore_ascii_case(sheet) {
            return;
        }
        let Some(output) = self.last_output_range else {
            return;
        };
        if output.intersects(region) {
            self.needs_refresh = true;
        }
    }

    fn invalidate_if_structural_edit_overlaps_output(&mut self, edit: &StructuralEdit) {
        let Some(output) = self.last_output_range else {
            return;
        };
        let sheet = match edit {
            StructuralEdit::InsertRows { sheet, .. }
            | StructuralEdit::DeleteRows { sheet, .. }
            | StructuralEdit::InsertCols { sheet, .. }
            | StructuralEdit::DeleteCols { sheet, .. } => sheet.as_str(),
        };
        if !self.destination.sheet.eq_ignore_ascii_case(sheet) {
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
                let deleted = Range::new(
                    CellRef::new(*row, 0),
                    CellRef::new(del_end, u32::MAX),
                );
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
                let deleted = Range::new(
                    CellRef::new(0, *col),
                    CellRef::new(u32::MAX, del_end),
                );
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
    ) {
        let Some(output) = prev_output else {
            return;
        };
        if !self.destination.sheet.eq_ignore_ascii_case(&edit.sheet) {
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
        let Some(dst) = shift_grid_range_saturating(edit.moved_region, edit.delta_row, edit.delta_col) else {
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

fn shift_grid_range_saturating(range: GridRange, delta_row: i32, delta_col: i32) -> Option<GridRange> {
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
    let formula = format!("={}", cell.to_a1());
    let (out, _) = rewrite_formula_for_structural_edit(&formula, sheet, CellAddr::new(0, 0), edit);
    parse_a1_cell_from_formula(&out)
}

fn rewrite_range_for_structural_edit(range: Range, sheet: &str, edit: &StructuralEdit) -> Option<Range> {
    let formula = format!("={range}");
    let (out, _) = rewrite_formula_for_structural_edit(&formula, sheet, CellAddr::new(0, 0), edit);
    parse_a1_range_from_formula(&out)
}

fn rewrite_cell_ref_for_range_map_edit(
    cell: CellRef,
    sheet: &str,
    edit: &RangeMapEdit,
) -> Option<CellRef> {
    let formula = format!("={}", cell.to_a1());
    let (out, _) = rewrite_formula_for_range_map(&formula, sheet, CellAddr::new(0, 0), edit);
    parse_a1_cell_from_formula(&out)
}

fn rewrite_range_for_range_map_edit(range: Range, sheet: &str, edit: &RangeMapEdit) -> Option<Range> {
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
    fn write_cell(
        &mut self,
        sheet: &str,
        addr: &str,
        value: crate::value::Value,
    ) -> Result<(), crate::EngineError>;
    fn clear_cell(&mut self, sheet: &str, addr: &str) -> Result<(), crate::EngineError>;
    fn resolve_table(&mut self, table_id: u32) -> Option<(String, Range)>;
}

pub(crate) fn refresh_pivot(
    ctx: &mut impl PivotRefreshContext,
    def: &mut PivotTableDefinition,
) -> Result<PivotRefreshOutput, PivotRefreshError> {
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

    let mut data: Vec<Vec<PivotValue>> = Vec::new();
    for row in source_range.start.row..=source_range.end.row {
        let mut out_row = Vec::new();
        for col in source_range.start.col..=source_range.end.col {
            let addr = CellRef::new(row, col).to_a1();
            let value = ctx.read_cell(&source_sheet, &addr);
            let pivot_value = engine_value_to_pivot_value(value);
            let number_format = ctx.read_cell_number_format(&source_sheet, &addr);
            out_row.push(coerce_pivot_value_with_number_format(
                pivot_value,
                number_format.as_deref(),
                ctx.date_system(),
            ));
        }
        data.push(out_row);
    }

    let cache = PivotCache::from_range(&data)?;
    // Note: computed pivot caches may be reused in the future; for MVP we rebuild each time.
    let result = PivotEngine::calculate(&cache, &def.config)?;

    // Clear old output footprint.
    if let Some(prev) = def.last_output_range {
        for row in prev.start.row..=prev.end.row {
            for col in prev.start.col..=prev.end.col {
                let addr = CellRef::new(row, col).to_a1();
                ctx.clear_cell(&def.destination.sheet, &addr).ok();
            }
        }
    }

    let rows = result.data.len() as u32;
    let cols = result.data.first().map(|r| r.len()).unwrap_or(0) as u32;
    if rows == 0 || cols == 0 {
        // Treat empty results as clearing output.
        def.last_output_range = None;
        def.needs_refresh = false;
        return Ok(PivotRefreshOutput {
            output_range: None,
            result,
        });
    }

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
    let output_range = Range::new(def.destination.cell, CellRef::new(end_row, end_col));

    let pivot_cell_writes = result.to_cell_writes_with_formats(
        super::CellRef {
            row: def.destination.cell.row,
            col: def.destination.cell.col,
        },
        &def.config,
        &PivotApplyOptions::default(),
    );

    let mut style_cache: HashMap<String, u32> = HashMap::new();
    let date_system = ctx.date_system();

    for write in pivot_cell_writes {
        let addr = CellRef::new(write.row, write.col).to_a1();

        if let Some(fmt) = write.number_format.as_deref() {
            let style_id = *style_cache.entry(fmt.to_string()).or_insert_with(|| {
                ctx.intern_style(Style {
                    number_format: Some(fmt.to_string()),
                    ..Style::default()
                })
            });
            ctx.set_cell_style_id(&def.destination.sheet, &addr, style_id)
                .map_err(|_| PivotRefreshError::OutputOutOfBounds)?;
        }

        // Write values after styles so "precision as displayed" rounding (when enabled) uses the
        // final number format.
        let v = pivot_value_to_engine_value(write.value, date_system);
        ctx.write_cell(&def.destination.sheet, &addr, v)
            .map_err(|_| PivotRefreshError::OutputOutOfBounds)?;
    }

    def.last_output_range = Some(output_range);
    def.needs_refresh = false;

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
        // Dates are represented as numbers in the core engine today.
        Value::Error(_) => PivotValue::Blank,
        Value::Reference(_) => PivotValue::Blank,
        Value::ReferenceUnion(_) => PivotValue::Blank,
        Value::Array(_) => PivotValue::Blank,
        Value::Lambda(_) => PivotValue::Blank,
        Value::Spill { .. } => PivotValue::Blank,
    }
}

fn pivot_value_to_engine_value(value: PivotValue, date_system: ExcelDateSystem) -> crate::value::Value {
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
