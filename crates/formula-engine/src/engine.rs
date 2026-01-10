use crate::eval::{
    parse_a1, CellAddr, CompiledExpr, Expr, FormulaParseError, Parser, RangeRef, SheetReference,
};
use crate::editing::{
    CellChange, CellSnapshot, EditError, EditOp, EditResult, FormulaRewrite, MovedRange,
};
use crate::editing::rewrite::{
    rewrite_formula_for_copy_delta, rewrite_formula_for_range_map, rewrite_formula_for_structural_edit,
    GridRange, RangeMapEdit, StructuralEdit,
};
use crate::locale::{canonicalize_formula, FormulaLocale};
use crate::value::{ErrorKind, Value};
use formula_model::{CellRef, Range, Table};
use rayon::prelude::*;
use std::cmp::max;
use std::collections::{BTreeSet, HashMap, HashSet, VecDeque};
use thiserror::Error;

pub type SheetId = usize;

#[derive(Debug, Error)]
pub enum EngineError {
    #[error(transparent)]
    Address(#[from] crate::eval::AddressParseError),
    #[error(transparent)]
    Parse(#[from] FormulaParseError),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RecalcMode {
    SingleThreaded,
    MultiThreaded,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct CellKey {
    sheet: SheetId,
    addr: CellAddr,
}

#[derive(Debug, Clone)]
struct Cell {
    value: Value,
    formula: Option<String>,
    ast: Option<CompiledExpr>,
    volatile: bool,
    thread_safe: bool,
}

impl Default for Cell {
    fn default() -> Self {
        Self {
            value: Value::Blank,
            formula: None,
            ast: None,
            volatile: false,
            thread_safe: true,
        }
    }
}

#[derive(Debug, Default, Clone)]
struct Sheet {
    cells: HashMap<CellAddr, Cell>,
    tables: Vec<Table>,
}

#[derive(Debug, Default, Clone)]
struct Workbook {
    sheets: Vec<Sheet>,
    sheet_name_to_id: HashMap<String, SheetId>,
}

impl Workbook {
    fn ensure_sheet(&mut self, name: &str) -> SheetId {
        if let Some(id) = self.sheet_name_to_id.get(name).copied() {
            return id;
        }
        let id = self.sheets.len();
        self.sheets.push(Sheet {
            cells: HashMap::new(),
            tables: Vec::new(),
        });
        self.sheet_name_to_id.insert(name.to_string(), id);
        id
    }

    fn sheet_id(&self, name: &str) -> Option<SheetId> {
        self.sheet_name_to_id.get(name).copied()
    }

    fn get_cell(&self, key: CellKey) -> Option<&Cell> {
        self.sheets.get(key.sheet)?.cells.get(&key.addr)
    }

    fn get_or_create_cell_mut(&mut self, key: CellKey) -> &mut Cell {
        self.sheets[key.sheet].cells.entry(key.addr).or_default()
    }

    fn get_cell_value(&self, key: CellKey) -> Value {
        self.get_cell(key).map(|c| c.value.clone()).unwrap_or(Value::Blank)
    }

    fn set_tables(&mut self, sheet: SheetId, tables: Vec<Table>) {
        if let Some(s) = self.sheets.get_mut(sheet) {
            s.tables = tables;
        }
    }
}

#[derive(Debug, Default)]
struct DependencyGraph {
    precedents: HashMap<CellKey, HashSet<CellKey>>,
    dependents: HashMap<CellKey, HashSet<CellKey>>,
    volatile_cells: HashSet<CellKey>,
}

impl DependencyGraph {
    fn set_precedents(&mut self, cell: CellKey, new_precedents: HashSet<CellKey>) {
        if let Some(old) = self.precedents.remove(&cell) {
            for p in old {
                if let Some(deps) = self.dependents.get_mut(&p) {
                    deps.remove(&cell);
                    if deps.is_empty() {
                        self.dependents.remove(&p);
                    }
                }
            }
        }

        for p in &new_precedents {
            self.dependents.entry(*p).or_default().insert(cell);
        }

        if !new_precedents.is_empty() {
            self.precedents.insert(cell, new_precedents);
        }
    }

    fn clear_cell(&mut self, cell: CellKey) {
        self.set_precedents(cell, HashSet::new());
        self.volatile_cells.remove(&cell);
    }

    // Dirty-marking with reason tracking is implemented in `Engine` (where we can
    // store the predecessor edge used for diagnostics). The graph itself only
    // maintains adjacency lists.
}

pub struct Engine {
    workbook: Workbook,
    graph: DependencyGraph,
    dirty: HashSet<CellKey>,
    dirty_reasons: HashMap<CellKey, CellKey>,
}

impl Default for Engine {
    fn default() -> Self {
        Self::new()
    }
}

impl Engine {
    pub fn new() -> Self {
        Self {
            workbook: Workbook::default(),
            graph: DependencyGraph::default(),
            dirty: HashSet::new(),
            dirty_reasons: HashMap::new(),
        }
    }

    pub fn set_cell_value(
        &mut self,
        sheet: &str,
        addr: &str,
        value: impl Into<Value>,
    ) -> Result<(), EngineError> {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let addr = parse_a1(addr)?;
        let key = CellKey { sheet: sheet_id, addr };

        // Replace any existing formula and dependencies.
        self.graph.clear_cell(key);
        self.dirty.remove(&key);
        self.dirty_reasons.remove(&key);

        let cell = self.workbook.get_or_create_cell_mut(key);
        cell.value = value.into();
        cell.formula = None;
        cell.ast = None;
        cell.volatile = false;
        cell.thread_safe = true;

        // Mark downstream dependents dirty.
        self.mark_dirty_dependents_with_reasons(key);
        Ok(())
    }

    /// Replace the set of tables for a given worksheet.
    ///
    /// Tables are needed to resolve structured references like `Table1[Col]` and `[@Col]`.
    pub fn set_sheet_tables(&mut self, sheet: &str, tables: Vec<Table>) {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        self.workbook.set_tables(sheet_id, tables);
        // Mark all formulas dirty; structured reference dependencies may have changed.
        for (addr, cell) in self.workbook.sheets[sheet_id].cells.iter() {
            if cell.formula.is_some() {
                self.dirty.insert(CellKey {
                    sheet: sheet_id,
                    addr: *addr,
                });
            }
        }
    }

    pub fn set_cell_formula(
        &mut self,
        sheet: &str,
        addr: &str,
        formula: &str,
    ) -> Result<(), EngineError> {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let addr = parse_a1(addr)?;
        let key = CellKey { sheet: sheet_id, addr };

        let parsed = Parser::parse(formula)?;
        let compiled = self.compile_expr(&parsed, sheet_id);
        let tables_by_sheet: Vec<Vec<Table>> =
            self.workbook.sheets.iter().map(|s| s.tables.clone()).collect();
        let (precedents, volatile, thread_safe) = analyze_expr(&compiled, key, &tables_by_sheet);

        self.graph.set_precedents(key, precedents);

        if volatile {
            self.graph.volatile_cells.insert(key);
        } else {
            self.graph.volatile_cells.remove(&key);
        }

        let cell = self.workbook.get_or_create_cell_mut(key);
        cell.formula = Some(formula.to_string());
        cell.ast = Some(compiled);
        cell.volatile = volatile;
        cell.thread_safe = thread_safe;

        // Recalculate this cell and anything depending on it.
        self.mark_dirty_including_self_with_reasons(key);
        Ok(())
    }

    /// Set a cell formula that was entered in a locale-specific display format.
    ///
    /// This converts the incoming formula to canonical form before parsing and
    /// persistence. Canonical form uses English function names and `,`/`.` for
    /// separators, which matches XLSX expectations and keeps storage stable across
    /// UI locale changes.
    pub fn set_cell_formula_localized(
        &mut self,
        sheet: &str,
        addr: &str,
        localized_formula: &str,
        locale: &FormulaLocale,
    ) -> Result<(), EngineError> {
        let canonical = canonicalize_formula(localized_formula, locale)?;
        self.set_cell_formula(sheet, addr, &canonical)
    }

    pub fn get_cell_value(&self, sheet: &str, addr: &str) -> Value {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Value::Blank;
        };
        let Ok(addr) = parse_a1(addr) else {
            return Value::Error(ErrorKind::Ref);
        };
        self.workbook
            .get_cell_value(CellKey { sheet: sheet_id, addr })
    }

    pub fn get_cell_formula(&self, sheet: &str, addr: &str) -> Option<&str> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey { sheet: sheet_id, addr };
        self.workbook.get_cell(key)?.formula.as_deref()
    }

    pub fn apply_operation(&mut self, op: EditOp) -> Result<EditResult, EditError> {
        let before = self.workbook.clone();
        let mut formula_rewrites = Vec::new();
        let mut moved_ranges = Vec::new();

        let sheet_names = sheet_names_by_id(&self.workbook);

        match op {
            EditOp::InsertRows { sheet, row, count } => {
                if count == 0 {
                    return Err(EditError::InvalidCount);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                shift_rows(&mut self.workbook.sheets[sheet_id], row, count, true);
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    StructuralEdit::InsertRows { sheet, row, count },
                ));
            }
            EditOp::DeleteRows { sheet, row, count } => {
                if count == 0 {
                    return Err(EditError::InvalidCount);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                shift_rows(&mut self.workbook.sheets[sheet_id], row, count, false);
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    StructuralEdit::DeleteRows { sheet, row, count },
                ));
            }
            EditOp::InsertCols { sheet, col, count } => {
                if count == 0 {
                    return Err(EditError::InvalidCount);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                shift_cols(&mut self.workbook.sheets[sheet_id], col, count, true);
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    StructuralEdit::InsertCols { sheet, col, count },
                ));
            }
            EditOp::DeleteCols { sheet, col, count } => {
                if count == 0 {
                    return Err(EditError::InvalidCount);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                shift_cols(&mut self.workbook.sheets[sheet_id], col, count, false);
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    StructuralEdit::DeleteCols { sheet, col, count },
                ));
            }
            EditOp::InsertCellsShiftRight { sheet, range } => {
                let width = range.width();
                if width == 0 {
                    return Err(EditError::InvalidRange);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                insert_cells_shift_right(&mut self.workbook.sheets[sheet_id], range, width);
                let edit = RangeMapEdit {
                    sheet,
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
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
            }
            EditOp::InsertCellsShiftDown { sheet, range } => {
                let height = range.height();
                if height == 0 {
                    return Err(EditError::InvalidRange);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                insert_cells_shift_down(&mut self.workbook.sheets[sheet_id], range, height);
                let edit = RangeMapEdit {
                    sheet,
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
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
            }
            EditOp::DeleteCellsShiftLeft { sheet, range } => {
                let width = range.width();
                if width == 0 {
                    return Err(EditError::InvalidRange);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                delete_cells_shift_left(&mut self.workbook.sheets[sheet_id], range, width);
                let start_col = range.end.col.saturating_add(1);
                let edit = RangeMapEdit {
                    sheet,
                    moved_region: GridRange::new(range.start.row, start_col, range.end.row, u32::MAX),
                    delta_row: 0,
                    delta_col: -(width as i32),
                    deleted_region: Some(GridRange::new(
                        range.start.row,
                        range.start.col,
                        range.end.row,
                        range.end.col,
                    )),
                };
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
            }
            EditOp::DeleteCellsShiftUp { sheet, range } => {
                let height = range.height();
                if height == 0 {
                    return Err(EditError::InvalidRange);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                delete_cells_shift_up(&mut self.workbook.sheets[sheet_id], range, height);
                let start_row = range.end.row.saturating_add(1);
                let edit = RangeMapEdit {
                    sheet,
                    moved_region: GridRange::new(start_row, range.start.col, u32::MAX, range.end.col),
                    delta_row: -(height as i32),
                    delta_col: 0,
                    deleted_region: Some(GridRange::new(
                        range.start.row,
                        range.start.col,
                        range.end.row,
                        range.end.col,
                    )),
                };
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
            }
            EditOp::MoveRange {
                sheet,
                src,
                dst_top_left,
            } => {
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                if src.width() == 0 || src.height() == 0 {
                    return Err(EditError::InvalidRange);
                }
                let dst = Range::new(
                    dst_top_left,
                    CellRef::new(
                        dst_top_left.row + src.height() - 1,
                        dst_top_left.col + src.width() - 1,
                    ),
                );
                if ranges_overlap(src, dst) {
                    return Err(EditError::OverlappingMove);
                }
                move_range(&mut self.workbook.sheets[sheet_id], src, dst_top_left);
                let edit = RangeMapEdit {
                    sheet: sheet.clone(),
                    moved_region: GridRange::new(src.start.row, src.start.col, src.end.row, src.end.col),
                    delta_row: dst.start.row as i32 - src.start.row as i32,
                    delta_col: dst.start.col as i32 - src.start.col as i32,
                    deleted_region: None,
                };
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
                moved_ranges.push(MovedRange { sheet, from: src, to: dst });
            }
            EditOp::CopyRange {
                sheet,
                src,
                dst_top_left,
            } => {
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                if src.width() == 0 || src.height() == 0 {
                    return Err(EditError::InvalidRange);
                }
                copy_range(
                    &mut self.workbook.sheets[sheet_id],
                    &sheet,
                    src,
                    dst_top_left,
                    &mut formula_rewrites,
                );
            }
            EditOp::Fill { sheet, src, dst } => {
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                fill_range(
                    &mut self.workbook.sheets[sheet_id],
                    &sheet,
                    src,
                    dst,
                    &mut formula_rewrites,
                );
            }
        }

        self.rebuild_graph()
            .map_err(|e| EditError::Engine(e.to_string()))?;

        let sheet_names_after = sheet_names_by_id(&self.workbook);
        let changed_cells = diff_workbooks(&before, &self.workbook, &sheet_names_after);

        Ok(EditResult {
            changed_cells,
            moved_ranges,
            formula_rewrites,
        })
    }

    pub fn recalculate(&mut self) {
        // Default to multithreaded when rayon is available.
        self.recalculate_with_mode(RecalcMode::MultiThreaded);
    }

    pub fn recalculate_single_threaded(&mut self) {
        self.recalculate_with_mode(RecalcMode::SingleThreaded);
    }

    pub fn recalculate_multi_threaded(&mut self) {
        self.recalculate_with_mode(RecalcMode::MultiThreaded);
    }

    fn recalculate_with_mode(&mut self, mode: RecalcMode) {
        // Volatile formulas recalculate every time, and so do their dependents.
        let volatile_cells: Vec<CellKey> = self.graph.volatile_cells.iter().copied().collect();
        for cell in volatile_cells {
            self.mark_dirty_including_self_with_reasons(cell);
        }

        if self.dirty.is_empty() {
            return;
        }

        let mut snapshot = Snapshot::from_workbook(&self.workbook);

        let dirty_cells: HashSet<CellKey> = self.dirty.iter().copied().collect();
        let mut in_degree: HashMap<CellKey, usize> = HashMap::new();
        for &cell in &dirty_cells {
            let deg = self
                .graph
                .precedents
                .get(&cell)
                .map(|p| p.iter().filter(|c| dirty_cells.contains(c)).count())
                .unwrap_or(0);
            in_degree.insert(cell, deg);
        }

        let mut current_level: Vec<CellKey> = in_degree
            .iter()
            .filter_map(|(&k, &d)| (d == 0).then_some(k))
            .collect();

        let mut processed = HashSet::new();

        while !current_level.is_empty() {
            let level = std::mem::take(&mut current_level);
            let has_barrier = level.iter().any(|&k| {
                self.workbook
                    .get_cell(k)
                    .map(|c| c.volatile || !c.thread_safe)
                    .unwrap_or(false)
            });

            let tasks: Vec<(CellKey, CompiledExpr)> = level
                .iter()
                .filter_map(|&k| self.workbook.get_cell(k).and_then(|c| c.ast.clone().map(|a| (k, a))))
                .collect();

            let results: Vec<(CellKey, Value)> = if mode == RecalcMode::MultiThreaded && !has_barrier
            {
                tasks
                    .par_iter()
                    .map(|(k, expr)| {
                        let ctx = crate::eval::EvalContext {
                            current_sheet: k.sheet,
                            current_cell: k.addr,
                        };
                        let evaluator = crate::eval::Evaluator::new(&snapshot, ctx);
                        (*k, evaluator.eval_formula(expr))
                    })
                    .collect()
            } else {
                tasks
                    .iter()
                    .map(|(k, expr)| {
                        let ctx = crate::eval::EvalContext {
                            current_sheet: k.sheet,
                            current_cell: k.addr,
                        };
                        let evaluator = crate::eval::Evaluator::new(&snapshot, ctx);
                        (*k, evaluator.eval_formula(expr))
                    })
                    .collect()
            };

            for (k, v) in &results {
                let cell = self.workbook.get_or_create_cell_mut(*k);
                cell.value = v.clone();
                snapshot.values.insert(*k, v.clone());
            }

            for &k in &level {
                processed.insert(k);
                self.dirty.remove(&k);
                self.dirty_reasons.remove(&k);
                if let Some(deps) = self.graph.dependents.get(&k) {
                    for &d in deps {
                        if !dirty_cells.contains(&d) {
                            continue;
                        }
                        if let Some(entry) = in_degree.get_mut(&d) {
                            *entry = entry.saturating_sub(1);
                            if *entry == 0 {
                                current_level.push(d);
                            }
                        }
                    }
                }
            }
        }

        // Any remaining dirty cells are in a cycle. For now, surface a calc error.
        for cell in dirty_cells.difference(&processed).copied().collect::<Vec<_>>() {
            let c = self.workbook.get_or_create_cell_mut(cell);
            c.value = Value::Error(ErrorKind::Calc);
            self.dirty.remove(&cell);
            self.dirty_reasons.remove(&cell);
        }
    }

    fn compile_expr(&mut self, expr: &Expr<String>, _current_sheet: SheetId) -> CompiledExpr {
        let mut map = |sref: &SheetReference<String>| match sref {
            SheetReference::Current => SheetReference::Current,
            SheetReference::Sheet(name) => SheetReference::Sheet(self.workbook.ensure_sheet(name)),
            SheetReference::External(wb) => SheetReference::External(wb.clone()),
        };
        expr.map_sheets(&mut map)
    }

    fn rebuild_graph(&mut self) -> Result<(), EngineError> {
        let sheet_names = sheet_names_by_id(&self.workbook);
        let mut formulas: Vec<(String, CellAddr, String)> = Vec::new();
        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            let Some(sheet_name) = sheet_names.get(&sheet_id).cloned() else {
                continue;
            };
            for (addr, cell) in &sheet.cells {
                if let Some(formula) = &cell.formula {
                    formulas.push((sheet_name.clone(), *addr, formula.clone()));
                }
            }
        }

        self.graph = DependencyGraph::default();
        self.dirty.clear();
        self.dirty_reasons.clear();

        for (sheet_name, addr, formula) in formulas {
            let addr_a1 = cell_addr_to_a1(addr);
            self.set_cell_formula(&sheet_name, &addr_a1, &formula)?;
        }
        Ok(())
    }

    /// Returns whether a cell is currently marked dirty (needs recalculation).
    pub fn is_dirty(&self, sheet: &str, addr: &str) -> bool {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return false;
        };
        let Ok(addr) = parse_a1(addr) else {
            return false;
        };
        self.dirty.contains(&CellKey { sheet: sheet_id, addr })
    }

    /// Direct precedents (cells referenced by the formula in `cell`).
    pub fn precedents(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        self.precedents_impl(sheet, addr, false)
    }

    /// Transitive precedents (all cells that can influence `cell`).
    pub fn precedents_transitive(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        self.precedents_impl(sheet, addr, true)
    }

    /// Direct dependents (cells whose formulas reference `cell`).
    pub fn dependents(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        self.dependents_impl(sheet, addr, false)
    }

    /// Transitive dependents (all downstream cells that are affected by `cell`).
    pub fn dependents_transitive(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        self.dependents_impl(sheet, addr, true)
    }

    /// Returns a dependency path explaining why `cell` is currently dirty.
    ///
    /// The returned vector is ordered from the root cause (usually an edited
    /// input cell) to the provided `cell`.
    pub fn dirty_dependency_path(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Option<Vec<(SheetId, CellAddr)>> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey { sheet: sheet_id, addr };
        if !self.dirty.contains(&key) {
            return None;
        }

        let mut path = vec![key];
        let mut current = key;
        let mut guard = 0usize;
        while let Some(prev) = self.dirty_reasons.get(&current).copied() {
            path.push(prev);
            current = prev;
            guard += 1;
            if guard > 10_000 {
                break;
            }
        }
        path.reverse();
        Some(path.into_iter().map(|k| (k.sheet, k.addr)).collect())
    }

    /// Deterministically evaluates a cell's formula while capturing a per-node trace.
    ///
    /// This is intended for on-demand debugging and does **not** mutate engine state.
    pub fn debug_evaluate(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<crate::debug::DebugEvaluation, EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Err(EngineError::Parse(FormulaParseError::UnexpectedToken(format!(
                "unknown sheet '{sheet}'"
            ))));
        };
        let addr = parse_a1(addr)?;
        let key = CellKey { sheet: sheet_id, addr };
        let cell = self.workbook.get_cell(key);
        let Some(formula) = cell.and_then(|c| c.formula.as_deref()) else {
            return Err(EngineError::Parse(FormulaParseError::UnexpectedToken(
                "cell has no formula".to_string(),
            )));
        };

        let snapshot = Snapshot::from_workbook(&self.workbook);
        let ctx = crate::eval::EvalContext {
            current_sheet: sheet_id,
            current_cell: addr,
        };

        // Parse with spans, compile sheet references without mutating the workbook,
        // then evaluate with tracing.
        let parsed = crate::debug::parse_spanned_formula(formula)?;
        let mut map = |sref: &SheetReference<String>| match sref {
            SheetReference::Current => SheetReference::Current,
            SheetReference::Sheet(name) => self
                .workbook
                .sheet_id(name)
                .map(SheetReference::Sheet)
                .unwrap_or_else(|| SheetReference::External(name.clone())),
            SheetReference::External(wb) => SheetReference::External(wb.clone()),
        };
        let compiled = parsed.map_sheets(&mut map);

        let (value, trace) = crate::debug::evaluate_with_trace(&snapshot, ctx, &compiled);

        Ok(crate::debug::DebugEvaluation {
            formula: formula.to_string(),
            value,
            trace,
        })
    }

    fn precedents_impl(
        &self,
        sheet: &str,
        addr: &str,
        transitive: bool,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(Vec::new());
        };
        let addr = parse_a1(addr)?;
        let key = CellKey { sheet: sheet_id, addr };
        let nodes = if transitive {
            collect_transitive(&self.graph.precedents, key)
        } else {
            self.graph
                .precedents
                .get(&key)
                .map(|s| sorted_cell_keys(s))
                .unwrap_or_default()
        };
        Ok(nodes.into_iter().map(|k| (k.sheet, k.addr)).collect())
    }

    fn dependents_impl(
        &self,
        sheet: &str,
        addr: &str,
        transitive: bool,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(Vec::new());
        };
        let addr = parse_a1(addr)?;
        let key = CellKey { sheet: sheet_id, addr };
        let nodes = if transitive {
            collect_transitive(&self.graph.dependents, key)
        } else {
            self.graph
                .dependents
                .get(&key)
                .map(|s| sorted_cell_keys(s))
                .unwrap_or_default()
        };
        Ok(nodes.into_iter().map(|k| (k.sheet, k.addr)).collect())
    }

    fn mark_dirty_including_self_with_reasons(&mut self, from: CellKey) {
        self.dirty.insert(from);
        self.dirty_reasons.remove(&from);
        self.mark_dirty_dependents_with_reasons(from);
    }

    fn mark_dirty_dependents_with_reasons(&mut self, from: CellKey) {
        let mut queue: VecDeque<(CellKey, CellKey)> = VecDeque::new();
        if let Some(deps) = self.graph.dependents.get(&from) {
            for dep in sorted_cell_keys(deps) {
                queue.push_back((from, dep));
            }
        }

        while let Some((cause, cell)) = queue.pop_front() {
            if !self.dirty.insert(cell) {
                continue;
            }
            self.dirty_reasons.entry(cell).or_insert(cause);
            if let Some(deps) = self.graph.dependents.get(&cell) {
                for dep in sorted_cell_keys(deps) {
                    queue.push_back((cell, dep));
                }
            }
        }
    }
}

fn sheet_names_by_id(workbook: &Workbook) -> HashMap<SheetId, String> {
    workbook
        .sheet_name_to_id
        .iter()
        .map(|(name, id)| (*id, name.clone()))
        .collect()
}

fn cell_ref_from_addr(addr: CellAddr) -> CellRef {
    CellRef::new(addr.row, addr.col)
}

fn cell_addr_from_cell_ref(cell: CellRef) -> CellAddr {
    CellAddr { row: cell.row, col: cell.col }
}

fn cell_addr_to_a1(addr: CellAddr) -> String {
    format!("{}{}", col_to_name(addr.col), addr.row + 1)
}

fn col_to_name(col: u32) -> String {
    let mut n = col + 1;
    let mut out = Vec::<u8>::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        out.push(b'A' + rem as u8);
        n = (n - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).expect("column letters are ASCII")
}

fn ranges_overlap(a: Range, b: Range) -> bool {
    !(a.end.row < b.start.row
        || a.start.row > b.end.row
        || a.end.col < b.start.col
        || a.start.col > b.end.col)
}

fn shift_rows(sheet: &mut Sheet, row: u32, count: u32, insert: bool) {
    let del_end = row.saturating_add(count.saturating_sub(1));
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if insert {
            if addr.row >= row {
                new_cells.insert(
                    CellAddr {
                        row: addr.row + count,
                        col: addr.col,
                    },
                    cell,
                );
            } else {
                new_cells.insert(addr, cell);
            }
            continue;
        }

        if addr.row < row {
            new_cells.insert(addr, cell);
        } else if addr.row > del_end {
            new_cells.insert(
                CellAddr {
                    row: addr.row - count,
                    col: addr.col,
                },
                cell,
            );
        }
    }
    sheet.cells = new_cells;
}

fn shift_cols(sheet: &mut Sheet, col: u32, count: u32, insert: bool) {
    let del_end = col.saturating_add(count.saturating_sub(1));
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if insert {
            if addr.col >= col {
                new_cells.insert(
                    CellAddr {
                        row: addr.row,
                        col: addr.col + count,
                    },
                    cell,
                );
            } else {
                new_cells.insert(addr, cell);
            }
            continue;
        }

        if addr.col < col {
            new_cells.insert(addr, cell);
        } else if addr.col > del_end {
            new_cells.insert(
                CellAddr {
                    row: addr.row,
                    col: addr.col - count,
                },
                cell,
            );
        }
    }
    sheet.cells = new_cells;
}

fn insert_cells_shift_right(sheet: &mut Sheet, range: Range, width: u32) {
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if addr.row >= range.start.row && addr.row <= range.end.row && addr.col >= range.start.col {
            new_cells.insert(
                CellAddr {
                    row: addr.row,
                    col: addr.col + width,
                },
                cell,
            );
        } else {
            new_cells.insert(addr, cell);
        }
    }
    sheet.cells = new_cells;
}

fn insert_cells_shift_down(sheet: &mut Sheet, range: Range, height: u32) {
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if addr.col >= range.start.col && addr.col <= range.end.col && addr.row >= range.start.row {
            new_cells.insert(
                CellAddr {
                    row: addr.row + height,
                    col: addr.col,
                },
                cell,
            );
        } else {
            new_cells.insert(addr, cell);
        }
    }
    sheet.cells = new_cells;
}

fn delete_cells_shift_left(sheet: &mut Sheet, range: Range, width: u32) {
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if addr.row >= range.start.row && addr.row <= range.end.row {
            if addr.col >= range.start.col && addr.col <= range.end.col {
                continue;
            }
            if addr.col > range.end.col {
                new_cells.insert(
                    CellAddr {
                        row: addr.row,
                        col: addr.col - width,
                    },
                    cell,
                );
            } else {
                new_cells.insert(addr, cell);
            }
        } else {
            new_cells.insert(addr, cell);
        }
    }
    sheet.cells = new_cells;
}

fn delete_cells_shift_up(sheet: &mut Sheet, range: Range, height: u32) {
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if addr.col >= range.start.col && addr.col <= range.end.col {
            if addr.row >= range.start.row && addr.row <= range.end.row {
                continue;
            }
            if addr.row > range.end.row {
                new_cells.insert(
                    CellAddr {
                        row: addr.row - height,
                        col: addr.col,
                    },
                    cell,
                );
            } else {
                new_cells.insert(addr, cell);
            }
        } else {
            new_cells.insert(addr, cell);
        }
    }
    sheet.cells = new_cells;
}

fn move_range(sheet: &mut Sheet, src: Range, dst_top_left: CellRef) {
    let dst_top_left_addr = cell_addr_from_cell_ref(dst_top_left);
    let dst = Range::new(
        dst_top_left,
        CellRef::new(
            dst_top_left.row + src.height() - 1,
            dst_top_left.col + src.width() - 1,
        ),
    );

    let mut extracted: Vec<(CellRef, Option<Cell>)> = Vec::new();
    for cell in src.iter() {
        extracted.push((cell, sheet.cells.remove(&cell_addr_from_cell_ref(cell))));
    }

    for cell in dst.iter() {
        sheet.cells.remove(&cell_addr_from_cell_ref(cell));
    }

    for (cell, value) in extracted {
        let Some(value) = value else { continue };
        let dr = cell.row - src.start.row;
        let dc = cell.col - src.start.col;
        sheet.cells.insert(
            CellAddr {
                row: dst_top_left_addr.row + dr,
                col: dst_top_left_addr.col + dc,
            },
            value,
        );
    }
}

fn copy_range(
    sheet: &mut Sheet,
    sheet_name: &str,
    src: Range,
    dst_top_left: CellRef,
    formula_rewrites: &mut Vec<FormulaRewrite>,
) {
    let dst = Range::new(
        dst_top_left,
        CellRef::new(
            dst_top_left.row + src.height() - 1,
            dst_top_left.col + src.width() - 1,
        ),
    );
    let delta_row = dst.start.row as i32 - src.start.row as i32;
    let delta_col = dst.start.col as i32 - src.start.col as i32;

    let mut extracted: Vec<(CellRef, Option<Cell>)> = Vec::new();
    for cell in src.iter() {
        extracted.push((cell, sheet.cells.get(&cell_addr_from_cell_ref(cell)).cloned()));
    }

    for cell in dst.iter() {
        sheet.cells.remove(&cell_addr_from_cell_ref(cell));
    }

    for (cell, value) in extracted {
        let Some(mut value) = value else { continue };
        let dr = cell.row - src.start.row;
        let dc = cell.col - src.start.col;
        let target = CellRef::new(dst.start.row + dr, dst.start.col + dc);

        if let Some(formula) = &value.formula {
            let (new_formula, _) =
                rewrite_formula_for_copy_delta(formula, sheet_name, delta_row, delta_col);
            if &new_formula != formula {
                formula_rewrites.push(FormulaRewrite {
                    sheet: sheet_name.to_string(),
                    cell: target,
                    before: formula.clone(),
                    after: new_formula.clone(),
                });
            }
            value.formula = Some(new_formula);
        }

        sheet.cells.insert(cell_addr_from_cell_ref(target), value);
    }
}

fn fill_range(
    sheet: &mut Sheet,
    sheet_name: &str,
    src: Range,
    dst: Range,
    formula_rewrites: &mut Vec<FormulaRewrite>,
) {
    let height = src.height() as i32;
    let width = src.width() as i32;
    if height <= 0 || width <= 0 {
        return;
    }

    for cell in dst.iter() {
        if src.contains(cell) {
            continue;
        }
        sheet.cells.remove(&cell_addr_from_cell_ref(cell));

        let rel_row = cell.row as i32 - src.start.row as i32;
        let rel_col = cell.col as i32 - src.start.col as i32;
        let src_row = src.start.row + rel_row.rem_euclid(height) as u32;
        let src_col = src.start.col + rel_col.rem_euclid(width) as u32;
        let src_cell = CellRef::new(src_row, src_col);

        let Some(mut value) = sheet.cells.get(&cell_addr_from_cell_ref(src_cell)).cloned() else {
            continue;
        };
        if let Some(formula) = &value.formula {
            let delta_row = cell.row as i32 - src_cell.row as i32;
            let delta_col = cell.col as i32 - src_cell.col as i32;
            let (new_formula, _) =
                rewrite_formula_for_copy_delta(formula, sheet_name, delta_row, delta_col);
            if &new_formula != formula {
                formula_rewrites.push(FormulaRewrite {
                    sheet: sheet_name.to_string(),
                    cell,
                    before: formula.clone(),
                    after: new_formula.clone(),
                });
            }
            value.formula = Some(new_formula);
        }
        sheet.cells.insert(cell_addr_from_cell_ref(cell), value);
    }
}

fn rewrite_all_formulas_structural(
    workbook: &mut Workbook,
    sheet_names: &HashMap<SheetId, String>,
    edit: StructuralEdit,
) -> Vec<FormulaRewrite> {
    let mut rewrites = Vec::new();
    for (sheet_id, sheet) in workbook.sheets.iter_mut().enumerate() {
        let Some(ctx_sheet) = sheet_names.get(&sheet_id) else { continue };
        for (addr, cell) in sheet.cells.iter_mut() {
            let Some(formula) = &cell.formula else { continue };
            let (new_formula, changed) =
                rewrite_formula_for_structural_edit(formula, ctx_sheet, &edit);
            if changed {
                rewrites.push(FormulaRewrite {
                    sheet: ctx_sheet.clone(),
                    cell: cell_ref_from_addr(*addr),
                    before: formula.clone(),
                    after: new_formula.clone(),
                });
                cell.formula = Some(new_formula);
            }
        }
    }
    rewrites
}

fn rewrite_all_formulas_range_map(
    workbook: &mut Workbook,
    sheet_names: &HashMap<SheetId, String>,
    edit: &RangeMapEdit,
) -> Vec<FormulaRewrite> {
    let mut rewrites = Vec::new();
    for (sheet_id, sheet) in workbook.sheets.iter_mut().enumerate() {
        let Some(ctx_sheet) = sheet_names.get(&sheet_id) else { continue };
        for (addr, cell) in sheet.cells.iter_mut() {
            let Some(formula) = &cell.formula else { continue };
            let (new_formula, changed) = rewrite_formula_for_range_map(formula, ctx_sheet, edit);
            if changed {
                rewrites.push(FormulaRewrite {
                    sheet: ctx_sheet.clone(),
                    cell: cell_ref_from_addr(*addr),
                    before: formula.clone(),
                    after: new_formula.clone(),
                });
                cell.formula = Some(new_formula);
            }
        }
    }
    rewrites
}

fn diff_workbooks(
    before: &Workbook,
    after: &Workbook,
    sheet_names: &HashMap<SheetId, String>,
) -> Vec<CellChange> {
    let mut out = Vec::new();
    let max_sheets = max(before.sheets.len(), after.sheets.len());
    for sheet_id in 0..max_sheets {
        let sheet_name = sheet_names
            .get(&sheet_id)
            .cloned()
            .unwrap_or_else(|| format!("Sheet{sheet_id}"));
        let before_sheet = before.sheets.get(sheet_id);
        let after_sheet = after.sheets.get(sheet_id);
        let mut addrs: BTreeSet<CellAddr> = BTreeSet::new();
        if let Some(sheet) = before_sheet {
            addrs.extend(sheet.cells.keys().copied());
        }
        if let Some(sheet) = after_sheet {
            addrs.extend(sheet.cells.keys().copied());
        }
        for addr in addrs {
            let before_cell = before_sheet.and_then(|s| s.cells.get(&addr));
            let after_cell = after_sheet.and_then(|s| s.cells.get(&addr));
            let before_snap = before_cell.map(cell_snapshot);
            let after_snap = after_cell.map(cell_snapshot);
            if before_snap == after_snap {
                continue;
            }
            out.push(CellChange {
                sheet: sheet_name.clone(),
                cell: cell_ref_from_addr(addr),
                before: before_snap,
                after: after_snap,
            });
        }
    }
    out
}

fn cell_snapshot(cell: &Cell) -> CellSnapshot {
    CellSnapshot {
        value: cell.value.clone(),
        formula: cell.formula.clone(),
    }
}

fn sorted_cell_keys(set: &HashSet<CellKey>) -> Vec<CellKey> {
    let mut out: Vec<CellKey> = set.iter().copied().collect();
    out.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));
    out
}

fn collect_transitive(map: &HashMap<CellKey, HashSet<CellKey>>, start: CellKey) -> Vec<CellKey> {
    let mut visited: HashSet<CellKey> = HashSet::new();
    let mut out: Vec<CellKey> = Vec::new();
    let mut queue = VecDeque::new();

    visited.insert(start);
    queue.push_back(start);

    while let Some(cell) = queue.pop_front() {
        let neighbors = map.get(&cell).map(sorted_cell_keys).unwrap_or_default();
        for n in neighbors {
            if visited.insert(n) {
                out.push(n);
                queue.push_back(n);
            }
        }
    }

    out.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));
    out
}

#[derive(Debug)]
struct Snapshot {
    sheets: HashSet<SheetId>,
    values: HashMap<CellKey, Value>,
    tables: Vec<Vec<Table>>,
}

impl Snapshot {
    fn from_workbook(workbook: &Workbook) -> Self {
        let sheets: HashSet<SheetId> = (0..workbook.sheets.len()).collect();
        let mut values = HashMap::new();
        for (sheet_id, sheet) in workbook.sheets.iter().enumerate() {
            for (addr, cell) in &sheet.cells {
                values.insert(CellKey { sheet: sheet_id, addr: *addr }, cell.value.clone());
            }
        }
        let tables = workbook.sheets.iter().map(|s| s.tables.clone()).collect();
        Self { sheets, values, tables }
    }
}

impl crate::eval::ValueResolver for Snapshot {
    fn sheet_exists(&self, sheet_id: usize) -> bool {
        self.sheets.contains(&sheet_id)
    }

    fn get_cell_value(&self, sheet_id: usize, addr: CellAddr) -> Value {
        self.values
            .get(&CellKey { sheet: sheet_id, addr })
            .cloned()
            .unwrap_or(Value::Blank)
    }

    fn resolve_structured_ref(
        &self,
        ctx: crate::eval::EvalContext,
        sref: &crate::structured_refs::StructuredRef,
    ) -> Option<(usize, CellAddr, CellAddr)> {
        crate::structured_refs::resolve_structured_ref(&self.tables, ctx.current_sheet, ctx.current_cell, sref).ok()
    }
}

fn analyze_expr(
    expr: &CompiledExpr,
    current_cell: CellKey,
    tables_by_sheet: &[Vec<Table>],
) -> (HashSet<CellKey>, bool, bool) {
    let mut precedents = HashSet::new();
    let mut volatile = false;
    let mut thread_safe = true;
    walk_expr(
        expr,
        current_cell,
        tables_by_sheet,
        &mut precedents,
        &mut volatile,
        &mut thread_safe,
    );
    (precedents, volatile, thread_safe)
}

fn walk_expr(
    expr: &CompiledExpr,
    current_cell: CellKey,
    tables_by_sheet: &[Vec<Table>],
    precedents: &mut HashSet<CellKey>,
    volatile: &mut bool,
    thread_safe: &mut bool,
) {
    match expr {
        Expr::CellRef(r) => {
            if let Some(sheet) = resolve_sheet(&r.sheet, current_cell.sheet) {
                precedents.insert(CellKey { sheet, addr: r.addr });
            }
        }
        Expr::RangeRef(RangeRef { sheet, start, end }) => {
            if let Some(sheet) = resolve_sheet(sheet, current_cell.sheet) {
                let (r1, r2) = if start.row <= end.row {
                    (start.row, end.row)
                } else {
                    (end.row, start.row)
                };
                let (c1, c2) = if start.col <= end.col {
                    (start.col, end.col)
                } else {
                    (end.col, start.col)
                };
                for row in r1..=r2 {
                    for col in c1..=c2 {
                        precedents.insert(CellKey { sheet, addr: CellAddr { row, col } });
                    }
                }
            }
        }
        Expr::StructuredRef(sref) => {
            if let Ok((sheet_id, start, end)) = crate::structured_refs::resolve_structured_ref(
                tables_by_sheet,
                current_cell.sheet,
                current_cell.addr,
                sref,
            ) {
                let (r1, r2) = if start.row <= end.row {
                    (start.row, end.row)
                } else {
                    (end.row, start.row)
                };
                let (c1, c2) = if start.col <= end.col {
                    (start.col, end.col)
                } else {
                    (end.col, start.col)
                };
                for row in r1..=r2 {
                    for col in c1..=c2 {
                        precedents.insert(CellKey {
                            sheet: sheet_id,
                            addr: CellAddr { row, col },
                        });
                    }
                }
            }
        }
        Expr::Unary { expr, .. } => walk_expr(
            expr,
            current_cell,
            tables_by_sheet,
            precedents,
            volatile,
            thread_safe,
        ),
        Expr::Binary { left, right, .. } | Expr::Compare { left, right, .. } => {
            walk_expr(
                left,
                current_cell,
                tables_by_sheet,
                precedents,
                volatile,
                thread_safe,
            );
            walk_expr(
                right,
                current_cell,
                tables_by_sheet,
                precedents,
                volatile,
                thread_safe,
            );
        }
        Expr::FunctionCall { name, args } => {
            if is_volatile_function(name) {
                *volatile = true;
            }
            // Placeholder: treat unknown/UDFs as non-thread-safe.
            if !is_known_thread_safe_function(name) {
                *thread_safe = false;
            }
            for a in args {
                walk_expr(
                    a,
                    current_cell,
                    tables_by_sheet,
                    precedents,
                    volatile,
                    thread_safe,
                );
            }
        }
        Expr::ImplicitIntersection(inner) => {
            walk_expr(
                inner,
                current_cell,
                tables_by_sheet,
                precedents,
                volatile,
                thread_safe,
            )
        }
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Blank | Expr::Error(_) => {}
    }
}

fn resolve_sheet(sheet: &SheetReference<usize>, current_sheet: SheetId) -> Option<SheetId> {
    match sheet {
        SheetReference::Current => Some(current_sheet),
        SheetReference::Sheet(id) => Some(*id),
        SheetReference::External(_) => None,
    }
}

fn is_volatile_function(name: &str) -> bool {
    matches!(
        name,
        "NOW" | "TODAY" | "RAND" | "RANDBETWEEN" | "OFFSET" | "INDIRECT" | "INFO" | "CELL"
    )
}

fn is_known_thread_safe_function(name: &str) -> bool {
    // For now we only implement a small set of built-ins, all thread-safe.
    matches!(
        name,
        "IF"
            | "IFERROR"
            | "ISERROR"
            | "SUM"
            | "PV"
            | "FV"
            | "PMT"
            | "NPER"
            | "RATE"
            | "IPMT"
            | "PPMT"
            | "SLN"
            | "SYD"
            | "DDB"
            | "NPV"
            | "IRR"
            | "XNPV"
            | "XIRR"
    )
}
