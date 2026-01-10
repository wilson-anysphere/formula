use crate::file_io::Workbook;
use serde_json::Value as JsonValue;
use std::collections::{HashMap, HashSet, VecDeque};

#[derive(Debug, thiserror::Error)]
pub enum AppStateError {
    #[error("no workbook loaded")]
    NoWorkbookLoaded,
    #[error("no undo history")]
    NoUndoHistory,
    #[error("no redo history")]
    NoRedoHistory,
    #[error("unknown sheet id: {0}")]
    UnknownSheet(String),
    #[error("invalid range: start ({start_row},{start_col}) end ({end_row},{end_col})")]
    InvalidRange {
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    },
}

#[derive(Clone, Debug, PartialEq)]
pub enum CellScalar {
    Empty,
    Number(f64),
    Text(String),
    Bool(bool),
    Error(String),
}

impl CellScalar {
    pub fn display(&self) -> String {
        match self {
            CellScalar::Empty => String::new(),
            CellScalar::Number(n) => {
                // Avoid showing ".0" for whole numbers in the UI layer.
                if (n.fract() - 0.0).abs() < f64::EPSILON {
                    format!("{:.0}", n)
                } else {
                    n.to_string()
                }
            }
            CellScalar::Text(s) => s.clone(),
            CellScalar::Bool(true) => "TRUE".to_string(),
            CellScalar::Bool(false) => "FALSE".to_string(),
            CellScalar::Error(e) => e.clone(),
        }
    }

    pub fn as_json(&self) -> Option<JsonValue> {
        match self {
            CellScalar::Empty => None,
            CellScalar::Number(n) => Some(JsonValue::from(*n)),
            CellScalar::Text(s) => Some(JsonValue::from(s.clone())),
            CellScalar::Bool(b) => Some(JsonValue::from(*b)),
            CellScalar::Error(e) => Some(JsonValue::from(e.clone())),
        }
    }

    pub fn from_json(value: &JsonValue) -> Self {
        match value {
            JsonValue::Null => CellScalar::Empty,
            JsonValue::Bool(b) => CellScalar::Bool(*b),
            JsonValue::Number(n) => n
                .as_f64()
                .map(CellScalar::Number)
                .unwrap_or_else(|| CellScalar::Error("#NUM!".to_string())),
            JsonValue::String(s) => CellScalar::Text(s.clone()),
            other => CellScalar::Text(other.to_string()),
        }
    }

    fn coerce_number(&self) -> Result<f64, CellScalar> {
        match self {
            CellScalar::Empty => Ok(0.0),
            CellScalar::Number(n) => Ok(*n),
            CellScalar::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
            CellScalar::Text(s) => s
                .trim()
                .parse::<f64>()
                .map_err(|_| CellScalar::Error("#VALUE!".to_string())),
            CellScalar::Error(e) => Err(CellScalar::Error(e.clone())),
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct Cell {
    pub(crate) input_value: Option<CellScalar>,
    pub(crate) formula: Option<String>,
    pub(crate) computed_value: CellScalar,
}

impl Cell {
    pub(crate) fn empty() -> Self {
        Self {
            input_value: None,
            formula: None,
            computed_value: CellScalar::Empty,
        }
    }

    pub(crate) fn from_literal(value: Option<CellScalar>) -> Self {
        let computed_value = value.clone().unwrap_or(CellScalar::Empty);
        Self {
            input_value: value,
            formula: None,
            computed_value,
        }
    }

    pub(crate) fn from_formula(formula: String) -> Self {
        Self {
            input_value: None,
            formula: Some(formula),
            computed_value: CellScalar::Empty,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct CellData {
    pub value: CellScalar,
    pub formula: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct CellUpdateData {
    pub sheet_id: String,
    pub row: usize,
    pub col: usize,
    pub value: CellScalar,
    pub formula: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct SheetInfoData {
    pub id: String,
    pub name: String,
}

#[derive(Clone, Debug, PartialEq)]
pub struct WorkbookInfoData {
    pub path: Option<String>,
    pub sheets: Vec<SheetInfoData>,
}

#[derive(Clone, Debug, PartialEq)]
struct CellInputSnapshot {
    sheet_id: String,
    row: usize,
    col: usize,
    value: Option<CellScalar>,
    formula: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
struct UndoEntry {
    before: Vec<CellInputSnapshot>,
    after: Vec<CellInputSnapshot>,
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
struct CellKey {
    sheet_id: String,
    row: usize,
    col: usize,
}

#[derive(Default, Debug)]
struct DependencyGraph {
    deps: HashMap<CellKey, HashSet<CellKey>>,
    rev_deps: HashMap<CellKey, HashSet<CellKey>>,
}

impl DependencyGraph {
    fn clear(&mut self) {
        self.deps.clear();
        self.rev_deps.clear();
    }

    fn set_deps(&mut self, cell: &CellKey, new_deps: HashSet<CellKey>) {
        let old = self.deps.insert(cell.clone(), new_deps.clone());
        if let Some(old_deps) = old {
            for dep in old_deps {
                if let Some(dependents) = self.rev_deps.get_mut(&dep) {
                    dependents.remove(cell);
                    if dependents.is_empty() {
                        self.rev_deps.remove(&dep);
                    }
                }
            }
        }

        for dep in new_deps {
            self.rev_deps.entry(dep).or_default().insert(cell.clone());
        }
    }

    fn remove_cell(&mut self, cell: &CellKey) {
        if let Some(old_deps) = self.deps.remove(cell) {
            for dep in old_deps {
                if let Some(dependents) = self.rev_deps.get_mut(&dep) {
                    dependents.remove(cell);
                    if dependents.is_empty() {
                        self.rev_deps.remove(&dep);
                    }
                }
            }
        }
    }

    fn dependents_closure(&self, start: &CellKey) -> HashSet<CellKey> {
        let mut out = HashSet::new();
        let mut queue = VecDeque::new();
        queue.push_back(start.clone());

        while let Some(cell) = queue.pop_front() {
            if !out.insert(cell.clone()) {
                continue;
            }

            if let Some(dependents) = self.rev_deps.get(&cell) {
                for dependent in dependents {
                    queue.push_back(dependent.clone());
                }
            }
        }

        out
    }
}

#[derive(Default, Debug)]
struct CalcEngine {
    graph: DependencyGraph,
}

impl CalcEngine {
    fn rebuild(&mut self, workbook: &Workbook) {
        self.graph.clear();
        for sheet in &workbook.sheets {
            for ((row, col), cell) in sheet.cells_iter() {
                if let Some(formula) = &cell.formula {
                    let key = CellKey {
                        sheet_id: sheet.id.clone(),
                        row,
                        col,
                    };
                    let deps = extract_dependencies(formula, &sheet.id);
                    self.graph.set_deps(&key, deps);
                }
            }
        }
    }

    fn update_cell_formula(
        &mut self,
        sheet_id: &str,
        row: usize,
        col: usize,
        formula: Option<&str>,
    ) {
        let key = CellKey {
            sheet_id: sheet_id.to_string(),
            row,
            col,
        };

        if let Some(formula) = formula {
            let deps = extract_dependencies(formula, sheet_id);
            self.graph.set_deps(&key, deps);
        } else {
            self.graph.remove_cell(&key);
        }
    }
}

pub struct AppState {
    workbook: Option<Workbook>,
    engine: CalcEngine,
    dirty: bool,
    undo_stack: Vec<UndoEntry>,
    redo_stack: Vec<UndoEntry>,
}

impl Default for AppState {
    fn default() -> Self {
        Self::new()
    }
}

pub type SharedAppState = std::sync::Arc<std::sync::Mutex<AppState>>;

impl AppState {
    pub fn new() -> Self {
        Self {
            workbook: None,
            engine: CalcEngine::default(),
            dirty: false,
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
        }
    }

    pub fn has_unsaved_changes(&self) -> bool {
        self.dirty
    }

    pub fn workbook_info(&self) -> Result<WorkbookInfoData, AppStateError> {
        let workbook = self
            .workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        Ok(WorkbookInfoData {
            path: workbook.path.clone(),
            sheets: workbook
                .sheets
                .iter()
                .map(|sheet| SheetInfoData {
                    id: sheet.id.clone(),
                    name: sheet.name.clone(),
                })
                .collect(),
        })
    }

    pub fn load_workbook(&mut self, mut workbook: Workbook) -> WorkbookInfoData {
        workbook.ensure_sheet_ids();
        self.engine.rebuild(&workbook);

        // Calculate formulas once on load so the UI sees fresh values even if the file
        // only has cached values.
        let _ = recalculate_all_formulas(&mut workbook, &self.engine.graph);

        self.workbook = Some(workbook);
        self.dirty = false;
        self.undo_stack.clear();
        self.redo_stack.clear();
        self.workbook_info()
            .expect("workbook_info should succeed right after load")
    }

    pub fn mark_saved(&mut self, new_path: Option<String>) -> Result<(), AppStateError> {
        let workbook = self
            .workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        if let Some(path) = new_path {
            workbook.path = Some(path);
        }
        self.dirty = false;
        Ok(())
    }

    pub fn get_workbook(&self) -> Result<&Workbook, AppStateError> {
        self.workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)
    }

    pub fn get_workbook_mut(&mut self) -> Result<&mut Workbook, AppStateError> {
        self.workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)
    }

    pub fn get_cell(
        &self,
        sheet_id: &str,
        row: usize,
        col: usize,
    ) -> Result<CellData, AppStateError> {
        let workbook = self
            .workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let sheet = workbook
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
        let cell = sheet.get_cell(row, col);
        Ok(CellData {
            value: cell.computed_value.clone(),
            formula: cell.formula.clone(),
        })
    }

    pub fn get_range(
        &self,
        sheet_id: &str,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> Result<Vec<Vec<CellData>>, AppStateError> {
        if start_row > end_row || start_col > end_col {
            return Err(AppStateError::InvalidRange {
                start_row,
                start_col,
                end_row,
                end_col,
            });
        }

        let workbook = self
            .workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let sheet = workbook
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;

        let mut rows = Vec::new();
        for row in start_row..=end_row {
            let mut cols = Vec::new();
            for col in start_col..=end_col {
                let cell = sheet.get_cell(row, col);
                cols.push(CellData {
                    value: cell.computed_value.clone(),
                    formula: cell.formula.clone(),
                });
            }
            rows.push(cols);
        }

        Ok(rows)
    }

    pub fn set_cell(
        &mut self,
        sheet_id: &str,
        row: usize,
        col: usize,
        value: Option<JsonValue>,
        formula: Option<String>,
    ) -> Result<Vec<CellUpdateData>, AppStateError> {
        let before = self.snapshot_cell(sheet_id, row, col)?;
        let after_cell = CellInputSnapshot {
            sheet_id: sheet_id.to_string(),
            row,
            col,
            value: value.as_ref().map(CellScalar::from_json),
            formula: normalize_formula(formula),
        };

        if before == after_cell {
            return Ok(Vec::new());
        }

        self.apply_snapshots(&[after_cell.clone()])?;
        let mut updates = self.recalculate_from_cells(&[CellKey {
            sheet_id: sheet_id.to_string(),
            row,
            col,
        }]);

        // Ensure the edited cell is included even if its computed value didn't change
        // (e.g. formula edit).
        if !updates
            .iter()
            .any(|u| u.sheet_id == sheet_id && u.row == row && u.col == col)
        {
            let cell = self.get_cell(sheet_id, row, col)?;
            updates.push(CellUpdateData {
                sheet_id: sheet_id.to_string(),
                row,
                col,
                value: cell.value,
                formula: cell.formula,
            });
        }

        self.dirty = true;
        self.redo_stack.clear();
        self.undo_stack.push(UndoEntry {
            before: vec![before],
            after: vec![after_cell],
        });

        Ok(updates)
    }

    pub fn set_range(
        &mut self,
        sheet_id: &str,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        values: Vec<Vec<(Option<JsonValue>, Option<String>)>>,
    ) -> Result<Vec<CellUpdateData>, AppStateError> {
        if start_row > end_row || start_col > end_col {
            return Err(AppStateError::InvalidRange {
                start_row,
                start_col,
                end_row,
                end_col,
            });
        }

        let expected_rows = end_row - start_row + 1;
        let expected_cols = end_col - start_col + 1;
        if values.len() != expected_rows || values.iter().any(|row| row.len() != expected_cols) {
            return Err(AppStateError::InvalidRange {
                start_row,
                start_col,
                end_row,
                end_col,
            });
        }

        let mut before = Vec::new();
        let mut after = Vec::new();
        let mut changed = Vec::new();

        for (r_off, row_values) in values.iter().enumerate() {
            for (c_off, (value, formula)) in row_values.iter().enumerate() {
                let row = start_row + r_off;
                let col = start_col + c_off;

                let snapshot_before = self.snapshot_cell(sheet_id, row, col)?;
                let snapshot_after = CellInputSnapshot {
                    sheet_id: sheet_id.to_string(),
                    row,
                    col,
                    value: value.as_ref().map(CellScalar::from_json),
                    formula: normalize_formula(formula.clone()),
                };

                if snapshot_before != snapshot_after {
                    before.push(snapshot_before);
                    after.push(snapshot_after.clone());
                    changed.push(CellKey {
                        sheet_id: sheet_id.to_string(),
                        row,
                        col,
                    });
                }
            }
        }

        if changed.is_empty() {
            return Ok(Vec::new());
        }

        self.apply_snapshots(&after)?;
        let mut updates = self.recalculate_from_cells(&changed);

        // Ensure all explicitly edited cells are included in the update list.
        for cell in changed {
            if !updates
                .iter()
                .any(|u| u.sheet_id == cell.sheet_id && u.row == cell.row && u.col == cell.col)
            {
                let cell_data = self.get_cell(&cell.sheet_id, cell.row, cell.col)?;
                updates.push(CellUpdateData {
                    sheet_id: cell.sheet_id,
                    row: cell.row,
                    col: cell.col,
                    value: cell_data.value,
                    formula: cell_data.formula,
                });
            }
        }

        self.dirty = true;
        self.redo_stack.clear();
        self.undo_stack.push(UndoEntry { before, after });
        Ok(updates)
    }

    pub fn recalculate_all(&mut self) -> Result<Vec<CellUpdateData>, AppStateError> {
        let workbook = self
            .workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let updates = recalculate_all_formulas(workbook, &self.engine.graph);
        Ok(updates)
    }

    pub fn undo(&mut self) -> Result<Vec<CellUpdateData>, AppStateError> {
        if self.workbook.is_none() {
            return Err(AppStateError::NoWorkbookLoaded);
        }
        let entry = self.undo_stack.pop().ok_or(AppStateError::NoUndoHistory)?;
        self.apply_snapshots(&entry.before)?;
        let changed: Vec<CellKey> = entry
            .before
            .iter()
            .map(|s| CellKey {
                sheet_id: s.sheet_id.clone(),
                row: s.row,
                col: s.col,
            })
            .collect();
        let mut updates = self.recalculate_from_cells(&changed);
        for cell in &changed {
            if !updates
                .iter()
                .any(|u| u.sheet_id == cell.sheet_id && u.row == cell.row && u.col == cell.col)
            {
                let cell_data = self.get_cell(&cell.sheet_id, cell.row, cell.col)?;
                updates.push(CellUpdateData {
                    sheet_id: cell.sheet_id.clone(),
                    row: cell.row,
                    col: cell.col,
                    value: cell_data.value,
                    formula: cell_data.formula,
                });
            }
        }
        self.redo_stack.push(entry);
        self.dirty = true;
        Ok(updates)
    }

    pub fn redo(&mut self) -> Result<Vec<CellUpdateData>, AppStateError> {
        if self.workbook.is_none() {
            return Err(AppStateError::NoWorkbookLoaded);
        }
        let entry = self.redo_stack.pop().ok_or(AppStateError::NoRedoHistory)?;
        self.apply_snapshots(&entry.after)?;
        let changed: Vec<CellKey> = entry
            .after
            .iter()
            .map(|s| CellKey {
                sheet_id: s.sheet_id.clone(),
                row: s.row,
                col: s.col,
            })
            .collect();
        let mut updates = self.recalculate_from_cells(&changed);
        for cell in &changed {
            if !updates
                .iter()
                .any(|u| u.sheet_id == cell.sheet_id && u.row == cell.row && u.col == cell.col)
            {
                let cell_data = self.get_cell(&cell.sheet_id, cell.row, cell.col)?;
                updates.push(CellUpdateData {
                    sheet_id: cell.sheet_id.clone(),
                    row: cell.row,
                    col: cell.col,
                    value: cell_data.value,
                    formula: cell_data.formula,
                });
            }
        }
        self.undo_stack.push(entry);
        self.dirty = true;
        Ok(updates)
    }

    fn snapshot_cell(
        &self,
        sheet_id: &str,
        row: usize,
        col: usize,
    ) -> Result<CellInputSnapshot, AppStateError> {
        let workbook = self
            .workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let sheet = workbook
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
        let cell = sheet.get_cell(row, col);

        Ok(CellInputSnapshot {
            sheet_id: sheet_id.to_string(),
            row,
            col,
            value: cell.input_value.clone(),
            formula: cell.formula.clone(),
        })
    }

    fn apply_snapshots(&mut self, snapshots: &[CellInputSnapshot]) -> Result<(), AppStateError> {
        let workbook = self
            .workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        for snap in snapshots {
            let sheet = workbook
                .sheet_mut(&snap.sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(snap.sheet_id.clone()))?;

            let new_cell = match (&snap.formula, &snap.value) {
                (Some(formula), _) => Cell::from_formula(formula.clone()),
                (None, Some(value)) => Cell::from_literal(Some(value.clone())),
                (None, None) => Cell::empty(),
            };

            sheet.set_cell(snap.row, snap.col, new_cell);
            self.engine.update_cell_formula(
                &snap.sheet_id,
                snap.row,
                snap.col,
                snap.formula.as_deref(),
            );
        }

        Ok(())
    }

    fn recalculate_from_cells(&mut self, changed: &[CellKey]) -> Vec<CellUpdateData> {
        let workbook = match self.workbook.as_mut() {
            Some(workbook) => workbook,
            None => return Vec::new(),
        };

        let mut impacted = HashSet::new();
        for cell in changed {
            impacted.extend(self.engine.graph.dependents_closure(cell));
        }

        recalculate_impacted(workbook, &self.engine.graph, &impacted)
    }
}

fn normalize_formula(formula: Option<String>) -> Option<String> {
    let mut formula = formula?.trim().to_string();
    if formula.is_empty() {
        return None;
    }
    if !formula.starts_with('=') {
        formula.insert(0, '=');
    }
    Some(formula)
}

fn recalculate_all_formulas(
    workbook: &mut Workbook,
    graph: &DependencyGraph,
) -> Vec<CellUpdateData> {
    let formula_cells: HashSet<CellKey> = workbook
        .sheets
        .iter()
        .flat_map(|sheet| {
            sheet
                .cells_iter()
                .filter(|(_, cell)| cell.formula.is_some())
                .map(|((row, col), _)| CellKey {
                    sheet_id: sheet.id.clone(),
                    row,
                    col,
                })
                .collect::<Vec<_>>()
        })
        .collect();

    recalculate_impacted(workbook, graph, &formula_cells)
}

fn recalculate_impacted(
    workbook: &mut Workbook,
    graph: &DependencyGraph,
    impacted: &HashSet<CellKey>,
) -> Vec<CellUpdateData> {
    if impacted.is_empty() {
        return Vec::new();
    }

    let formula_set: HashSet<CellKey> = impacted
        .iter()
        .filter(|key| workbook.cell_has_formula(&key.sheet_id, key.row, key.col))
        .cloned()
        .collect();

    let mut in_degree: HashMap<CellKey, usize> = formula_set
        .iter()
        .map(|key| (key.clone(), 0usize))
        .collect();

    for cell in &formula_set {
        if let Some(deps) = graph.deps.get(cell) {
            let mut count = 0usize;
            for dep in deps {
                if formula_set.contains(dep) {
                    count += 1;
                }
            }
            in_degree.insert(cell.clone(), count);
        }
    }

    let mut queue: VecDeque<CellKey> = in_degree
        .iter()
        .filter(|(_, &deg)| deg == 0)
        .map(|(cell, _)| cell.clone())
        .collect();

    let mut ordered = Vec::new();
    while let Some(cell) = queue.pop_front() {
        ordered.push(cell.clone());
        if let Some(dependents) = graph.rev_deps.get(&cell) {
            for dependent in dependents {
                if !formula_set.contains(dependent) {
                    continue;
                }
                if let Some(deg) = in_degree.get_mut(dependent) {
                    *deg = deg.saturating_sub(1);
                    if *deg == 0 {
                        queue.push_back(dependent.clone());
                    }
                }
            }
        }
    }

    let mut updates = Vec::new();

    // Any remaining nodes have a cycle. Mark them as errors.
    if ordered.len() != formula_set.len() {
        let ordered_set: HashSet<CellKey> = ordered.iter().cloned().collect();
        for cell in formula_set.difference(&ordered_set) {
            let old = workbook.cell_value(&cell.sheet_id, cell.row, cell.col);
            if old != CellScalar::Error("#CYCLE!".to_string()) {
                workbook.set_computed_value(
                    &cell.sheet_id,
                    cell.row,
                    cell.col,
                    CellScalar::Error("#CYCLE!".to_string()),
                );
                updates.push(CellUpdateData {
                    sheet_id: cell.sheet_id.clone(),
                    row: cell.row,
                    col: cell.col,
                    value: CellScalar::Error("#CYCLE!".to_string()),
                    formula: workbook.cell_formula(&cell.sheet_id, cell.row, cell.col),
                });
            }
        }
    }

    for cell in ordered {
        let old = workbook.cell_value(&cell.sheet_id, cell.row, cell.col);
        let formula = workbook.cell_formula(&cell.sheet_id, cell.row, cell.col);
        let new = match formula.as_deref() {
            Some(formula) => evaluate_formula(formula, &cell.sheet_id, workbook),
            None => old.clone(),
        };

        if new != old {
            workbook.set_computed_value(&cell.sheet_id, cell.row, cell.col, new.clone());
            updates.push(CellUpdateData {
                sheet_id: cell.sheet_id.clone(),
                row: cell.row,
                col: cell.col,
                value: new,
                formula,
            });
        }
    }

    updates
}

#[derive(Clone, Debug, PartialEq)]
enum Expr {
    Number(f64),
    CellRef {
        row: usize,
        col: usize,
    },
    Unary {
        op: char,
        expr: Box<Expr>,
    },
    Binary {
        op: char,
        left: Box<Expr>,
        right: Box<Expr>,
    },
}

struct Parser {
    chars: Vec<char>,
    pos: usize,
}

impl Parser {
    fn new(input: &str) -> Self {
        Self {
            chars: input.chars().collect(),
            pos: 0,
        }
    }

    fn parse(mut self) -> Result<Expr, CellScalar> {
        let expr = self.parse_expr()?;
        self.skip_ws();
        if self.pos < self.chars.len() {
            return Err(CellScalar::Error("#PARSE!".to_string()));
        }
        Ok(expr)
    }

    fn parse_expr(&mut self) -> Result<Expr, CellScalar> {
        let mut node = self.parse_term()?;
        loop {
            self.skip_ws();
            let op = match self.peek() {
                Some('+') => '+',
                Some('-') => '-',
                _ => break,
            };
            self.pos += 1;
            let rhs = self.parse_term()?;
            node = Expr::Binary {
                op,
                left: Box::new(node),
                right: Box::new(rhs),
            };
        }
        Ok(node)
    }

    fn parse_term(&mut self) -> Result<Expr, CellScalar> {
        let mut node = self.parse_unary()?;
        loop {
            self.skip_ws();
            let op = match self.peek() {
                Some('*') => '*',
                Some('/') => '/',
                _ => break,
            };
            self.pos += 1;
            let rhs = self.parse_unary()?;
            node = Expr::Binary {
                op,
                left: Box::new(node),
                right: Box::new(rhs),
            };
        }
        Ok(node)
    }

    fn parse_unary(&mut self) -> Result<Expr, CellScalar> {
        self.skip_ws();
        match self.peek() {
            Some('+') | Some('-') => {
                let op = self.peek().unwrap();
                self.pos += 1;
                let expr = self.parse_unary()?;
                Ok(Expr::Unary {
                    op,
                    expr: Box::new(expr),
                })
            }
            _ => self.parse_primary(),
        }
    }

    fn parse_primary(&mut self) -> Result<Expr, CellScalar> {
        self.skip_ws();
        match self.peek() {
            Some('(') => {
                self.pos += 1;
                let expr = self.parse_expr()?;
                self.skip_ws();
                if self.peek() != Some(')') {
                    return Err(CellScalar::Error("#PARSE!".to_string()));
                }
                self.pos += 1;
                Ok(expr)
            }
            Some(c) if c.is_ascii_digit() || c == '.' => self.parse_number(),
            Some(_) => self.parse_cell_ref(),
            None => Err(CellScalar::Error("#PARSE!".to_string())),
        }
    }

    fn parse_number(&mut self) -> Result<Expr, CellScalar> {
        self.skip_ws();
        let start = self.pos;
        let mut saw_digit = false;
        while let Some(c) = self.peek() {
            if c.is_ascii_digit() {
                saw_digit = true;
                self.pos += 1;
                continue;
            }
            if c == '.' {
                self.pos += 1;
                continue;
            }
            break;
        }
        if !saw_digit {
            return Err(CellScalar::Error("#PARSE!".to_string()));
        }
        let s: String = self.chars[start..self.pos].iter().collect();
        let n = s
            .parse::<f64>()
            .map_err(|_| CellScalar::Error("#NUM!".to_string()))?;
        Ok(Expr::Number(n))
    }

    fn parse_cell_ref(&mut self) -> Result<Expr, CellScalar> {
        self.skip_ws();
        let start = self.pos;
        if self.peek() == Some('$') {
            self.pos += 1;
        }

        let col_start = self.pos;
        while let Some(c) = self.peek() {
            if c.is_ascii_alphabetic() {
                self.pos += 1;
                continue;
            }
            break;
        }

        if col_start == self.pos {
            return Err(CellScalar::Error("#PARSE!".to_string()));
        }

        if self.peek() == Some('$') {
            self.pos += 1;
        }

        let row_start = self.pos;
        while let Some(c) = self.peek() {
            if c.is_ascii_digit() {
                self.pos += 1;
                continue;
            }
            break;
        }

        if row_start == self.pos {
            // Not a cell ref, treat as unsupported token (e.g. function name)
            let token: String = self.chars[start..self.pos].iter().collect();
            return Err(CellScalar::Error(format!("#UNSUPPORTED({})!", token)));
        }

        let col_str: String = self.chars[col_start..row_start]
            .iter()
            .filter(|c| c.is_ascii_alphabetic())
            .collect();
        let row_str: String = self.chars[row_start..self.pos].iter().collect();
        let col =
            col_letters_to_index(&col_str).ok_or_else(|| CellScalar::Error("#REF!".to_string()))?;
        let row_num = row_str
            .parse::<usize>()
            .map_err(|_| CellScalar::Error("#REF!".to_string()))?;
        if row_num == 0 {
            return Err(CellScalar::Error("#REF!".to_string()));
        }
        Ok(Expr::CellRef {
            row: row_num - 1,
            col,
        })
    }

    fn skip_ws(&mut self) {
        while let Some(c) = self.peek() {
            if c.is_whitespace() {
                self.pos += 1;
            } else {
                break;
            }
        }
    }

    fn peek(&self) -> Option<char> {
        self.chars.get(self.pos).copied()
    }
}

fn evaluate_formula(formula: &str, sheet_id: &str, workbook: &Workbook) -> CellScalar {
    let formula = formula.trim();
    let body = formula.strip_prefix('=').unwrap_or(formula);

    let expr = match Parser::new(body).parse() {
        Ok(expr) => expr,
        Err(err) => return err,
    };

    eval_expr(&expr, sheet_id, workbook)
}

fn eval_expr(expr: &Expr, sheet_id: &str, workbook: &Workbook) -> CellScalar {
    match expr {
        Expr::Number(n) => CellScalar::Number(*n),
        Expr::CellRef { row, col } => workbook.cell_value(sheet_id, *row, *col),
        Expr::Unary { op, expr } => {
            let value = eval_expr(expr, sheet_id, workbook);
            let n = match value.coerce_number() {
                Ok(n) => n,
                Err(err) => return err,
            };
            match op {
                '+' => CellScalar::Number(n),
                '-' => CellScalar::Number(-n),
                _ => CellScalar::Error("#PARSE!".to_string()),
            }
        }
        Expr::Binary { op, left, right } => {
            let left_val = eval_expr(left, sheet_id, workbook);
            let right_val = eval_expr(right, sheet_id, workbook);

            let l = match left_val.coerce_number() {
                Ok(n) => n,
                Err(err) => return err,
            };
            let r = match right_val.coerce_number() {
                Ok(n) => n,
                Err(err) => return err,
            };

            match op {
                '+' => CellScalar::Number(l + r),
                '-' => CellScalar::Number(l - r),
                '*' => CellScalar::Number(l * r),
                '/' => {
                    if r == 0.0 {
                        CellScalar::Error("#DIV/0!".to_string())
                    } else {
                        CellScalar::Number(l / r)
                    }
                }
                _ => CellScalar::Error("#PARSE!".to_string()),
            }
        }
    }
}

fn extract_dependencies(formula: &str, sheet_id: &str) -> HashSet<CellKey> {
    let formula = formula.trim();
    let body = formula.strip_prefix('=').unwrap_or(formula);
    let mut deps = HashSet::new();

    let chars: Vec<char> = body.chars().collect();
    let mut i = 0usize;
    while i < chars.len() {
        if chars[i] == '$' || chars[i].is_ascii_alphabetic() {
            let mut j = i;
            if chars[j] == '$' {
                j += 1;
            }
            let col_start = j;
            while j < chars.len() && chars[j].is_ascii_alphabetic() {
                j += 1;
            }
            let col_end = j;

            if col_end == col_start {
                i += 1;
                continue;
            }

            // Avoid false positives like "Sheet1" or function names.
            if col_end - col_start > 3 {
                i += 1;
                continue;
            }

            if j < chars.len() && chars[j] == '$' {
                j += 1;
            }
            let row_start = j;
            while j < chars.len() && chars[j].is_ascii_digit() {
                j += 1;
            }
            if row_start == j {
                i += 1;
                continue;
            }

            let col: String = chars[col_start..col_end].iter().collect();
            let row: String = chars[row_start..j].iter().collect();
            if let (Some(col_idx), Ok(row_idx)) = (col_letters_to_index(&col), row.parse::<usize>())
            {
                if row_idx > 0 {
                    deps.insert(CellKey {
                        sheet_id: sheet_id.to_string(),
                        row: row_idx - 1,
                        col: col_idx,
                    });
                }
            }

            i = j;
        } else {
            i += 1;
        }
    }

    deps
}

fn col_letters_to_index(input: &str) -> Option<usize> {
    if input.is_empty() {
        return None;
    }
    let mut col = 0usize;
    for ch in input.chars() {
        let ch = ch.to_ascii_uppercase();
        if !('A'..='Z').contains(&ch) {
            return None;
        }
        col = col * 26 + (ch as usize - 'A' as usize + 1);
    }
    Some(col - 1)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::file_io::{read_xlsx_blocking, write_xlsx_blocking};

    #[test]
    fn set_cell_recalculates_dependents() {
        let mut workbook = Workbook::new_empty(Some("fixture.xlsx".to_string()));
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        // A1 = 1, B1 = =A1+1
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(1.0))),
        );
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_formula("=A1+1".to_string()),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let b1 = state.get_cell(&sheet_id, 0, 1).unwrap();
        assert_eq!(b1.value, CellScalar::Number(2.0));

        let updates = state
            .set_cell(&sheet_id, 0, 0, Some(JsonValue::from(10)), None)
            .unwrap();
        assert!(updates.iter().any(|u| u.row == 0 && u.col == 1));

        let b1_after = state.get_cell(&sheet_id, 0, 1).unwrap();
        assert_eq!(b1_after.value, CellScalar::Number(11.0));
        assert!(state.has_unsaved_changes());
    }

    #[test]
    fn undo_redo_round_trip() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(1.0))),
        );
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_formula("=A1+1".to_string()),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .set_cell(&sheet_id, 0, 0, Some(JsonValue::from(10)), None)
            .unwrap();
        assert_eq!(
            state.get_cell(&sheet_id, 0, 1).unwrap().value,
            CellScalar::Number(11.0)
        );

        state.undo().unwrap();
        assert_eq!(
            state.get_cell(&sheet_id, 0, 1).unwrap().value,
            CellScalar::Number(2.0)
        );

        state.redo().unwrap();
        assert_eq!(
            state.get_cell(&sheet_id, 0, 1).unwrap().value,
            CellScalar::Number(11.0)
        );
    }

    #[test]
    fn xlsx_round_trip_through_file_io() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(3.0))),
        );
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_formula("=A1+4".to_string()),
        );

        let tmp_dir = tempfile::tempdir().expect("temp dir");
        let path = tmp_dir.path().join("roundtrip.xlsx");
        write_xlsx_blocking(&path, &workbook).expect("write");

        let loaded = read_xlsx_blocking(&path).expect("read");
        let mut state = AppState::new();
        let info = state.load_workbook(loaded);
        assert_eq!(info.sheets.len(), 1);

        let sheet_id = info.sheets[0].id.clone();
        let b1 = state.get_cell(&sheet_id, 0, 1).unwrap();
        assert_eq!(b1.value, CellScalar::Number(7.0));
    }
}
