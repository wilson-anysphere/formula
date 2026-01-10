use crate::eval::{
    parse_a1, CellAddr, CompiledExpr, Expr, FormulaParseError, Parser, RangeRef, SheetReference,
};
use crate::value::{ErrorKind, Value};
use rayon::prelude::*;
use std::collections::{HashMap, HashSet, VecDeque};
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

#[derive(Debug, Default)]
struct Sheet {
    cells: HashMap<CellAddr, Cell>,
}

#[derive(Debug, Default)]
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
        self.sheets.push(Sheet { cells: HashMap::new() });
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

    fn mark_dirty_dependents(&self, from: CellKey, dirty: &mut HashSet<CellKey>) {
        let mut queue = VecDeque::new();
        if let Some(deps) = self.dependents.get(&from) {
            for &d in deps {
                queue.push_back(d);
            }
        }

        while let Some(cell) = queue.pop_front() {
            if !dirty.insert(cell) {
                continue;
            }
            if let Some(deps) = self.dependents.get(&cell) {
                for &d in deps {
                    queue.push_back(d);
                }
            }
        }
    }

    fn mark_dirty_including_self(&self, from: CellKey, dirty: &mut HashSet<CellKey>) {
        let mut queue = VecDeque::new();
        queue.push_back(from);
        while let Some(cell) = queue.pop_front() {
            if !dirty.insert(cell) {
                continue;
            }
            if let Some(deps) = self.dependents.get(&cell) {
                for &d in deps {
                    queue.push_back(d);
                }
            }
        }
    }
}

pub struct Engine {
    workbook: Workbook,
    graph: DependencyGraph,
    dirty: HashSet<CellKey>,
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

        let cell = self.workbook.get_or_create_cell_mut(key);
        cell.value = value.into();
        cell.formula = None;
        cell.ast = None;
        cell.volatile = false;
        cell.thread_safe = true;

        // Mark downstream dependents dirty.
        self.graph.mark_dirty_dependents(key, &mut self.dirty);
        Ok(())
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
        let (precedents, volatile, thread_safe) = analyze_expr(&compiled, sheet_id);

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
        self.graph
            .mark_dirty_including_self(key, &mut self.dirty);
        Ok(())
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
            self.graph.mark_dirty_including_self(cell, &mut self.dirty);
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
}

#[derive(Debug)]
struct Snapshot {
    sheets: HashSet<SheetId>,
    values: HashMap<CellKey, Value>,
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
        Self { sheets, values }
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
}

fn analyze_expr(expr: &CompiledExpr, current_sheet: SheetId) -> (HashSet<CellKey>, bool, bool) {
    let mut precedents = HashSet::new();
    let mut volatile = false;
    let mut thread_safe = true;
    walk_expr(expr, current_sheet, &mut precedents, &mut volatile, &mut thread_safe);
    (precedents, volatile, thread_safe)
}

fn walk_expr(
    expr: &CompiledExpr,
    current_sheet: SheetId,
    precedents: &mut HashSet<CellKey>,
    volatile: &mut bool,
    thread_safe: &mut bool,
) {
    match expr {
        Expr::CellRef(r) => {
            if let Some(sheet) = resolve_sheet(&r.sheet, current_sheet) {
                precedents.insert(CellKey { sheet, addr: r.addr });
            }
        }
        Expr::RangeRef(RangeRef { sheet, start, end }) => {
            if let Some(sheet) = resolve_sheet(sheet, current_sheet) {
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
        Expr::Unary { expr, .. } => walk_expr(expr, current_sheet, precedents, volatile, thread_safe),
        Expr::Binary { left, right, .. } | Expr::Compare { left, right, .. } => {
            walk_expr(left, current_sheet, precedents, volatile, thread_safe);
            walk_expr(right, current_sheet, precedents, volatile, thread_safe);
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
                walk_expr(a, current_sheet, precedents, volatile, thread_safe);
            }
        }
        Expr::ImplicitIntersection(inner) => {
            walk_expr(inner, current_sheet, precedents, volatile, thread_safe)
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
